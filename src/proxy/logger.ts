import { promises as fs } from "node:fs";
import path from "node:path";
import type { JsonObject, LogLevel, WrapperOptions } from "./types";
import { isJsonObject, tryGetInt, tryGetString, tryParseJsonObject, tryPrettyJson } from "./utils";

const LogLevelPriority: Record<LogLevel, number> = {
  error: 0,
  warning: 1,
  info: 2,
  debug: 3,
  trace: 4,
};

type FrameAnalysis = {
  index: number;
  raw: string;
  json: JsonObject | null;
  type: number | null;
  target: string | null;
  requestId: string | null;
  messageId: string | null;
  messageText: string | null;
  writeAtCursor: string | null;
  resultValue: string | null;
  resultMessage: string | null;
  error: string | null;
  hasCursor: boolean;
  isTerminal: boolean;
  isEmptyObject: boolean;
  hasCompleteMarkdownJson: boolean;
};

type ResponseProgress = {
  sawFirstText: boolean;
  sawFirstDelta: boolean;
  sawCompleteJson: boolean;
  maxTextLength: number;
};

const CompleteMarkdownJsonPattern = /```json\s*\{[\s\S]*\}\s*```/i;

export class DebugMarkdownLogger {
  private sequence = 0;
  private readonly logLevel: LogLevel;
  private readonly responseProgressByKey = new Map<string, ResponseProgress>();

  constructor(
    private readonly options: WrapperOptions,
    private readonly isEnabled: boolean,
  ) {
    this.logLevel = normalizeLogLevel(options.logLevel);
  }

  async logIncomingRequest(
    request: Request,
    rawBody: string | null,
  ): Promise<void> {
    await this.logHttpLike(
      "Incoming Request",
      [
        ["Method", request.method],
        ["Uri", request.url],
      ],
      [...request.headers.entries()],
      rawBody,
      "incoming-request",
    );
  }

  async logOutgoingResponse(
    statusCode: number,
    headers: Iterable<[string, string]>,
    rawBody: string | null,
  ): Promise<void> {
    await this.logHttpLike(
      "Outgoing Response",
      [["Status", String(statusCode)]],
      [...headers],
      rawBody,
      "outgoing-response",
      statusCode,
    );
  }

  async logUpstreamRequest(
    method: string,
    uri: string,
    headers: Iterable<[string, string]>,
    body: string | null,
  ): Promise<void> {
    await this.logHttpLike(
      "Upstream Request",
      [
        ["Method", method],
        ["Uri", uri],
      ],
      [...headers],
      body,
      "request",
    );
  }

  async logUpstreamResponse(
    statusCode: number,
    uri: string,
    headers: Iterable<[string, string]>,
    body: string | null,
    includeBody: boolean,
  ): Promise<void> {
    await this.logHttpLike(
      "Upstream Response",
      [
        ["Status", String(statusCode)],
        ["Uri", uri],
      ],
      [...headers],
      includeBody ? body : null,
      includeBody ? "response" : "response-headers",
    );
  }

  async logSubstrateFrame(
    uri: string,
    direction: string,
    payload: string,
  ): Promise<void> {
    const normalizedDirection = direction.trim().toLowerCase() || "frame";
    if (
      normalizedDirection === "request" &&
      this.logLevel === "info" &&
      isTrivialSubstrateRequestPayload(payload)
    ) {
      return;
    }

    const structuredBody = this.buildStructuredSubstrateBody(
      normalizedDirection,
      payload,
    );
    if (!structuredBody) {
      return;
    }

    await this.logHttpLike(
      `Substrate WebSocket ${normalizedDirection}`,
      [
        ["Uri", redactUriTokens(uri)],
        ["Direction", normalizedDirection],
      ],
      [],
      structuredBody,
      `substrate-${normalizedDirection}`,
    );
  }

  async logSimulatedStreamingDiagnostics(details: JsonObject): Promise<void> {
    if (
      !this.isEnabled ||
      !this.options.debugPath ||
      !this.options.debugPath.trim() ||
      !this.isLevelEnabled("trace")
    ) {
      return;
    }

    await this.logHttpLike(
      "Simulated Streaming Diagnostics",
      [],
      [],
      JSON.stringify(details),
      "simulated-streaming",
    );
  }

  private buildStructuredSubstrateBody(
    direction: string,
    payload: string,
  ): string | null {
    const frameTexts = splitSignalRFrames(payload);
    const analyses = frameTexts.map((raw, index) => analyzeFrame(raw, index));

    if (direction !== "response" || this.logLevel === "trace") {
      const frames = analyses.map((analysis) =>
        buildFrameLogEntry(analysis, "all", true),
      );
      return JSON.stringify({
        format: "signalr-json-v1",
        direction,
        frameCount: analyses.length,
        includedFrameCount: frames.length,
        omittedFrameCount: 0,
        frames,
      });
    }

    const selected: JsonObject[] = [];
    for (const analysis of analyses) {
      const reason = this.selectResponseFrameReason(analysis);
      if (!reason) {
        continue;
      }
      selected.push(buildFrameLogEntry(analysis, reason, false));
      if (analysis.isTerminal) {
        this.clearFrameProgress(analysis);
      }
    }

    if (selected.length === 0) {
      return null;
    }

    return JSON.stringify({
      format: "signalr-json-v1",
      direction,
      frameCount: analyses.length,
      includedFrameCount: selected.length,
      omittedFrameCount: analyses.length - selected.length,
      frames: selected,
    });
  }

  private selectResponseFrameReason(analysis: FrameAnalysis): string | null {
    if (!analysis.json) {
      return "non_json_frame";
    }

    if (analysis.error) {
      return "error";
    }

    if (analysis.resultValue && !isSubstrateResultSuccess(analysis.resultValue)) {
      return "result_error";
    }

    if (analysis.isTerminal) {
      return "terminal";
    }

    const progress = this.getFrameProgress(analysis);

    if (analysis.writeAtCursor && !progress.sawFirstDelta) {
      progress.sawFirstDelta = true;
      return "first_delta";
    }

    if (analysis.messageText) {
      const messageLength = analysis.messageText.length;
      if (!progress.sawFirstText) {
        progress.sawFirstText = true;
        progress.maxTextLength = Math.max(progress.maxTextLength, messageLength);
        return "first_text";
      }
      if (analysis.hasCompleteMarkdownJson && !progress.sawCompleteJson) {
        progress.sawCompleteJson = true;
        progress.maxTextLength = Math.max(progress.maxTextLength, messageLength);
        return "complete_markdown_json";
      }
      if (messageLength - progress.maxTextLength >= 1500) {
        progress.maxTextLength = messageLength;
        return "text_growth_milestone";
      }
      progress.maxTextLength = Math.max(progress.maxTextLength, messageLength);
    }

    if (analysis.isEmptyObject || isPureMetadataUpdate(analysis)) {
      return null;
    }

    return null;
  }

  private getFrameProgress(analysis: FrameAnalysis): ResponseProgress {
    const key = frameProgressKey(analysis);
    if (!key) {
      return {
        sawFirstText: false,
        sawFirstDelta: false,
        sawCompleteJson: false,
        maxTextLength: 0,
      };
    }

    const existing = this.responseProgressByKey.get(key);
    if (existing) {
      return existing;
    }

    const created: ResponseProgress = {
      sawFirstText: false,
      sawFirstDelta: false,
      sawCompleteJson: false,
      maxTextLength: 0,
    };
    this.responseProgressByKey.set(key, created);
    if (this.responseProgressByKey.size > 1024) {
      const oldestKey = this.responseProgressByKey.keys().next().value;
      if (typeof oldestKey === "string") {
        this.responseProgressByKey.delete(oldestKey);
      }
    }
    return created;
  }

  private clearFrameProgress(analysis: FrameAnalysis): void {
    const key = frameProgressKey(analysis);
    if (!key) {
      return;
    }
    this.responseProgressByKey.delete(key);
  }

  private async logHttpLike(
    title: string,
    metadata: readonly (readonly [string, string])[],
    headers: readonly [string, string][],
    body: string | null,
    suffix: string,
    statusCode: number | null = null,
  ): Promise<void> {
    if (
      !this.isEnabled ||
      !this.options.debugPath ||
      !this.options.debugPath.trim() ||
      !this.shouldLogByLevel(suffix, statusCode)
    ) {
      return;
    }

    await fs.mkdir(this.options.debugPath, { recursive: true });

    const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
    const sequence = String(++this.sequence).padStart(4, "0");
    const filePath = path.resolve(
      this.options.debugPath,
      `${timestamp}-${sequence}-${suffix}.md`,
    );

    const lines: string[] = [];
    lines.push(`# ${title}`, "");
    for (const [key, value] of metadata) {
      lines.push(`**${key}**: ${value}`);
    }
    lines.push("", "| Header | Value |", "| --- | --- |");
    for (const [header, value] of headers) {
      lines.push(`| ${header} | ${redactHeaderValue(header, value)} |`);
    }
    if (body && body.trim()) {
      lines.push("", "```json", tryPrettyJson(body), "```");
    }
    await fs.writeFile(filePath, lines.join("\n"), "utf8");
  }

  private shouldLogByLevel(suffix: string, statusCode: number | null): boolean {
    if (isRequestSuffix(suffix)) {
      return true;
    }

    if (suffix === "substrate-delta") {
      return this.logLevel === "trace";
    }

    if (suffix === "substrate-response") {
      return this.isLevelEnabled("debug");
    }

    if (suffix === "outgoing-response") {
      if (this.logLevel === "warning") {
        return statusCode !== null && statusCode >= 400 && statusCode <= 499;
      }
      if (this.logLevel === "error") {
        return statusCode !== null && statusCode >= 500 && statusCode <= 599;
      }
      return true;
    }

    if (suffix === "response" || suffix === "response-headers") {
      return this.isLevelEnabled("debug");
    }

    return true;
  }

  private isLevelEnabled(level: LogLevel): boolean {
    return LogLevelPriority[this.logLevel] >= LogLevelPriority[level];
  }
}

function normalizeLogLevel(raw: string | null | undefined): LogLevel {
  if (!raw) {
    return "info";
  }
  const normalized = raw.trim().toLowerCase();
  if (
    normalized === "trace" ||
    normalized === "debug" ||
    normalized === "info" ||
    normalized === "warning" ||
    normalized === "error"
  ) {
    return normalized;
  }
  return "info";
}

function isRequestSuffix(suffix: string): boolean {
  return (
    suffix === "incoming-request" ||
    suffix === "request" ||
    suffix === "substrate-request"
  );
}

function buildFrameLogEntry(
  analysis: FrameAnalysis,
  reason: string,
  includePayload: boolean,
): JsonObject {
  const entry: JsonObject = {
    index: analysis.index,
    reason,
  };
  if (analysis.type !== null) {
    entry.type = analysis.type;
  }
  if (analysis.target) {
    entry.target = analysis.target;
  }
  if (analysis.requestId) {
    entry.requestId = analysis.requestId;
  }
  if (analysis.messageId) {
    entry.messageId = analysis.messageId;
  }
  if (analysis.messageText) {
    entry.messageTextLength = analysis.messageText.length;
    entry.messagePreview = truncate(analysis.messageText, 200);
  }
  if (analysis.writeAtCursor) {
    entry.writeAtCursorLength = analysis.writeAtCursor.length;
    entry.writeAtCursorPreview = truncate(analysis.writeAtCursor, 120);
  }
  if (analysis.error) {
    entry.error = analysis.error;
  }
  if (analysis.resultValue) {
    entry.resultValue = analysis.resultValue;
  }
  if (analysis.resultMessage) {
    entry.resultMessage = analysis.resultMessage;
  }
  if (analysis.hasCursor) {
    entry.hasCursor = true;
  }
  if (analysis.hasCompleteMarkdownJson) {
    entry.hasCompleteMarkdownJson = true;
  }
  if (analysis.isTerminal) {
    entry.isTerminal = true;
  }
  if (includePayload) {
    entry.payload = analysis.json ? analysis.json : analysis.raw;
  }
  return entry;
}

function frameProgressKey(analysis: FrameAnalysis): string | null {
  if (analysis.requestId) {
    return `request:${analysis.requestId}`;
  }
  if (analysis.messageId) {
    return `message:${analysis.messageId}`;
  }
  return null;
}

function isPureMetadataUpdate(analysis: FrameAnalysis): boolean {
  return (
    analysis.type === 1 &&
    analysis.target === "update" &&
    !analysis.messageText &&
    !analysis.writeAtCursor &&
    !analysis.error &&
    !analysis.resultValue &&
    !analysis.resultMessage
  );
}

function analyzeFrame(rawFrame: string, index: number): FrameAnalysis {
  const json = tryParseJsonObject(rawFrame);
  if (!json) {
    return {
      index,
      raw: rawFrame,
      json: null,
      type: null,
      target: null,
      requestId: null,
      messageId: null,
      messageText: null,
      writeAtCursor: null,
      resultValue: null,
      resultMessage: null,
      error: null,
      hasCursor: false,
      isTerminal: false,
      isEmptyObject: false,
      hasCompleteMarkdownJson: false,
    };
  }

  const message = extractBotMessage(json);
  const messageText = message?.text ?? null;
  const writeAtCursor = extractWriteAtCursor(json);
  const type = tryGetInt(json, "type");
  const resultValue = extractResultValue(json);
  const resultMessage = extractResultMessage(json);

  return {
    index,
    raw: rawFrame,
    json,
    type,
    target: tryGetString(json, "target"),
    requestId: extractRequestId(json),
    messageId: message?.id ?? null,
    messageText,
    writeAtCursor,
    resultValue,
    resultMessage,
    error: tryGetString(json, "error"),
    hasCursor: hasCursor(json),
    isTerminal: type === 2 || type === 3 || type === 7,
    isEmptyObject: Object.keys(json).length === 0,
    hasCompleteMarkdownJson: Boolean(
      messageText && CompleteMarkdownJsonPattern.test(messageText),
    ),
  };
}

function splitSignalRFrames(payload: string): string[] {
  const frames = payload
    .split("\u001e")
    .map((part) => part.trim())
    .filter((part) => part.length > 0);
  return frames.length > 0 ? frames : [payload];
}

function extractRequestId(json: JsonObject): string | null {
  const direct = tryGetString(json, "requestId");
  if (direct) {
    return direct;
  }
  if (isJsonObject(json.item)) {
    const itemRequestId = tryGetString(json.item, "requestId");
    if (itemRequestId) {
      return itemRequestId;
    }
  }
  if (!Array.isArray(json.arguments)) {
    return null;
  }
  for (const arg of json.arguments) {
    if (!isJsonObject(arg)) {
      continue;
    }
    const requestId = tryGetString(arg, "requestId");
    if (requestId) {
      return requestId;
    }
    if (isJsonObject(arg.item)) {
      const nested = tryGetString(arg.item, "requestId");
      if (nested) {
        return nested;
      }
    }
  }
  return null;
}

function extractBotMessage(
  json: JsonObject,
): { id: string | null; text: string | null } | null {
  const messages = collectMessages(json);
  let result: { id: string | null; text: string | null } | null = null;
  for (const message of messages) {
    const author = (tryGetString(message, "author") ?? "").toLowerCase();
    if (author !== "bot") {
      continue;
    }
    const text =
      tryGetString(message, "text") ??
      tryGetString(message, "hiddenText") ??
      tryGetString(message, "spokenText");
    if (!text) {
      continue;
    }
    result = {
      id: tryGetString(message, "messageId"),
      text,
    };
  }
  return result;
}

function extractWriteAtCursor(json: JsonObject): string | null {
  if (!Array.isArray(json.arguments)) {
    return null;
  }
  for (const arg of json.arguments) {
    if (!isJsonObject(arg)) {
      continue;
    }
    const delta = tryGetString(arg, "writeAtCursor");
    if (delta) {
      return delta;
    }
  }
  return null;
}

function hasCursor(json: JsonObject): boolean {
  if (!Array.isArray(json.arguments)) {
    return false;
  }
  for (const arg of json.arguments) {
    if (!isJsonObject(arg)) {
      continue;
    }
    if (isJsonObject(arg.cursor)) {
      return true;
    }
  }
  return false;
}

function extractResultValue(json: JsonObject): string | null {
  if (isJsonObject(json.item) && isJsonObject(json.item.result)) {
    const itemValue = tryGetString(json.item.result, "value");
    if (itemValue) {
      return itemValue;
    }
  }
  if (isJsonObject(json.result)) {
    return tryGetString(json.result, "value");
  }
  return null;
}

function extractResultMessage(json: JsonObject): string | null {
  if (isJsonObject(json.item) && isJsonObject(json.item.result)) {
    const itemMessage = tryGetString(json.item.result, "message");
    if (itemMessage) {
      return itemMessage;
    }
  }
  if (isJsonObject(json.result)) {
    return tryGetString(json.result, "message");
  }
  return null;
}

function collectMessages(json: JsonObject): JsonObject[] {
  const messages: JsonObject[] = [];
  const pushArray = (value: unknown) => {
    if (!Array.isArray(value)) {
      return;
    }
    for (const item of value) {
      if (isJsonObject(item)) {
        messages.push(item);
      }
    }
  };

  pushArray(json.messages);
  if (isJsonObject(json.item)) {
    pushArray(json.item.messages);
  }
  if (Array.isArray(json.arguments)) {
    for (const arg of json.arguments) {
      if (!isJsonObject(arg)) {
        continue;
      }
      pushArray(arg.messages);
      if (isJsonObject(arg.item)) {
        pushArray(arg.item.messages);
      }
    }
  }

  return messages;
}

function truncate(value: string, maxLength: number): string {
  if (value.length <= maxLength) {
    return value;
  }
  return `${value.slice(0, maxLength)}...`;
}

function redactHeaderValue(header: string, value: string): string {
  const normalizedHeader = header.trim().toLowerCase();
  if (
    normalizedHeader !== "authorization" &&
    normalizedHeader !== "proxy-authorization"
  ) {
    return value;
  }

  const trimmed = value.trim();
  const bearerPrefix = /^bearer\s+/i;
  if (!bearerPrefix.test(trimmed)) {
    return "[redacted]";
  }

  const token = trimmed.replace(bearerPrefix, "").trim();
  if (!token) {
    return "Bearer [redacted]";
  }

  const prefix = token.slice(0, 4);
  const suffix = token.slice(-3);
  if (token.length <= 8) {
    return `Bearer ${prefix}...`;
  }

  return `Bearer ${prefix}...${suffix}`;
}

function redactUriTokens(uri: string): string {
  try {
    const parsed = new URL(uri);
    if (parsed.searchParams.has("access_token")) {
      const token = parsed.searchParams.get("access_token") ?? "";
      const redacted = redactTokenValue(token);
      parsed.searchParams.set("access_token", redacted);
      return parsed.toString();
    }
  } catch {
    // fall through
  }

  return uri.replace(
    /(access_token=)([^&]+)/gi,
    (_match, prefix, token) => `${prefix}${redactTokenValue(String(token))}`,
  );
}

function redactTokenValue(token: string): string {
  const normalized = token.trim();
  if (!normalized) {
    return "[redacted]";
  }
  const prefix = normalized.slice(0, 4);
  const suffix = normalized.slice(-3);
  if (normalized.length <= 8) {
    return `${prefix}...`;
  }
  return `${prefix}...${suffix}`;
}

function isTrivialSubstrateRequestPayload(payload: string): boolean {
  const frames = splitSignalRFrames(payload);
  if (frames.length === 0) {
    return false;
  }
  return frames.every((frame) => isTrivialSubstrateRequestFrame(frame));
}

function isTrivialSubstrateRequestFrame(frame: string): boolean {
  const json = tryParseJsonObject(frame);
  if (!json) {
    return false;
  }

  const keys = Object.keys(json);
  if (keys.length === 2 && keys.includes("protocol") && keys.includes("version")) {
    return json.protocol === "json" && json.version === 1;
  }

  return keys.length === 1 && keys[0] === "type" && json.type === 6;
}

function isSubstrateResultSuccess(resultValue: string): boolean {
  const normalized = resultValue.toLowerCase();
  return normalized === "success" || normalized === "apologyresponsereturned";
}
