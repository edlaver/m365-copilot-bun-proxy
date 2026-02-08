import { randomUUID } from "node:crypto";
import WebSocket from "ws";
import type {
  ChatResult,
  CreateConversationResult,
  JsonObject,
  JsonValue,
  ParsedOpenAiRequest,
  SubstrateStreamUpdate,
  WrapperOptions,
} from "./types";
import {
  extractBearerToken,
  isJsonObject,
  tryGetString,
  tryParseJsonObject,
  tryReadJwtPayload,
  extractGraphErrorMessage,
  tryGetInt,
} from "./utils";
import { DebugMarkdownLogger } from "./logger";

export class CopilotGraphClient {
  constructor(
    private readonly options: WrapperOptions,
    private readonly logger: DebugMarkdownLogger,
  ) {}

  async createConversation(
    authorizationHeader: string,
  ): Promise<CreateConversationResult> {
    const uri = this.buildAbsoluteUri(this.options.createConversationPath);
    const headers = new Headers({
      Authorization: authorizationHeader,
      Accept: "application/json",
      "Content-Type": "application/json",
    });
    const body = "{}";

    await this.logger.logUpstreamRequest(
      "POST",
      uri.toString(),
      headers.entries(),
      body,
    );
    const response = await fetch(uri, { method: "POST", headers, body });
    const rawBody = await response.text();
    await this.logger.logUpstreamResponse(
      response.status,
      uri.toString(),
      response.headers.entries(),
      rawBody,
      true,
    );

    const json = tryParseJsonObject(rawBody);
    return {
      isSuccess: response.ok,
      statusCode: response.status,
      conversationId: tryGetString(json, "id"),
      rawBody,
    };
  }

  async chat(
    authorizationHeader: string,
    conversationId: string,
    payload: JsonObject,
  ): Promise<ChatResult> {
    const uri = this.buildAbsoluteUri(
      resolveConversationPath(this.options.chatPathTemplate, conversationId),
    );
    const headers = new Headers({
      Authorization: authorizationHeader,
      Accept: "application/json",
      "Content-Type": "application/json",
    });
    const body = JSON.stringify(payload);

    await this.logger.logUpstreamRequest(
      "POST",
      uri.toString(),
      headers.entries(),
      body,
    );
    const response = await fetch(uri, { method: "POST", headers, body });
    const rawBody = await response.text();
    await this.logger.logUpstreamResponse(
      response.status,
      uri.toString(),
      response.headers.entries(),
      rawBody,
      true,
    );

    return {
      isSuccess: response.ok,
      statusCode: response.status,
      responseJson: tryParseJsonObject(rawBody),
      rawBody,
      assistantText: null,
      conversationId: null,
    };
  }

  async chatOverStream(
    authorizationHeader: string,
    conversationId: string,
    payload: JsonObject,
  ): Promise<Response> {
    const uri = this.buildAbsoluteUri(
      resolveConversationPath(
        this.options.chatOverStreamPathTemplate,
        conversationId,
      ),
    );
    const headers = new Headers({
      Authorization: authorizationHeader,
      Accept: "text/event-stream",
      "Content-Type": "application/json",
    });
    const body = JSON.stringify(payload);

    await this.logger.logUpstreamRequest(
      "POST",
      uri.toString(),
      headers.entries(),
      body,
    );
    const response = await fetch(uri, { method: "POST", headers, body });
    await this.logger.logUpstreamResponse(
      response.status,
      uri.toString(),
      response.headers.entries(),
      null,
      false,
    );
    return response;
  }

  private buildAbsoluteUri(relativePath: string): URL {
    try {
      return new URL(relativePath);
    } catch {
      const baseUrl =
        this.options.graphBaseUrl?.trim() || "https://graph.microsoft.com";
      const normalized = baseUrl.endsWith("/") ? baseUrl : `${baseUrl}/`;
      return new URL(relativePath.replace(/^\/+/, ""), normalized);
    }
  }
}

export class CopilotSubstrateClient {
  constructor(
    private readonly options: WrapperOptions,
    private readonly logger: DebugMarkdownLogger,
  ) {}

  createConversation(): CreateConversationResult {
    return {
      isSuccess: true,
      statusCode: 200,
      conversationId: randomUUID(),
      rawBody: "",
    };
  }

  async chat(
    authorizationHeader: string,
    conversationId: string,
    request: ParsedOpenAiRequest,
    isStartOfSession: boolean,
  ): Promise<ChatResult> {
    return this.chatCore(
      authorizationHeader,
      conversationId,
      request,
      isStartOfSession,
      null,
    );
  }

  async chatStream(
    authorizationHeader: string,
    conversationId: string,
    request: ParsedOpenAiRequest,
    isStartOfSession: boolean,
    onStreamUpdate: (update: SubstrateStreamUpdate) => Promise<void>,
  ): Promise<ChatResult> {
    return this.chatCore(
      authorizationHeader,
      conversationId,
      request,
      isStartOfSession,
      onStreamUpdate,
    );
  }

  private async chatCore(
    authorizationHeader: string,
    conversationId: string,
    request: ParsedOpenAiRequest,
    isStartOfSession: boolean,
    onStreamUpdate: ((update: SubstrateStreamUpdate) => Promise<void>) | null,
  ): Promise<ChatResult> {
    const rawToken = extractBearerToken(authorizationHeader);
    if (!rawToken) {
      return buildFailure(401, "Missing Bearer token.");
    }

    const tokenPayload = tryReadJwtPayload(rawToken);
    if (!tokenPayload) {
      return buildFailure(
        400,
        "Authorization token must be a JWT so oid/tid can be resolved for Substrate.",
      );
    }

    const objectId = tryGetString(tokenPayload, "oid");
    const tenantId = tryGetString(tokenPayload, "tid");
    if (!objectId || !tenantId) {
      return buildFailure(
        400,
        "Authorization token is missing required 'oid' and/or 'tid' claims for Substrate.",
      );
    }

    const clientRequestId = randomUUID();
    const sessionId = randomUUID();
    const requestUri = buildSubstrateHubUri(
      this.options,
      objectId,
      tenantId,
      rawToken,
      clientRequestId,
      sessionId,
      conversationId,
    );
    const timeoutMs =
      (this.options.substrate.invocationTimeoutSeconds > 0
        ? this.options.substrate.invocationTimeoutSeconds
        : 120) * 1000;

    const ws = await connectWebSocket(
      requestUri,
      this.options.substrate.origin ?? undefined,
      timeoutMs,
      this.options.substrate.keepAliveSeconds > 0
        ? this.options.substrate.keepAliveSeconds * 1000
        : 15_000,
    ).catch((error) => error as Error);
    if (ws instanceof Error) {
      return buildFailure(
        502,
        `Substrate websocket request failed. ${ws.message}`,
      );
    }

    const transcript: string[] = [];
    try {
      await sendFrame(ws, requestUri, this.logger, {
        protocol: "json",
        version: 1,
      });
      const handshakePayload = await receiveMessage(ws, timeoutMs);
      if (handshakePayload === null) {
        return buildFailure(
          502,
          "Substrate websocket closed during handshake.",
        );
      }
      await this.logger.logSubstrateFrame(
        requestUri.toString(),
        "response",
        handshakePayload,
      );
      transcript.push(handshakePayload);

      for (const frame of splitFrames(handshakePayload)) {
        const frameJson = tryParseJsonObject(frame);
        const handshakeError = tryGetString(frameJson, "error");
        if (handshakeError) {
          return buildFailure(
            502,
            `Substrate handshake failed. ${handshakeError}`,
          );
        }
      }

      await sendFrame(ws, requestUri, this.logger, { type: 6 });
      await sendFrame(
        ws,
        requestUri,
        this.logger,
        buildInvocationPayload(
          request,
          conversationId,
          sessionId,
          clientRequestId,
          isStartOfSession,
          this.options,
        ),
      );

      let assistantText = "";
      let deltaBuilder = "";
      let resolvedConversationId = conversationId;
      let responseError: string | null = null;
      let completed = false;

      while (!completed && ws.readyState === WebSocket.OPEN) {
        const payload = await receiveMessage(ws, timeoutMs);
        if (payload === null) {
          break;
        }
        await this.logger.logSubstrateFrame(
          requestUri.toString(),
          "response",
          payload,
        );
        transcript.push(payload);

        for (const frame of splitFrames(payload)) {
          if (!frame.trim()) {
            continue;
          }
          const json = tryParseJsonObject(frame);
          if (!json) {
            continue;
          }

          const extractedConversationId = extractSubstrateConversationId(json);
          if (
            extractedConversationId &&
            extractedConversationId !== resolvedConversationId
          ) {
            resolvedConversationId = extractedConversationId;
            if (onStreamUpdate) {
              await onStreamUpdate({
                deltaText: null,
                conversationId: resolvedConversationId,
              });
            }
          }

          const extractedAssistantText = extractSubstrateAssistantText(json);
          if (extractedAssistantText) {
            assistantText = extractedAssistantText;
          } else {
            const deltaText = extractSubstrateDeltaText(json);
            if (deltaText) {
              deltaBuilder += deltaText;
              if (onStreamUpdate) {
                await onStreamUpdate({
                  deltaText,
                  conversationId: resolvedConversationId,
                });
              }
            }
          }

          const frameError = tryGetString(json, "error");
          if (frameError) {
            responseError = frameError;
          }
          const resultValue = extractSubstrateResultValue(json);
          if (resultValue && !isSubstrateResultSuccess(resultValue)) {
            responseError =
              extractSubstrateResultMessage(json) ??
              `Substrate returned result '${resultValue}'.`;
          }

          const frameType = tryGetInt(json, "type");
          if (
            frameType !== null &&
            (frameType === 2 || frameType === 3 || frameType === 7)
          ) {
            completed = true;
            break;
          }
        }
      }

      try {
        ws.close(1000, "completed");
      } catch {
        // ignore
      }

      if (responseError && !assistantText) {
        return buildFailure(502, `Substrate chat failed. ${responseError}`);
      }

      if (!assistantText && deltaBuilder) {
        assistantText = deltaBuilder;
      }

      if (!assistantText) {
        return buildFailure(
          502,
          "Substrate chat returned no assistant content.",
        );
      }

      return {
        isSuccess: true,
        statusCode: 200,
        responseJson: buildNormalizedConversation(
          resolvedConversationId,
          request.promptText,
          assistantText,
        ),
        rawBody: transcript.join("\n"),
        assistantText,
        conversationId: resolvedConversationId,
      };
    } catch (error) {
      const message = String(error);
      if (message.toLowerCase().includes("timeout")) {
        return buildFailure(504, "Substrate websocket request timed out.");
      }
      return buildFailure(
        502,
        `Unexpected Substrate websocket failure. ${message}`,
      );
    } finally {
      try {
        ws.close();
      } catch {
        // ignore
      }
    }
  }
}

function resolveConversationPath(
  pathTemplate: string,
  conversationId: string,
): string {
  const template = pathTemplate?.trim()
    ? pathTemplate
    : "/beta/copilot/conversations/{conversationId}/chat";
  return template.replaceAll(
    "{conversationId}",
    encodeURIComponent(conversationId),
  );
}

function buildFailure(statusCode: number, message: string): ChatResult {
  return {
    isSuccess: false,
    statusCode,
    responseJson: null,
    rawBody: JSON.stringify({ message }),
    assistantText: null,
    conversationId: null,
  };
}

function buildSubstrateHubUri(
  options: WrapperOptions,
  objectId: string,
  tenantId: string,
  accessToken: string,
  clientRequestId: string,
  sessionId: string,
  conversationId: string,
): URL {
  let baseHub =
    options.substrate.hubPath?.trim() ||
    "wss://substrate.office.com/m365Copilot/Chathub/";
  if (!baseHub.endsWith("/")) {
    baseHub += "/";
  }
  const hubUri = new URL(
    `${encodeURIComponent(objectId)}@${encodeURIComponent(tenantId)}`,
    baseHub,
  );

  const query = new URLSearchParams({
    ClientRequestId: clientRequestId,
    "X-SessionId": sessionId,
    ConversationId: conversationId,
    access_token: accessToken,
  });

  if (options.substrate.source?.trim()) {
    const sourceValue = options.substrate.quoteSourceInQuery
      ? `"${options.substrate.source}"`
      : options.substrate.source;
    query.set("source", sourceValue);
  }
  if (options.substrate.scenario?.trim()) {
    query.set("scenario", options.substrate.scenario);
  }
  if (options.substrate.product?.trim()) {
    query.set("product", options.substrate.product);
  }
  if (options.substrate.agentHost?.trim()) {
    query.set("agentHost", options.substrate.agentHost);
  }
  if (options.substrate.licenseType?.trim()) {
    query.set("licenseType", options.substrate.licenseType);
  }
  if (options.substrate.agent?.trim()) {
    query.set("agent", options.substrate.agent);
  }
  if (options.substrate.variants?.trim()) {
    query.set("variants", options.substrate.variants);
  }

  hubUri.search = query.toString();
  return hubUri;
}

function buildInvocationPayload(
  request: ParsedOpenAiRequest,
  conversationId: string,
  sessionId: string,
  clientRequestId: string,
  isStartOfSession: boolean,
  options: WrapperOptions,
): JsonObject {
  const message: JsonObject = {
    author: "user",
    text: buildPromptWithAdditionalContext(request),
    inputMethod: "Keyboard",
    messageType: "Chat",
    requestId: clientRequestId,
    messageId: randomUUID(),
    locale: options.substrate.locale?.trim() || "en-US",
    experienceType: options.substrate.experienceType?.trim() || "Default",
  };

  if (options.substrate.entityAnnotationTypes.length > 0) {
    message.entityAnnotationTypes = options.substrate.entityAnnotationTypes
      .map((v) => v.trim())
      .filter((v) => v.length > 0);
  }

  const locationInfo = buildLocationInfo(request.locationHint);
  if (locationInfo) {
    message.locationInfo = locationInfo;
  }

  const argument: JsonObject = {
    source: options.substrate.source?.trim() || "officeweb",
    clientCorrelationId: clientRequestId,
    sessionId,
    conversationId,
    traceId: randomUUID().replaceAll("-", ""),
    isStartOfSession,
    productThreadType: options.substrate.productThreadType?.trim() || "Office",
    clientInfo: {
      clientPlatform: options.substrate.clientPlatform?.trim() || "web",
    },
    message,
  };

  if (options.substrate.optionsSets.length > 0) {
    argument.optionsSets = options.substrate.optionsSets
      .map((v) => v.trim())
      .filter((v) => v.length > 0);
  }
  if (options.substrate.allowedMessageTypes.length > 0) {
    argument.allowedMessageTypes = options.substrate.allowedMessageTypes
      .map((v) => v.trim())
      .filter((v) => v.length > 0);
  }
  if (request.contextualResources) {
    argument.contextualResources = request.contextualResources;
  }

  return {
    arguments: [argument],
    invocationId: "0",
    target: options.substrate.invocationTarget?.trim() || "update",
    type:
      options.substrate.invocationType > 0
        ? options.substrate.invocationType
        : 1,
  };
}

function buildPromptWithAdditionalContext(
  request: ParsedOpenAiRequest,
): string {
  if (request.additionalContext.length === 0) {
    return request.promptText;
  }
  const lines = ["Context:"];
  for (const ctx of request.additionalContext) {
    if (!ctx.text.trim()) {
      continue;
    }
    lines.push(`${ctx.description ? `${ctx.description}: ` : ""}${ctx.text}`);
  }
  lines.push("", `User: ${request.promptText}`);
  return lines.join("\n");
}

function buildLocationInfo(locationHint: JsonObject): JsonObject | null {
  const timeZone = tryGetString(locationHint, "timeZone");
  if (!timeZone) {
    return null;
  }
  const locationInfo: JsonObject = {
    timeZone,
    timeZoneOffset: resolveTimeZoneOffsetMinutes(timeZone),
  };
  const countryOrRegion = tryGetString(locationHint, "countryOrRegion");
  if (countryOrRegion) {
    locationInfo.countryOrRegion = countryOrRegion;
  }
  return locationInfo;
}

function resolveTimeZoneOffsetMinutes(timeZoneId: string): number {
  try {
    const now = new Date();
    const zoned = new Date(
      now.toLocaleString("en-US", { timeZone: timeZoneId }),
    );
    return Math.round((zoned.getTime() - now.getTime()) / 60000);
  } catch {
    return 0;
  }
}

function splitFrames(payload: string): string[] {
  return payload
    .split("\u001e")
    .map((frame) => frame.trim())
    .filter((frame) => frame.length > 0);
}

async function connectWebSocket(
  url: URL,
  origin: string | undefined,
  timeoutMs: number,
  keepAliveMs: number,
): Promise<WebSocket> {
  return new Promise<WebSocket>((resolve, reject) => {
    const ws = new WebSocket(url, {
      handshakeTimeout: timeoutMs,
      headers: origin ? { Origin: origin } : undefined,
    });
    const timeout = setTimeout(() => {
      try {
        ws.terminate();
      } catch {
        // ignore
      }
      reject(new Error("timeout"));
    }, timeoutMs);

    ws.once("open", () => {
      clearTimeout(timeout);
      if (keepAliveMs > 0) {
        const timer = setInterval(() => {
          if (ws.readyState !== WebSocket.OPEN) {
            clearInterval(timer);
            return;
          }
          try {
            ws.ping();
          } catch {
            clearInterval(timer);
          }
        }, keepAliveMs);
        ws.once("close", () => clearInterval(timer));
      }
      resolve(ws);
    });
    ws.once("error", (error) => {
      clearTimeout(timeout);
      reject(error);
    });
  });
}

async function sendFrame(
  ws: WebSocket,
  requestUri: URL,
  logger: DebugMarkdownLogger,
  frame: JsonObject,
): Promise<void> {
  const payload = `${JSON.stringify(frame)}\u001e`;
  await logger.logSubstrateFrame(requestUri.toString(), "request", payload);
  await new Promise<void>((resolve, reject) => {
    ws.send(payload, (error) => {
      if (error) {
        reject(error);
      } else {
        resolve();
      }
    });
  });
}

async function receiveMessage(
  ws: WebSocket,
  timeoutMs: number,
): Promise<string | null> {
  return new Promise<string | null>((resolve) => {
    const timer = setTimeout(() => {
      cleanup();
      resolve(null);
    }, timeoutMs);

    const cleanup = () => {
      clearTimeout(timer);
      ws.off("message", onMessage);
      ws.off("close", onClose);
      ws.off("error", onError);
    };

    const onMessage = (data: WebSocket.RawData) => {
      cleanup();
      if (typeof data === "string") {
        resolve(data);
        return;
      }
      if (Buffer.isBuffer(data)) {
        resolve(data.toString("utf8"));
        return;
      }
      if (Array.isArray(data)) {
        resolve(Buffer.concat(data).toString("utf8"));
        return;
      }
      resolve(Buffer.from(data).toString("utf8"));
    };
    const onClose = () => {
      cleanup();
      resolve(null);
    };
    const onError = () => {
      cleanup();
      resolve(null);
    };

    ws.once("message", onMessage);
    ws.once("close", onClose);
    ws.once("error", onError);
  });
}

function extractSubstrateAssistantText(envelope: JsonObject): string | null {
  const messages = collectMessageObjects(envelope);
  let fallback: string | null = null;
  for (const message of messages) {
    if ((tryGetString(message, "author") ?? "").toLowerCase() !== "bot") {
      continue;
    }
    const messageType = (
      tryGetString(message, "messageType") ?? "Chat"
    ).toLowerCase();
    if (messageType !== "chat" && messageType !== "disengaged") {
      continue;
    }
    const text =
      tryGetString(message, "text") ??
      tryGetString(message, "hiddenText") ??
      tryGetString(message, "spokenText");
    if (text) {
      fallback = text;
    }
  }
  return fallback ?? extractSubstrateResultMessage(envelope);
}

function extractSubstrateDeltaText(envelope: JsonObject): string | null {
  const args = envelope.arguments;
  if (!Array.isArray(args)) {
    return null;
  }
  for (const arg of args) {
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

function extractSubstrateConversationId(envelope: JsonObject): string | null {
  const direct = tryGetString(envelope, "conversationId");
  if (direct) {
    return direct;
  }

  const item = envelope.item;
  if (isJsonObject(item)) {
    const itemId = tryGetString(item, "conversationId");
    if (itemId) {
      return itemId;
    }
  }

  const args = envelope.arguments;
  if (!Array.isArray(args)) {
    return null;
  }
  for (const arg of args) {
    if (!isJsonObject(arg)) {
      continue;
    }
    const argId = tryGetString(arg, "conversationId");
    if (argId) {
      return argId;
    }
    const argItem = arg.item;
    if (isJsonObject(argItem)) {
      const argItemId = tryGetString(argItem, "conversationId");
      if (argItemId) {
        return argItemId;
      }
    }
  }
  return null;
}

function extractSubstrateResultMessage(envelope: JsonObject): string | null {
  const item = envelope.item;
  if (isJsonObject(item) && isJsonObject(item.result)) {
    const itemMessage = tryGetString(item.result, "message");
    if (itemMessage) {
      return itemMessage;
    }
  }
  if (isJsonObject(envelope.result)) {
    return tryGetString(envelope.result, "message");
  }
  return null;
}

function extractSubstrateResultValue(envelope: JsonObject): string | null {
  const item = envelope.item;
  if (isJsonObject(item) && isJsonObject(item.result)) {
    const itemValue = tryGetString(item.result, "value");
    if (itemValue) {
      return itemValue;
    }
  }
  if (isJsonObject(envelope.result)) {
    return tryGetString(envelope.result, "value");
  }
  return null;
}

function isSubstrateResultSuccess(resultValue: string): boolean {
  const normalized = resultValue.toLowerCase();
  return normalized === "success" || normalized === "apologyresponsereturned";
}

function buildNormalizedConversation(
  conversationId: string,
  prompt: string,
  assistantText: string,
): JsonObject {
  return {
    id: conversationId,
    messages: [
      { author: "user", text: prompt },
      { author: "assistant", text: assistantText },
    ],
  };
}

function collectMessageObjects(envelope: JsonObject): JsonObject[] {
  const messages: JsonObject[] = [];
  const pushArray = (value: JsonValue | undefined) => {
    if (!Array.isArray(value)) {
      return;
    }
    for (const item of value) {
      if (isJsonObject(item)) {
        messages.push(item);
      }
    }
  };

  pushArray(envelope.messages);
  if (isJsonObject(envelope.item)) {
    pushArray(envelope.item.messages);
  }
  const args = envelope.arguments;
  if (Array.isArray(args)) {
    for (const arg of args) {
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

export function summarizeUpstreamFailure(
  statusCode: number,
  responseBody: string | null,
  fallbackMessage: string,
): { statusCode: number; message: string } {
  const details = extractGraphErrorMessage(responseBody);
  const message = details ? `${fallbackMessage} ${details}` : fallbackMessage;
  const normalizedStatusCode =
    statusCode >= 400 && statusCode <= 599 ? statusCode : 502;
  return { statusCode: normalizedStatusCode, message };
}
