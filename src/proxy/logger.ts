import { promises as fs } from "node:fs";
import path from "node:path";
import type { LogLevel, WrapperOptions } from "./types";
import { tryPrettyJson } from "./utils";

const LogLevelPriority: Record<LogLevel, number> = {
  error: 0,
  warning: 1,
  info: 2,
  debug: 3,
  trace: 4,
};

export class DebugMarkdownLogger {
  private sequence = 0;
  private readonly logLevel: LogLevel;

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
    await this.logHttpLike(
      `Substrate WebSocket ${normalizedDirection}`,
      [
        ["Uri", redactUriTokens(uri)],
        ["Direction", normalizedDirection],
      ],
      [],
      payload,
      `substrate-${normalizedDirection}`,
    );
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

    return this.isLevelEnabled("debug");
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
