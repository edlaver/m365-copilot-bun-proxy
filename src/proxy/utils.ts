import type { JsonObject, JsonPayload, JsonValue } from "./types";

export function nowUnix(): number {
  return Math.floor(Date.now() / 1000);
}

export function normalizeNullableString(
  value: string | null | undefined,
): string | null {
  return value && value.trim() ? value.trim() : null;
}

export function firstNonEmpty(
  ...values: Array<string | null | undefined>
): string | null {
  for (const value of values) {
    if (value && value.trim()) {
      return value;
    }
  }
  return null;
}

export function parseBooleanString(value: string | null | undefined): boolean {
  if (!value || !value.trim()) {
    return false;
  }
  return value.trim().toLowerCase() === "true";
}

export function isJsonObject(value: unknown): value is JsonObject {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

export function cloneJsonValue<T extends JsonValue>(value: T): T {
  if (value === null) {
    return value;
  }
  if (Array.isArray(value)) {
    return value.map((item) => cloneJsonValue(item)) as T;
  }
  if (isJsonObject(value)) {
    const clone: JsonObject = {};
    for (const [key, childValue] of Object.entries(value)) {
      clone[key] = cloneJsonValue(childValue);
    }
    return clone as T;
  }
  return value;
}

export function deepMerge(target: JsonObject, source: JsonObject): JsonObject {
  for (const [key, sourceValue] of Object.entries(source)) {
    const targetValue = target[key];
    if (isJsonObject(sourceValue) && isJsonObject(targetValue)) {
      deepMerge(targetValue, sourceValue);
      continue;
    }
    target[key] = cloneJsonValue(sourceValue);
  }
  return target;
}

export function setDeepValue(
  root: JsonObject,
  pathParts: string[],
  value: JsonValue,
): void {
  let cursor: JsonObject = root;
  for (let index = 0; index < pathParts.length; index++) {
    const part = pathParts[index];
    const isLast = index === pathParts.length - 1;

    if (isLast) {
      cursor[part] = value;
      return;
    }

    const existing = cursor[part];
    if (!isJsonObject(existing)) {
      cursor[part] = {};
    }
    cursor = cursor[part] as JsonObject;
  }
}

export function parseEnvValue(raw: string): JsonValue {
  const trimmed = raw.trim();
  if (!trimmed) {
    return "";
  }
  if (trimmed === "true") {
    return true;
  }
  if (trimmed === "false") {
    return false;
  }
  if (/^-?\d+(\.\d+)?$/.test(trimmed)) {
    return Number.parseFloat(trimmed);
  }
  if (
    (trimmed.startsWith("{") && trimmed.endsWith("}")) ||
    (trimmed.startsWith("[") && trimmed.endsWith("]"))
  ) {
    try {
      return JSON.parse(trimmed) as JsonValue;
    } catch {
      // ignore
    }
  }
  if (trimmed.includes(",")) {
    return trimmed.split(",").map((part) => part.trim());
  }
  return trimmed;
}

export function tryParseJsonObject(
  rawJson: string | null | undefined,
): JsonObject | null {
  if (!rawJson || !rawJson.trim()) {
    return null;
  }
  try {
    const parsed = JSON.parse(rawJson) as JsonValue;
    return isJsonObject(parsed) ? parsed : null;
  } catch {
    return null;
  }
}

export async function tryReadJsonPayload(
  request: Request,
): Promise<JsonPayload | null> {
  const contentType = request.headers.get("content-type");
  if (!contentType || !contentType.toLowerCase().includes("application/json")) {
    return null;
  }
  const rawText = await request.text();
  if (!rawText.trim()) {
    return null;
  }
  const json = tryParseJsonObject(rawText);
  if (!json) {
    return null;
  }
  return { json, rawText };
}

export function tryGetString(
  json: JsonObject | null | undefined,
  propertyName: string,
): string | null {
  if (!json || !(propertyName in json)) {
    return null;
  }
  const value = json[propertyName];
  if (typeof value !== "string") {
    return null;
  }
  return value.trim() ? value.trim() : null;
}

export function tryGetBoolean(
  json: JsonObject | null | undefined,
  propertyName: string,
): boolean | null {
  if (!json || !(propertyName in json)) {
    return null;
  }
  const value = json[propertyName];
  if (typeof value === "boolean") {
    return value;
  }
  if (typeof value === "number") {
    return value !== 0;
  }
  if (typeof value === "string") {
    const lowered = value.trim().toLowerCase();
    if (lowered === "true") {
      return true;
    }
    if (lowered === "false") {
      return false;
    }
  }
  return null;
}

export function tryGetInt(
  json: JsonObject | null | undefined,
  propertyName: string,
): number | null {
  if (!json || !(propertyName in json)) {
    return null;
  }
  const value = json[propertyName];
  if (typeof value === "number" && Number.isFinite(value)) {
    return Math.trunc(value);
  }
  if (typeof value === "string") {
    const parsed = Number.parseInt(value, 10);
    return Number.isFinite(parsed) ? parsed : null;
  }
  return null;
}

export function tryGetDouble(
  json: JsonObject | null | undefined,
  propertyName: string,
): number | null {
  if (!json || !(propertyName in json)) {
    return null;
  }
  const value = json[propertyName];
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }
  if (typeof value === "string") {
    const parsed = Number.parseFloat(value);
    return Number.isFinite(parsed) ? parsed : null;
  }
  return null;
}

export function normalizeBearerToken(
  authorizationHeader: string | null | undefined,
): string | null {
  if (!authorizationHeader || !authorizationHeader.trim()) {
    return null;
  }
  const trimmed = authorizationHeader.trim();
  if (trimmed.toLowerCase().startsWith("bearer ")) {
    const token = trimmed.slice("Bearer ".length).trim();
    return token ? `Bearer ${token}` : null;
  }
  return `Bearer ${trimmed}`;
}

export function extractBearerToken(
  authorizationHeader: string | null | undefined,
): string | null {
  const normalized = normalizeBearerToken(authorizationHeader);
  if (!normalized) {
    return null;
  }
  return normalized.toLowerCase().startsWith("bearer ")
    ? normalized.slice("Bearer ".length).trim()
    : normalized.trim();
}

export function decodeBase64Url(encoded: string): Uint8Array {
  const normalized = encoded.replaceAll("-", "+").replaceAll("_", "/");
  const paddingLength = normalized.length % 4;
  const padded =
    paddingLength > 0
      ? normalized.padEnd(normalized.length + (4 - paddingLength), "=")
      : normalized;
  return Uint8Array.from(Buffer.from(padded, "base64"));
}

export function tryReadJwtPayload(rawToken: string): JsonObject | null {
  if (!rawToken || !rawToken.trim()) {
    return null;
  }
  const parts = rawToken.split(".");
  if (parts.length < 2 || !parts[1]) {
    return null;
  }
  try {
    const payloadBytes = decodeBase64Url(parts[1]);
    const payloadText = new TextDecoder().decode(payloadBytes);
    return tryParseJsonObject(payloadText);
  } catch {
    return null;
  }
}

export function computeTrailingDelta(
  alreadyEmitted: string,
  latestText: string,
): string {
  if (!latestText) {
    return "";
  }
  if (!alreadyEmitted) {
    return latestText;
  }
  if (latestText.length < alreadyEmitted.length) {
    return "";
  }
  return latestText.startsWith(alreadyEmitted)
    ? latestText.slice(alreadyEmitted.length)
    : "";
}

export function extractGraphErrorMessage(
  graphBody: string | null | undefined,
): string | null {
  if (!graphBody || !graphBody.trim()) {
    return null;
  }
  const json = tryParseJsonObject(graphBody);
  if (json) {
    const error = json.error;
    if (isJsonObject(error)) {
      const message = tryGetString(error, "message");
      if (message) {
        return message;
      }
    }
    const directMessage = tryGetString(json, "message");
    if (directMessage) {
      return directMessage;
    }
  }
  return graphBody.length > 400 ? `${graphBody.slice(0, 400)}...` : graphBody;
}

export function parseListenUrl(listenUrl: string): {
  hostname: string;
  port: number;
} {
  const url = new URL(listenUrl);
  return {
    hostname: url.hostname || "0.0.0.0",
    port: url.port ? Number.parseInt(url.port, 10) : 4000,
  };
}

export function tryPrettyJson(rawPayload: string): string {
  if (!rawPayload.trim()) {
    return "";
  }
  const trimmed = rawPayload.trim();
  const cleaned = trimmed.replace(
    /^[\s\u0000-\u001f]+|[\s\u0000-\u001f]+$/g,
    "",
  );
  try {
    return JSON.stringify(JSON.parse(cleaned), null, 2);
  } catch {
    try {
      return JSON.stringify(JSON.parse(trimmed), null, 2);
    } catch {
      return rawPayload;
    }
  }
}

export async function* readSseEvents(
  stream: ReadableStream<Uint8Array>,
): AsyncGenerator<{ event: string; data: string }> {
  const reader = stream.getReader();
  const decoder = new TextDecoder();
  let buffer = "";
  let eventName = "message";
  let dataBuffer: string[] = [];

  while (true) {
    const { done, value } = await reader.read();
    if (done) {
      break;
    }
    buffer += decoder.decode(value, { stream: true });

    while (true) {
      const nl = buffer.indexOf("\n");
      if (nl < 0) {
        break;
      }
      const rawLine = buffer.slice(0, nl);
      buffer = buffer.slice(nl + 1);
      const line = rawLine.endsWith("\r") ? rawLine.slice(0, -1) : rawLine;

      if (line.length === 0) {
        if (dataBuffer.length > 0) {
          yield { event: eventName, data: dataBuffer.join("\n") };
          dataBuffer = [];
        }
        eventName = "message";
        continue;
      }
      if (line.startsWith("event:")) {
        eventName = line.slice("event:".length).trim();
        continue;
      }
      if (line.startsWith("data:")) {
        dataBuffer.push(line.slice("data:".length).trimStart());
      }
    }
  }

  if (dataBuffer.length > 0) {
    yield { event: eventName, data: dataBuffer.join("\n") };
  }
}
