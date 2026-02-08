import { randomUUID } from "node:crypto";
import {
  ResponseFormatTypes,
  ToolChoiceModes,
  type JsonObject,
  type JsonValue,
  type OpenAiAssistantResponse,
  type OpenAiAssistantToolCall,
  type OpenAiResponseFormat,
  type OpenAiTooling,
  type ParsedOpenAiRequest,
} from "./types";
import {
  cloneJsonValue,
  computeTrailingDelta,
  isJsonObject,
  nowUnix,
  tryParseJsonObject,
} from "./utils";

export function buildChatCompletion(
  model: string,
  assistantResponse: OpenAiAssistantResponse,
  conversationId: string,
  includeConversationId: boolean,
): JsonObject {
  const message: JsonObject = { role: "assistant" };
  if (assistantResponse.toolCalls.length > 0) {
    message.content = null;
    message.tool_calls = assistantResponse.toolCalls.map((toolCall) => ({
      id: toolCall.id,
      type: "function",
      function: {
        name: toolCall.name,
        arguments: toolCall.argumentsJson,
      },
    }));
  } else {
    message.content = assistantResponse.content ?? "";
  }

  const response: JsonObject = {
    id: `chatcmpl-${randomUUID().replaceAll("-", "")}`,
    object: "chat.completion",
    created: nowUnix(),
    model,
    choices: [
      {
        index: 0,
        message,
        finish_reason: assistantResponse.finishReason,
      },
    ],
  };

  if (includeConversationId) {
    response.conversation_id = conversationId;
  }

  return response;
}

export function buildChatCompletionChunk(
  completionId: string,
  created: number,
  model: string,
  role: string | null,
  contentDelta: string | null,
  finishReason: string | null,
  conversationId: string | null,
  toolCallsDelta?: JsonValue,
): JsonObject {
  const delta: JsonObject = {};
  if (role?.trim()) {
    delta.role = role;
  }
  if (contentDelta !== null) {
    delta.content = contentDelta;
  }
  if (toolCallsDelta !== undefined) {
    delta.tool_calls = toolCallsDelta;
  }

  const chunk: JsonObject = {
    id: completionId,
    object: "chat.completion.chunk",
    created,
    model,
    choices: [
      {
        index: 0,
        delta,
        finish_reason: finishReason,
      },
    ],
  };

  if (conversationId?.trim()) {
    chunk.conversation_id = conversationId;
  }

  return chunk;
}

export function requiresBufferedAssistantResponse(request: ParsedOpenAiRequest): boolean {
  return request.tooling.tools.length > 0 || request.responseFormat !== null;
}

export function buildAssistantResponse(
  request: ParsedOpenAiRequest,
  assistantText: string,
): OpenAiAssistantResponse {
  const normalizedText = assistantText ?? "";
  if (
    request.tooling.tools.length > 0 &&
    request.tooling.toolChoiceMode !== ToolChoiceModes.None
  ) {
    const toolCalls = tryExtractToolCalls(normalizedText, request.tooling);
    if (toolCalls.length > 0) {
      return {
        content: null,
        toolCalls,
        finishReason: "tool_calls",
      };
    }
  }

  return {
    content: normalizeStructuredContent(normalizedText, request.responseFormat),
    toolCalls: [],
    finishReason: "stop",
  };
}

function tryExtractToolCalls(
  assistantText: string,
  tooling: OpenAiTooling,
): OpenAiAssistantToolCall[] {
  for (const candidate of enumerateJsonCandidates(assistantText)) {
    const node = tryParseJsonNode(candidate);
    if (node === null) {
      continue;
    }
    const calls = extractToolCallsFromNode(node, tooling);
    if (calls.length > 0) {
      return calls;
    }
  }
  return [];
}

function extractToolCallsFromNode(node: JsonValue, tooling: OpenAiTooling): OpenAiAssistantToolCall[] {
  if (Array.isArray(node)) {
    return extractToolCallsFromArray(node, tooling);
  }
  if (!isJsonObject(node)) {
    return [];
  }

  const toolCalls = node.tool_calls;
  if (Array.isArray(toolCalls)) {
    return extractToolCallsFromArray(toolCalls, tooling);
  }

  const singleCall = tryBuildToolCall(node, tooling);
  return singleCall ? [singleCall] : [];
}

function extractToolCallsFromArray(
  toolCallsArray: JsonValue[],
  tooling: OpenAiTooling,
): OpenAiAssistantToolCall[] {
  const toolCalls: OpenAiAssistantToolCall[] = [];
  for (const item of toolCallsArray) {
    if (!isJsonObject(item)) {
      continue;
    }
    const toolCall = tryBuildToolCall(item, tooling);
    if (toolCall) {
      toolCalls.push(toolCall);
    }
  }
  return toolCalls;
}

function tryBuildToolCall(
  callObject: JsonObject,
  tooling: OpenAiTooling,
): OpenAiAssistantToolCall | null {
  const functionObject = isJsonObject(callObject.function) ? callObject.function : null;

  const name = pickString(
    callObject.name,
    functionObject?.name,
    callObject.tool_name,
  );
  if (!name) {
    return null;
  }

  if (
    tooling.toolChoiceMode === ToolChoiceModes.Function &&
    tooling.toolChoiceFunctionName &&
    name !== tooling.toolChoiceFunctionName
  ) {
    return null;
  }

  if (
    tooling.tools.length > 0 &&
    !tooling.tools.some((tool) => tool.name === name)
  ) {
    return null;
  }

  const argumentsJson = normalizeArgumentsJson(callObject.arguments ?? functionObject?.arguments ?? null);
  const id = pickString(callObject.id) ?? `call_${randomUUID().replaceAll("-", "")}`;
  return { id, name, argumentsJson };
}

function normalizeArgumentsJson(argumentsNode: JsonValue | null): string {
  if (argumentsNode === null) {
    return "{}";
  }
  if (typeof argumentsNode === "string") {
    if (!argumentsNode.trim()) {
      return "{}";
    }
    const parsed = tryParseJsonNode(argumentsNode);
    if (parsed !== null) {
      return JSON.stringify(parsed);
    }
    return JSON.stringify({ input: argumentsNode });
  }
  if (Array.isArray(argumentsNode) || isJsonObject(argumentsNode)) {
    return JSON.stringify(argumentsNode);
  }
  return JSON.stringify({ value: argumentsNode });
}

function normalizeStructuredContent(
  assistantText: string,
  responseFormat: OpenAiResponseFormat | null,
): string {
  if (!responseFormat) {
    return assistantText;
  }
  const node = tryExtractJsonNode(assistantText);
  if (node === null) {
    return assistantText;
  }
  if (responseFormat.type === ResponseFormatTypes.JsonObject && !isJsonObject(node)) {
    return assistantText;
  }
  return JSON.stringify(node);
}

function tryExtractJsonNode(rawText: string): JsonValue | null {
  for (const candidate of enumerateJsonCandidates(rawText)) {
    const node = tryParseJsonNode(candidate);
    if (node !== null) {
      return node;
    }
  }
  return null;
}

function tryParseJsonNode(rawText: string): JsonValue | null {
  if (!rawText.trim()) {
    return null;
  }
  try {
    return JSON.parse(rawText) as JsonValue;
  } catch {
    return null;
  }
}

function* enumerateJsonCandidates(rawText: string): Iterable<string> {
  if (!rawText.trim()) {
    return;
  }
  yield rawText.trim();

  let cursor = 0;
  while (cursor < rawText.length) {
    const fenceStart = rawText.indexOf("```", cursor);
    if (fenceStart < 0) {
      return;
    }
    const bodyStart = rawText.indexOf("\n", fenceStart + 3);
    if (bodyStart < 0) {
      return;
    }
    const fenceEnd = rawText.indexOf("```", bodyStart + 1);
    if (fenceEnd < 0) {
      return;
    }
    const body = rawText.slice(bodyStart + 1, fenceEnd).trim();
    if (body) {
      yield body;
    }
    cursor = fenceEnd + 3;
  }
}

function pickString(...values: Array<JsonValue | undefined>): string | null {
  for (const value of values) {
    if (typeof value === "string" && value.trim()) {
      return value.trim();
    }
  }
  return null;
}

export function extractCopilotAssistantText(
  conversationJson: JsonObject | null,
  promptText: string,
): string | null {
  const messages = conversationJson?.messages;
  if (!Array.isArray(messages) || messages.length === 0) {
    return null;
  }

  let fallback: string | null = null;
  for (const item of messages) {
    if (!isJsonObject(item) || typeof item.text !== "string" || !item.text.trim()) {
      continue;
    }
    fallback = item.text;
  }

  const prompt = promptText.trim();
  for (let i = messages.length - 1; i >= 0; i--) {
    const item = messages[i];
    if (!isJsonObject(item) || typeof item.text !== "string" || !item.text.trim()) {
      continue;
    }
    if (item.text.trim() !== prompt) {
      return item.text;
    }
  }
  return fallback;
}

export function extractCopilotAssistantTextFromStreamData(
  streamData: string,
  promptText: string,
): string | null {
  return extractCopilotAssistantText(tryParseJsonObject(streamData), promptText);
}

export function extractCopilotConversationIdFromStream(streamData: string): string | null {
  const json = tryParseJsonObject(streamData);
  if (!json || typeof json.id !== "string" || !json.id.trim()) {
    return null;
  }
  return json.id.trim();
}

export { computeTrailingDelta };

export function buildToolCallsDelta(toolCalls: OpenAiAssistantToolCall[]): JsonValue[] {
  return toolCalls.map((toolCall, index) => ({
    index,
    id: toolCall.id,
    type: "function",
    function: {
      name: toolCall.name,
      arguments: toolCall.argumentsJson,
    },
  }));
}

export function cloneJsonObject(value: JsonObject): JsonObject {
  return cloneJsonValue(value);
}

