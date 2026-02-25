import { randomUUID } from "node:crypto";
import {
  OpenAiTransformModes,
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
  return (
    request.transformMode === OpenAiTransformModes.Simulated ||
    request.tooling.tools.length > 0 ||
    request.responseFormat !== null
  );
}

export function buildAssistantResponse(
  request: ParsedOpenAiRequest,
  assistantText: string,
): OpenAiAssistantResponse {
  const normalizedText = assistantText ?? "";
  const shouldTryToolCalls =
    request.tooling.tools.length > 0 &&
    request.tooling.toolChoiceMode !== ToolChoiceModes.None;
  if (shouldTryToolCalls) {
    const extraction = tryExtractToolCalls(normalizedText, request.tooling);
    if (extraction.length > 0) {
      return {
        content: null,
        toolCalls: extraction,
        finishReason: "tool_calls",
        strictToolErrorMessage: null,
      };
    }
    if (requiresStrictToolOutput(request.tooling)) {
      return {
        content: null,
        toolCalls: [],
        finishReason: "stop",
        strictToolErrorMessage: buildStrictToolErrorMessage(request.tooling),
      };
    }
  }

  return {
    content: normalizeStructuredContent(normalizedText, request.responseFormat),
    toolCalls: [],
    finishReason: "stop",
    strictToolErrorMessage: null,
  };
}

export function tryExtractSimulatedResponsePayload(
  assistantText: string,
  endpoint: "chat.completions" | "responses",
): JsonObject | null {
  if (!assistantText.trim()) {
    return null;
  }

  const candidates: JsonObject[] = [];
  for (const rawCandidate of enumerateJsonCandidates(assistantText)) {
    const parsed = tryParseJsonNode(rawCandidate);
    if (isJsonObject(parsed)) {
      candidates.push(parsed);
    }
  }
  if (candidates.length === 0) {
    return null;
  }

  let best: JsonObject | null = null;
  let bestScore = Number.NEGATIVE_INFINITY;
  for (const candidate of candidates) {
    const score = scoreSimulatedResponseCandidate(candidate, endpoint);
    if (score > bestScore) {
      bestScore = score;
      best = candidate;
    }
  }
  if (!best || bestScore <= 0) {
    return null;
  }
  return best;
}

function scoreSimulatedResponseCandidate(
  candidate: JsonObject,
  endpoint: "chat.completions" | "responses",
): number {
  let score = 0;
  if (isRequestLikeSimulatedPayload(candidate)) {
    score -= 180;
  }

  if (endpoint === "chat.completions") {
    const choices = candidate.choices;
    if (Array.isArray(choices) && choices.length > 0) {
      score += 220;
      const first = choices[0];
      if (isJsonObject(first)) {
        const message = first.message;
        if (isJsonObject(message)) {
          score += 80;
          if (pickString(message.role)?.toLowerCase() === "assistant") {
            score += 20;
          }
          const toolCalls = message.tool_calls;
          if (Array.isArray(toolCalls) && toolCalls.length > 0) {
            score += 90;
          }
          const content = message.content;
          if (typeof content === "string" && content.trim()) {
            score += 35;
          } else if (Array.isArray(content) && content.length > 0) {
            score += 20;
          }
        }
        if (pickString(first.finish_reason)) {
          score += 15;
        }
      }
    }

    if (looksLikeChatChoiceObject(candidate)) {
      score += 75;
    }
    if ((pickString(candidate.object) ?? "").toLowerCase() === "chat.completion") {
      score += 70;
    }
    if ((pickString(candidate.id) ?? "").toLowerCase().startsWith("chatcmpl")) {
      score += 50;
    }
    if (Array.isArray(candidate.output)) {
      score += 25;
    }
  } else {
    if ((pickString(candidate.object) ?? "").toLowerCase() === "response") {
      score += 120;
    }
    if (Array.isArray(candidate.output) && candidate.output.length > 0) {
      score += 160;
    }
    if (pickString(candidate.status)) {
      score += 25;
    }
    if (typeof candidate.output_text === "string" && candidate.output_text.trim()) {
      score += 40;
    }
    if (Array.isArray(candidate.choices)) {
      score -= 20;
    }
  }

  return score;
}

function isRequestLikeSimulatedPayload(candidate: JsonObject): boolean {
  const hasMessagesArray = Array.isArray(candidate.messages);
  const hasInputField = candidate.input !== undefined;
  const hasTools = Array.isArray(candidate.tools);
  const hasToolChoice =
    typeof candidate.tool_choice === "string" || isJsonObject(candidate.tool_choice);
  const hasParallelFlag = candidate.parallel_tool_calls !== undefined;
  const lacksResponseShape =
    !Array.isArray(candidate.choices) && !Array.isArray(candidate.output);

  return (
    (hasMessagesArray || hasInputField) &&
    (hasTools || hasToolChoice || hasParallelFlag || lacksResponseShape)
  );
}

function looksLikeChatChoiceObject(candidate: JsonObject): boolean {
  return (
    isJsonObject(candidate.message) ||
    isJsonObject(candidate.delta) ||
    candidate.finish_reason !== undefined
  );
}

export function tryBuildAssistantResponseFromChatCompletionPayload(
  payload: JsonObject,
): OpenAiAssistantResponse | null {
  const choices = payload.choices;
  if (!Array.isArray(choices) || choices.length === 0) {
    return null;
  }

  const first = choices[0];
  if (!isJsonObject(first)) {
    return null;
  }

  const finishReason = pickString(first.finish_reason) ?? "stop";
  const message = isJsonObject(first.message) ? first.message : null;
  if (!message) {
    return {
      content: "",
      toolCalls: [],
      finishReason,
      strictToolErrorMessage: null,
    };
  }

  const toolCallsNode = message.tool_calls;
  if (Array.isArray(toolCallsNode) && toolCallsNode.length > 0) {
    const toolCalls: OpenAiAssistantToolCall[] = [];
    for (const toolCallNode of toolCallsNode) {
      if (!isJsonObject(toolCallNode)) {
        continue;
      }
      const functionNode = isJsonObject(toolCallNode.function)
        ? toolCallNode.function
        : null;
      const name = pickString(toolCallNode.name, functionNode?.name);
      if (!name) {
        continue;
      }
      toolCalls.push({
        id:
          pickString(toolCallNode.id) ??
          `call_${randomUUID().replaceAll("-", "")}`,
        name,
        argumentsJson: normalizeArgumentsJson(
          toolCallNode.arguments ?? functionNode?.arguments ?? null,
        ),
      });
    }

    if (toolCalls.length > 0) {
      return {
        content: null,
        toolCalls,
        finishReason: "tool_calls",
        strictToolErrorMessage: null,
      };
    }
  }

  return {
    content: normalizeMessageContent(message.content),
    toolCalls: [],
    finishReason,
    strictToolErrorMessage: null,
  };
}

function requiresStrictToolOutput(tooling: OpenAiTooling): boolean {
  return (
    tooling.toolChoiceMode === ToolChoiceModes.Required ||
    tooling.toolChoiceMode === ToolChoiceModes.Function
  );
}

function buildStrictToolErrorMessage(tooling: OpenAiTooling): string {
  if (
    tooling.toolChoiceMode === ToolChoiceModes.Function &&
    tooling.toolChoiceFunctionName
  ) {
    return `Tool output was required, but no valid tool call for '${tooling.toolChoiceFunctionName}' was found in assistant output JSON.`;
  }
  return "Tool output was required, but no valid tool_calls JSON payload was found in assistant output.";
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

  const wrappedMessage = node.message;
  if (isJsonObject(wrappedMessage) && Array.isArray(wrappedMessage.tool_calls)) {
    return extractToolCallsFromArray(wrappedMessage.tool_calls, tooling);
  }

  const wrappedFromChoices = extractToolCallsFromChatCompletionNode(node, tooling);
  if (wrappedFromChoices.length > 0) {
    return wrappedFromChoices;
  }

  const wrappedFromOutput = extractToolCallsFromResponsesNode(node, tooling);
  if (wrappedFromOutput.length > 0) {
    return wrappedFromOutput;
  }

  const singleCall = tryBuildToolCall(node, tooling);
  return singleCall ? [singleCall] : [];
}

function extractToolCallsFromChatCompletionNode(
  node: JsonObject,
  tooling: OpenAiTooling,
): OpenAiAssistantToolCall[] {
  const choices = node.choices;
  if (!Array.isArray(choices)) {
    return [];
  }
  for (const choice of choices) {
    if (!isJsonObject(choice)) {
      continue;
    }
    const message = choice.message;
    if (isJsonObject(message) && Array.isArray(message.tool_calls)) {
      const calls = extractToolCallsFromArray(message.tool_calls, tooling);
      if (calls.length > 0) {
        return calls;
      }
    }
    const delta = choice.delta;
    if (isJsonObject(delta) && Array.isArray(delta.tool_calls)) {
      const calls = extractToolCallsFromArray(delta.tool_calls, tooling);
      if (calls.length > 0) {
        return calls;
      }
    }
  }
  return [];
}

function extractToolCallsFromResponsesNode(
  node: JsonObject,
  tooling: OpenAiTooling,
): OpenAiAssistantToolCall[] {
  const output = node.output;
  if (!Array.isArray(output)) {
    return [];
  }

  const toolCalls: OpenAiAssistantToolCall[] = [];
  for (const item of output) {
    if (!isJsonObject(item)) {
      continue;
    }
    const type = (pickString(item.type) ?? "").toLowerCase();
    if (type !== "function_call") {
      continue;
    }
    const toolCall = tryBuildToolCall(
      {
        id: pickString(item.call_id, item.tool_call_id, item.id) ?? null,
        name: pickString(item.name) ?? null,
        arguments: item.arguments ?? null,
      },
      tooling,
    );
    if (toolCall) {
      toolCalls.push(toolCall);
    }
  }
  return toolCalls;
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

function normalizeMessageContent(contentNode: JsonValue | undefined): string {
  if (typeof contentNode === "string") {
    return contentNode;
  }
  if (!Array.isArray(contentNode)) {
    return "";
  }

  const textParts: string[] = [];
  for (const item of contentNode) {
    if (typeof item === "string" && item.trim()) {
      textParts.push(item.trim());
      continue;
    }
    if (!isJsonObject(item)) {
      continue;
    }
    const type = (pickString(item.type) ?? "").toLowerCase();
    if (
      (type === "" || type === "text" || type === "output_text") &&
      pickString(item.text)
    ) {
      textParts.push(pickString(item.text) ?? "");
    }
  }
  return textParts.join("\n");
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
  const trimmed = rawText.trim();
  if (!trimmed) {
    return;
  }
  const seen = new Set<string>();
  const yieldCandidate = function* (candidate: string): Iterable<string> {
    const normalized = candidate.trim();
    if (!normalized || seen.has(normalized)) {
      return;
    }
    seen.add(normalized);
    yield normalized;
  };

  yield* yieldCandidate(trimmed);

  let cursor = 0;
  while (cursor < rawText.length) {
    const fenceStart = rawText.indexOf("```", cursor);
    if (fenceStart < 0) {
      break;
    }
    const bodyStart = rawText.indexOf("\n", fenceStart + 3);
    if (bodyStart < 0) {
      break;
    }
    const fenceEnd = rawText.indexOf("```", bodyStart + 1);
    if (fenceEnd < 0) {
      break;
    }
    const body = rawText.slice(bodyStart + 1, fenceEnd).trim();
    if (body) {
      yield* yieldCandidate(body);
    }
    cursor = fenceEnd + 3;
  }

  for (const balanced of extractBalancedJsonSegments(rawText)) {
    yield* yieldCandidate(balanced);
  }
}

function* extractBalancedJsonSegments(rawText: string): Iterable<string> {
  const maxCandidates = 128;
  let emitted = 0;
  for (let start = 0; start < rawText.length; start++) {
    if (emitted >= maxCandidates) {
      break;
    }
    const opening = rawText[start];
    if (opening !== "{" && opening !== "[") {
      continue;
    }
    const balanced = extractBalancedJsonSegment(rawText, start, opening);
    if (!balanced) {
      continue;
    }
    emitted++;
    yield balanced;
  }
}

function extractBalancedJsonSegment(
  rawText: string,
  start: number,
  opening: string,
): string | null {
  const closing = opening === "{" ? "}" : "]";
  let depth = 0;
  let inString = false;
  let escaped = false;

  for (let index = start; index < rawText.length; index++) {
    const ch = rawText[index];
    if (!ch) {
      continue;
    }
    if (inString) {
      if (escaped) {
        escaped = false;
        continue;
      }
      if (ch === "\\") {
        escaped = true;
        continue;
      }
      if (ch === "\"") {
        inString = false;
      }
      continue;
    }

    if (ch === "\"") {
      inString = true;
      continue;
    }
    if (ch === opening) {
      depth++;
      continue;
    }
    if (ch === closing) {
      depth--;
      if (depth === 0) {
        return rawText.slice(start, index + 1).trim();
      }
    }
  }
  return null;
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
