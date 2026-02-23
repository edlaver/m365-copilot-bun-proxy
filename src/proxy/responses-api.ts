import { randomUUID } from "node:crypto";
import type {
  JsonObject,
  OpenAiAssistantResponse,
  OpenAiAssistantToolCall,
  ParsedResponsesRequest,
} from "./types";
import { cloneJsonValue } from "./utils";

export function createOpenAiResponseId(): string {
  return `resp_${randomUUID().replaceAll("-", "")}`;
}

export function createOpenAiOutputItemId(prefix: string): string {
  return `${prefix}_${randomUUID().replaceAll("-", "")}`;
}

export function buildOpenAiResponseFromAssistant(
  responseId: string,
  createdAt: number,
  model: string,
  status: string,
  parsedRequest: ParsedResponsesRequest,
  assistantResponse: OpenAiAssistantResponse,
  includeConversationId: boolean,
  conversationId: string,
): JsonObject {
  const output =
    assistantResponse.toolCalls.length > 0
      ? buildFunctionCallOutputItems(assistantResponse.toolCalls, "completed")
      : [
          buildMessageOutputItem(
            createOpenAiOutputItemId("msg"),
            assistantResponse.content ?? "",
            "completed",
          ),
        ];

  return buildOpenAiResponseObject(
    responseId,
    createdAt,
    model,
    status,
    output,
    parsedRequest,
    includeConversationId ? conversationId : null,
  );
}

export function buildOpenAiResponseObject(
  responseId: string,
  createdAt: number,
  model: string,
  status: string,
  output: JsonObject[],
  parsedRequest: ParsedResponsesRequest,
  conversationId: string | null,
): JsonObject {
  const response: JsonObject = {
    id: responseId,
    object: "response",
    created_at: createdAt,
    status,
    model,
    output: output.map((item) => cloneJsonValue(item)),
    output_text: extractOutputText(output),
    parallel_tool_calls: parsedRequest.base.tooling.parallelToolCalls,
  };

  if (parsedRequest.previousResponseId) {
    response.previous_response_id = parsedRequest.previousResponseId;
  }
  if (parsedRequest.instructions) {
    response.instructions = parsedRequest.instructions;
  }
  if (parsedRequest.inputItemsForStorage.length > 0) {
    response.input = cloneJsonValue(parsedRequest.inputItemsForStorage);
  }
  if (conversationId?.trim()) {
    response.conversation_id = conversationId;
  }

  return response;
}

export function buildMessageOutputItem(
  itemId: string,
  text: string,
  status: string,
): JsonObject {
  return {
    id: itemId,
    type: "message",
    status,
    role: "assistant",
    content: [{ type: "output_text", text }],
  };
}

export function buildFunctionCallOutputItems(
  toolCalls: OpenAiAssistantToolCall[],
  status: string,
): JsonObject[] {
  return toolCalls.map((toolCall) =>
    buildFunctionCallOutputItem(
      createOpenAiOutputItemId("fc"),
      toolCall,
      status,
    ),
  );
}

export function buildFunctionCallOutputItem(
  itemId: string,
  toolCall: OpenAiAssistantToolCall,
  status: string,
): JsonObject {
  return {
    id: itemId,
    type: "function_call",
    status,
    call_id: toolCall.id,
    name: toolCall.name,
    arguments: toolCall.argumentsJson,
  };
}

export function buildResponseCreatedEvent(response: JsonObject): JsonObject {
  return { type: "response.created", response: cloneJsonValue(response) };
}

export function buildResponseInProgressEvent(response: JsonObject): JsonObject {
  return { type: "response.in_progress", response: cloneJsonValue(response) };
}

export function buildResponseOutputItemAddedEvent(
  responseId: string,
  outputIndex: number,
  item: JsonObject,
): JsonObject {
  return {
    type: "response.output_item.added",
    response_id: responseId,
    output_index: outputIndex,
    item: cloneJsonValue(item),
  };
}

export function buildResponseOutputTextDeltaEvent(
  responseId: string,
  outputIndex: number,
  itemId: string,
  delta: string,
): JsonObject {
  return {
    type: "response.output_text.delta",
    response_id: responseId,
    output_index: outputIndex,
    item_id: itemId,
    delta,
  };
}

export function buildResponseOutputTextDoneEvent(
  responseId: string,
  outputIndex: number,
  itemId: string,
  text: string,
): JsonObject {
  return {
    type: "response.output_text.done",
    response_id: responseId,
    output_index: outputIndex,
    item_id: itemId,
    text,
  };
}

export function buildResponseOutputItemDoneEvent(
  responseId: string,
  outputIndex: number,
  item: JsonObject,
): JsonObject {
  return {
    type: "response.output_item.done",
    response_id: responseId,
    output_index: outputIndex,
    item: cloneJsonValue(item),
  };
}

export function buildResponseCompletedEvent(response: JsonObject): JsonObject {
  return { type: "response.completed", response: cloneJsonValue(response) };
}

function extractOutputText(output: JsonObject[]): string {
  const segments: string[] = [];
  for (const item of output) {
    if ((item.type ?? "") !== "message") {
      continue;
    }
    const content = item.content;
    if (!Array.isArray(content)) {
      continue;
    }
    for (const contentItem of content) {
      if (
        !contentItem ||
        typeof contentItem !== "object" ||
        Array.isArray(contentItem)
      ) {
        continue;
      }
      const typed = contentItem as Record<string, unknown>;
      if ((typed.type ?? "") !== "output_text") {
        continue;
      }
      const text = typed.text;
      if (typeof text === "string" && text.length > 0) {
        segments.push(text);
      }
    }
  }
  return segments.join("");
}
