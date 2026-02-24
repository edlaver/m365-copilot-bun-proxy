import {
  ResponseFormatTypes,
  ToolChoiceModes,
  TransportNames,
  type ContextMessage,
  type JsonObject,
  type JsonValue,
  type OpenAiResponseFormat,
  type OpenAiToolDefinition,
  type OpenAiTooling,
  type ParsedOpenAiRequest,
  type ParsedResponsesRequest,
  type WrapperOptions,
} from "./types";
import {
  firstNonEmpty,
  normalizeNullableString,
  parseBooleanString,
  tryGetBoolean,
  tryGetDouble,
  tryGetString,
  isJsonObject,
  cloneJsonValue,
} from "./utils";

export function normalizeTransport(
  transport: string | null | undefined,
): string {
  if (!transport || !transport.trim()) {
    return TransportNames.Graph;
  }
  return transport.trim().toLowerCase();
}

export function isSupportedTransport(
  transport: string | null | undefined,
): boolean {
  const normalized = normalizeTransport(transport);
  return (
    normalized === TransportNames.Graph ||
    normalized === TransportNames.Substrate
  );
}

export function resolveTransport(
  request: Request,
  requestJson: JsonObject,
  options: WrapperOptions,
): string {
  const transport = firstNonEmpty(
    request.headers.get("x-m365-transport"),
    tryGetString(requestJson, "m365_transport"),
    tryGetString(requestJson, "transport"),
    options.transport,
  );
  return normalizeTransport(transport);
}

export function selectConversation(
  request: Request,
  requestJson: JsonObject,
  fallbackConversationKey: string | null,
): {
  conversationId: string | null;
  conversationKey: string | null;
  forceNewConversation: boolean;
} {
  const conversationId = normalizeNullableString(
    firstNonEmpty(
      request.headers.get("x-m365-conversation-id"),
      tryGetString(requestJson, "m365_conversation_id"),
      tryGetString(requestJson, "conversation_id"),
    ),
  );

  const conversationKey = normalizeNullableString(
    firstNonEmpty(
      request.headers.get("x-m365-conversation-key"),
      tryGetString(requestJson, "m365_conversation_key"),
      fallbackConversationKey,
    ),
  );

  const forceNewFromHeader = parseBooleanString(
    request.headers.get("x-m365-new-conversation"),
  );
  const forceNewFromBody =
    tryGetBoolean(requestJson, "m365_new_conversation") === true;

  return {
    conversationId,
    conversationKey,
    forceNewConversation: forceNewFromHeader || forceNewFromBody,
  };
}

export function scopeConversationKey(
  conversationKey: string | null,
  transportName: string,
): string | null {
  const normalizedKey = normalizeNullableString(conversationKey);
  if (!normalizedKey) {
    return null;
  }
  return `${normalizeTransport(transportName)}:${normalizedKey}`;
}

export function buildCopilotRequestPayload(
  request: ParsedOpenAiRequest,
): JsonObject {
  const payload: JsonObject = {
    message: { text: request.promptText },
    locationHint: cloneJsonValue(request.locationHint),
  };

  if (request.additionalContext.length > 0) {
    const additionalContext: JsonValue[] = [];
    for (const item of request.additionalContext) {
      if (!item.text.trim()) {
        continue;
      }
      const context: JsonObject = { text: item.text };
      if (item.description?.trim()) {
        context.description = item.description;
      }
      additionalContext.push(context);
    }
    if (additionalContext.length > 0) {
      payload.additionalContext = additionalContext;
    }
  }

  if (request.contextualResources) {
    payload.contextualResources = cloneJsonValue(request.contextualResources);
  }
  return payload;
}

export function tryParseOpenAiRequest(
  requestJson: JsonObject,
  options: WrapperOptions,
): { ok: true; request: ParsedOpenAiRequest } | { ok: false; error: string } {
  const messagesNode = requestJson.messages;
  if (!Array.isArray(messagesNode) || messagesNode.length === 0) {
    return {
      ok: false,
      error: "The 'messages' array is required and cannot be empty.",
    };
  }

  const tooling = parseTooling(requestJson);
  const responseFormat = parseResponseFormat(requestJson);
  const reasoningEffort = tryGetString(requestJson, "reasoning_effort");
  const temperature = tryGetDouble(requestJson, "temperature");

  const messages: { role: string; content: string; index: number }[] = [];
  for (let index = 0; index < messagesNode.length; index++) {
    const message = messagesNode[index];
    if (!isJsonObject(message)) {
      continue;
    }

    const role = (tryGetString(message, "role") ?? "user").toLowerCase();
    let content = extractMessageContent(message.content, role);

    if (role === "assistant") {
      const assistantToolCalls = tryExtractAssistantToolCalls(message, content);
      if (assistantToolCalls.length > 0) {
        content = convertToolCallsToContextText(assistantToolCalls);
      }
    }

    if (!content && role === "tool") {
      const toolName = tryGetString(message, "name");
      const toolCallId = tryGetString(message, "tool_call_id");
      let prefix = toolName ? `tool:${toolName}` : "tool";
      if (toolCallId) {
        prefix += `[${toolCallId}]`;
      }
      const toolPayload = stringifyJsonValue(message.content);
      if (toolPayload) {
        content = `${prefix}: ${toolPayload}`;
      }
    }

    if (!content.trim()) {
      continue;
    }
    messages.push({ role, content: content.trim(), index });
  }

  if (messages.length === 0) {
    return {
      ok: false,
      error: "No textual content could be extracted from 'messages'.",
    };
  }

  const prompt = resolvePrompt(messages);
  if (!prompt?.content?.trim()) {
    return {
      ok: false,
      error: "Unable to determine a prompt from the message list.",
    };
  }

  const stream = tryGetBoolean(requestJson, "stream") === true;
  const model =
    tryGetString(requestJson, "model") ||
    (options.defaultModel?.trim() ? options.defaultModel : "m365-copilot");

  const parsedRequest: ParsedOpenAiRequest = {
    model,
    stream,
    promptText: prompt.content,
    userKey: tryGetString(requestJson, "user"),
    locationHint: buildLocationHint(requestJson, options.defaultTimeZone),
    contextualResources: buildContextualResources(requestJson),
    additionalContext: buildAdditionalContext(
      messages,
      prompt.index,
      requestJson,
      options.maxAdditionalContextMessages,
      tooling,
      responseFormat,
      reasoningEffort,
      temperature,
    ),
    tooling,
    responseFormat,
    reasoningEffort,
    temperature,
  };

  return { ok: true, request: parsedRequest };
}

export function tryParseResponsesRequest(
  requestJson: JsonObject,
  options: WrapperOptions,
):
  | { ok: true; request: ParsedResponsesRequest }
  | { ok: false; error: string } {
  const normalized = cloneJsonValue(requestJson);
  const normalizedInput = normalizeResponsesInput(requestJson.input);
  if (normalizedInput.messages.length === 0) {
    return {
      ok: false,
      error:
        "The 'input' field is required and must contain at least one supported message item.",
    };
  }

  normalized.messages = normalizedInput.messages;

  if (normalized.response_format === undefined) {
    const mappedResponseFormat = mapResponsesTextFormat(requestJson);
    if (mappedResponseFormat) {
      normalized.response_format = mappedResponseFormat;
    }
  }

  if (normalized.reasoning_effort === undefined) {
    const reasoningEffort = mapResponsesReasoningEffort(requestJson);
    if (reasoningEffort) {
      normalized.reasoning_effort = reasoningEffort;
    }
  }

  const instructions = tryGetString(requestJson, "instructions");
  if (instructions && !tryGetString(normalized, "m365_system_prompt")) {
    normalized.m365_system_prompt = instructions;
  }

  const parsedBase = tryParseOpenAiRequest(normalized, options);
  if (!parsedBase.ok) {
    return parsedBase;
  }

  return {
    ok: true,
    request: {
      base: parsedBase.request,
      previousResponseId: tryGetString(requestJson, "previous_response_id"),
      inputItemsForStorage: normalizedInput.inputItemsForStorage,
      instructions,
    },
  };
}

function resolvePrompt(
  messages: { role: string; content: string; index: number }[],
): { role: string; content: string; index: number } | null {
  for (let i = messages.length - 1; i >= 0; i--) {
    if (messages[i].role === "user") {
      return messages[i];
    }
  }
  return messages.at(-1) ?? null;
}

function normalizeResponsesInput(inputNode: JsonValue | undefined): {
  messages: JsonObject[];
  inputItemsForStorage: JsonValue[];
} {
  if (typeof inputNode === "string" && inputNode.trim()) {
    return {
      messages: [{ role: "user", content: inputNode.trim() }],
      inputItemsForStorage: [
        {
          type: "message",
          role: "user",
          content: [{ type: "input_text", text: inputNode.trim() }],
        },
      ],
    };
  }

  const sourceItems: JsonValue[] = Array.isArray(inputNode)
    ? inputNode
    : inputNode !== undefined && inputNode !== null
      ? [inputNode]
      : [];
  const messages: JsonObject[] = [];
  const inputItemsForStorage: JsonValue[] = [];

  for (const item of sourceItems) {
    if (typeof item === "string") {
      if (!item.trim()) {
        continue;
      }
      messages.push({ role: "user", content: item.trim() });
      inputItemsForStorage.push({
        type: "message",
        role: "user",
        content: [{ type: "input_text", text: item.trim() }],
      });
      continue;
    }

    if (!isJsonObject(item)) {
      continue;
    }

    inputItemsForStorage.push(cloneJsonValue(item));
    const type = (tryGetString(item, "type") ?? "message").toLowerCase();
    if (type === "message") {
      const role = (tryGetString(item, "role") ?? "user").toLowerCase();
      const message: JsonObject = { role };
      if (item.content !== undefined) {
        message.content = cloneJsonValue(item.content);
      } else {
        const text =
          tryGetString(item, "text") ??
          tryGetString(item, "input_text") ??
          tryGetString(item, "output_text");
        if (text) {
          message.content = text;
        }
      }
      messages.push(message);
      continue;
    }

    if (type === "function_call_output") {
      const message: JsonObject = {
        role: "tool",
        content: stringifyJsonValue(item.output),
      };
      const toolName = tryGetString(item, "name");
      if (toolName) {
        message.name = toolName;
      }
      const toolCallId =
        tryGetString(item, "call_id") ?? tryGetString(item, "tool_call_id");
      if (toolCallId) {
        message.tool_call_id = toolCallId;
      }
      messages.push(message);
      continue;
    }

    if (type === "function_call") {
      const functionName = tryGetString(item, "name");
      if (!functionName) {
        continue;
      }
      const functionArguments =
        normalizeFunctionArguments(item.arguments) ?? "{}";
      const toolCall: JsonObject = {
        id:
          tryGetString(item, "call_id") ??
          tryGetString(item, "id") ??
          `call_${messages.length + 1}`,
        type: "function",
        function: {
          name: functionName,
          arguments: functionArguments,
        },
      };
      messages.push({
        role: "assistant",
        content: null,
        tool_calls: [toolCall],
      });
    }
  }

  return { messages, inputItemsForStorage };
}

function mapResponsesTextFormat(requestJson: JsonObject): JsonObject | null {
  const text = requestJson.text;
  if (!isJsonObject(text) || !isJsonObject(text.format)) {
    return null;
  }
  const format = text.format;
  const type = tryGetString(format, "type");
  if (!type) {
    return null;
  }
  const normalizedType = type.toLowerCase();
  if (normalizedType === ResponseFormatTypes.JsonObject) {
    return { type: ResponseFormatTypes.JsonObject };
  }
  if (normalizedType !== ResponseFormatTypes.JsonSchema) {
    return null;
  }

  if (isJsonObject(format.json_schema)) {
    return {
      type: ResponseFormatTypes.JsonSchema,
      json_schema: cloneJsonValue(format.json_schema),
    };
  }

  const jsonSchema: JsonObject = {};
  const name = tryGetString(format, "name");
  if (name) {
    jsonSchema.name = name;
  }
  if (isJsonObject(format.schema)) {
    jsonSchema.schema = cloneJsonValue(format.schema);
  }

  return {
    type: ResponseFormatTypes.JsonSchema,
    json_schema: jsonSchema,
  };
}

function mapResponsesReasoningEffort(requestJson: JsonObject): string | null {
  const reasoning = requestJson.reasoning;
  if (!isJsonObject(reasoning)) {
    return null;
  }
  return tryGetString(reasoning, "effort");
}

function stringifyJsonValue(value: JsonValue | undefined): string {
  if (typeof value === "string") {
    const normalized = normalizeJsonLikeString(value);
    return normalized ?? value;
  }
  if (value === undefined || value === null) {
    return "";
  }
  if (typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }
  return JSON.stringify(value);
}

function normalizeFunctionArguments(value: JsonValue | undefined): string | null {
  if (value === undefined || value === null) {
    return null;
  }
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (!trimmed) {
      return "{}";
    }
    const parsed = tryParseJsonValueFromText(trimmed);
    return parsed !== null ? JSON.stringify(parsed) : trimmed;
  }
  return JSON.stringify(value);
}

function extractMessageContent(
  contentNode: JsonValue | undefined,
  role: string,
): string {
  if (contentNode === undefined || contentNode === null) {
    return "";
  }
  if (typeof contentNode === "string") {
    const normalized = normalizeJsonLikeString(contentNode);
    return normalized ?? contentNode;
  }
  if (isJsonObject(contentNode)) {
    const directText =
      tryGetString(contentNode, "text") ?? tryGetString(contentNode, "value");
    if (directText) {
      const normalized = normalizeJsonLikeString(directText);
      return normalized ?? directText;
    }
    const imageFromObject = extractImageReference(contentNode);
    if (imageFromObject) {
      return `[${role} attached image: ${imageFromObject}]`;
    }
    return "";
  }
  if (!Array.isArray(contentNode)) {
    return "";
  }

  const textParts: string[] = [];
  const imageParts: string[] = [];
  for (const part of contentNode) {
    if (typeof part === "string" && part.trim()) {
      const normalized = normalizeJsonLikeString(part.trim());
      textParts.push(normalized ?? part.trim());
      continue;
    }
    if (!isJsonObject(part)) {
      continue;
    }
    const type = tryGetString(part, "type");
    const isTextPart =
      !type ||
      type.toLowerCase() === "text" ||
      type.toLowerCase() === "input_text";

    if (isTextPart) {
      let partText = tryGetString(part, "text");
      const nestedText = part.text;
      if (!partText && isJsonObject(nestedText)) {
        partText = tryGetString(nestedText, "value");
      }
      if (partText) {
        const normalized = normalizeJsonLikeString(partText.trim());
        textParts.push(normalized ?? partText.trim());
        continue;
      }
    }

    if (
      type?.toLowerCase() === "input_image" ||
      type?.toLowerCase() === "image_url" ||
      type?.toLowerCase() === "image"
    ) {
      imageParts.push(extractImageReference(part) ?? "(inline-image)");
    }
  }

  if (textParts.length === 0 && imageParts.length === 0) {
    return "";
  }

  const imageSuffix =
    imageParts.length === 0
      ? ""
      : `\n[${role} attached ${imageParts.length} image input(s): ${imageParts
          .slice(0, 4)
          .join(", ")}]`;

  if (textParts.length === 0) {
    return imageSuffix.trimStart();
  }
  return `${textParts.join("\n")}${imageSuffix}`;
}

function tryExtractAssistantToolCalls(
  message: JsonObject,
  extractedContent: string,
): JsonObject[] {
  if (Array.isArray(message.tool_calls)) {
    return message.tool_calls.filter(isJsonObject);
  }
  if (!extractedContent.trim()) {
    return [];
  }

  const parsedContent = tryParseJsonValueFromText(extractedContent);
  if (parsedContent === null) {
    return [];
  }
  return extractToolCallsFromJsonNode(parsedContent);
}

function extractToolCallsFromJsonNode(node: JsonValue): JsonObject[] {
  if (Array.isArray(node)) {
    return node.filter(isJsonObject).filter(isToolCallLikeObject);
  }
  if (!isJsonObject(node)) {
    return [];
  }

  if (Array.isArray(node.tool_calls)) {
    return node.tool_calls.filter(isJsonObject).filter(isToolCallLikeObject);
  }

  const wrappedFromMessage = node.message;
  if (isJsonObject(wrappedFromMessage) && Array.isArray(wrappedFromMessage.tool_calls)) {
    return wrappedFromMessage.tool_calls
      .filter(isJsonObject)
      .filter(isToolCallLikeObject);
  }

  const wrappedFromChoices = extractToolCallsFromChatCompletionNode(node);
  if (wrappedFromChoices.length > 0) {
    return wrappedFromChoices;
  }

  const wrappedFromOutput = extractToolCallsFromResponsesOutputNode(node);
  if (wrappedFromOutput.length > 0) {
    return wrappedFromOutput;
  }

  if (tryGetString(node, "name")) {
    return [node];
  }

  const functionObject = node.function;
  if (isJsonObject(functionObject) && tryGetString(functionObject, "name")) {
    return [node];
  }

  return [];
}

function isToolCallLikeObject(value: JsonObject): boolean {
  const functionObject = isJsonObject(value.function) ? value.function : null;
  const type = (tryGetString(value, "type") ?? "").toLowerCase();
  return Boolean(
    tryGetString(value, "name") ||
      tryGetString(value, "tool_name") ||
      tryGetString(functionObject, "name") ||
      type === "function" ||
      type === "function_call",
  );
}

function extractToolCallsFromChatCompletionNode(node: JsonObject): JsonObject[] {
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
      return message.tool_calls.filter(isJsonObject);
    }
    const delta = choice.delta;
    if (isJsonObject(delta) && Array.isArray(delta.tool_calls)) {
      return delta.tool_calls.filter(isJsonObject);
    }
  }
  return [];
}

function extractToolCallsFromResponsesOutputNode(node: JsonObject): JsonObject[] {
  const output = node.output;
  if (!Array.isArray(output)) {
    return [];
  }

  const toolCalls: JsonObject[] = [];
  for (const item of output) {
    if (!isJsonObject(item)) {
      continue;
    }
    const type = (tryGetString(item, "type") ?? "").toLowerCase();
    if (type !== "function_call") {
      continue;
    }
    const name = tryGetString(item, "name");
    if (!name) {
      continue;
    }
    toolCalls.push({
      id:
        tryGetString(item, "call_id") ??
        tryGetString(item, "tool_call_id") ??
        tryGetString(item, "id") ??
        `call_${toolCalls.length + 1}`,
      type: "function",
      function: {
        name,
        arguments: normalizeFunctionArguments(item.arguments) ?? "{}",
      },
    });
  }
  return toolCalls;
}

function parseTooling(requestJson: JsonObject): OpenAiTooling {
  const tools: OpenAiToolDefinition[] = [];
  const toolsArray = requestJson.tools;
  if (Array.isArray(toolsArray)) {
    for (const toolNode of toolsArray) {
      if (!isJsonObject(toolNode)) {
        continue;
      }
      const type = tryGetString(toolNode, "type");
      if (type?.toLowerCase() !== "function") {
        continue;
      }
      const functionObject = toolNode.function;
      if (!isJsonObject(functionObject)) {
        continue;
      }
      const name = tryGetString(functionObject, "name");
      if (!name) {
        continue;
      }
      const parameters = isJsonObject(functionObject.parameters)
        ? cloneJsonValue(functionObject.parameters)
        : {};

      tools.push({
        name: name.trim(),
        description: tryGetString(functionObject, "description"),
        parameters,
      });
    }
  }

  let toolChoiceMode: string =
    tools.length === 0 ? ToolChoiceModes.None : ToolChoiceModes.Auto;
  let toolChoiceFunctionName: string | null = null;
  const toolChoice = requestJson.tool_choice;
  if (typeof toolChoice === "string") {
    const normalized = toolChoice.trim().toLowerCase();
    if (
      normalized === ToolChoiceModes.Auto ||
      normalized === ToolChoiceModes.None ||
      normalized === ToolChoiceModes.Required
    ) {
      toolChoiceMode = normalized;
    }
  } else if (isJsonObject(toolChoice)) {
    const type = tryGetString(toolChoice, "type");
    const functionObject = toolChoice.function;
    if (type?.toLowerCase() === "function" && isJsonObject(functionObject)) {
      toolChoiceFunctionName = tryGetString(functionObject, "name");
      if (toolChoiceFunctionName) {
        toolChoiceMode = ToolChoiceModes.Function;
      }
    }
  }

  return {
    tools,
    toolChoiceMode,
    toolChoiceFunctionName,
    parallelToolCalls:
      tryGetBoolean(requestJson, "parallel_tool_calls") !== false,
  };
}

function parseResponseFormat(
  requestJson: JsonObject,
): OpenAiResponseFormat | null {
  const responseFormat = requestJson.response_format;
  if (!isJsonObject(responseFormat)) {
    return null;
  }
  const type = tryGetString(responseFormat, "type");
  if (!type) {
    return null;
  }
  const normalizedType = type.toLowerCase();
  if (normalizedType === ResponseFormatTypes.JsonObject) {
    return {
      type: ResponseFormatTypes.JsonObject,
      name: null,
      jsonSchema: null,
    };
  }
  if (normalizedType !== ResponseFormatTypes.JsonSchema) {
    return null;
  }

  const schemaNode = responseFormat.json_schema;
  if (!isJsonObject(schemaNode)) {
    return {
      type: ResponseFormatTypes.JsonSchema,
      name: null,
      jsonSchema: null,
    };
  }

  return {
    type: ResponseFormatTypes.JsonSchema,
    name: tryGetString(schemaNode, "name"),
    jsonSchema: isJsonObject(schemaNode.schema)
      ? cloneJsonValue(schemaNode.schema)
      : null,
  };
}

function buildAdditionalContext(
  messages: { role: string; content: string; index: number }[],
  promptMessageIndex: number,
  requestJson: JsonObject,
  maxContextMessages: number,
  tooling: OpenAiTooling,
  responseFormat: OpenAiResponseFormat | null,
  reasoningEffort: string | null,
  temperature: number | null,
): ContextMessage[] {
  let context: ContextMessage[] = [];
  for (const message of messages) {
    if (message.index === promptMessageIndex || !message.content.trim()) {
      continue;
    }
    context.push({
      text: `${message.role}: ${message.content}`,
      description: null,
    });
  }

  appendCustomContext(context, requestJson.m365_additional_context);

  const systemPrompt = tryGetString(requestJson, "m365_system_prompt");
  if (systemPrompt) {
    context.push({ text: systemPrompt, description: "System prompt override" });
  }

  appendOpenAiCompatibilityContext(
    context,
    tooling,
    responseFormat,
    reasoningEffort,
    temperature,
  );

  if (maxContextMessages > 0 && context.length > maxContextMessages) {
    context = context.slice(context.length - maxContextMessages);
  }
  return context;
}

function appendCustomContext(
  context: ContextMessage[],
  customNode: JsonValue | undefined,
): void {
  if (customNode === undefined || customNode === null) {
    return;
  }
  if (typeof customNode === "string" && customNode.trim()) {
    context.push({ text: customNode.trim(), description: null });
    return;
  }
  if (isJsonObject(customNode)) {
    appendCustomContextObject(context, customNode);
    return;
  }
  if (!Array.isArray(customNode)) {
    return;
  }
  for (const item of customNode) {
    if (typeof item === "string" && item.trim()) {
      context.push({ text: item.trim(), description: null });
      continue;
    }
    if (isJsonObject(item)) {
      appendCustomContextObject(context, item);
    }
  }
}

function appendCustomContextObject(
  context: ContextMessage[],
  contextObject: JsonObject,
): void {
  const text =
    tryGetString(contextObject, "text") ??
    tryGetString(contextObject, "content");
  if (!text) {
    return;
  }
  context.push({
    text: text.trim(),
    description: tryGetString(contextObject, "description"),
  });
}

function appendOpenAiCompatibilityContext(
  context: ContextMessage[],
  tooling: OpenAiTooling,
  responseFormat: OpenAiResponseFormat | null,
  reasoningEffort: string | null,
  temperature: number | null,
): void {
  if (tooling.tools.length > 0) {
    context.push({
      text: 'If you call a tool, respond ONLY as minified JSON with this exact shape: {"tool_calls":[{"name":"<tool-name>","arguments":{}}]}. No markdown, no prose, no extra keys.',
      description: "OpenAI tool-calling contract",
    });
    context.push({
      text: "If no tool call is needed, return a normal assistant response (unless response_format requires JSON-only output).",
      description: "OpenAI tool-calling contract",
    });
    context.push({
      text: JSON.stringify(
        tooling.tools.map((tool) => ({
          name: tool.name,
          description: tool.description,
          parameters: cloneJsonValue(tool.parameters),
        })),
      ),
      description: "Available tools",
    });

    if (tooling.toolChoiceMode === ToolChoiceModes.None) {
      context.push({
        text: "Tool calls are disabled for this response.",
        description: "Tool choice",
      });
    } else if (tooling.toolChoiceMode === ToolChoiceModes.Required) {
      context.push({
        text: "A tool call is required in this turn. Do not return plain assistant text first.",
        description: "Tool choice",
      });
    } else if (
      tooling.toolChoiceMode === ToolChoiceModes.Function &&
      tooling.toolChoiceFunctionName
    ) {
      context.push({
        text: `You must call only tool '${tooling.toolChoiceFunctionName}' in this turn. Do not return plain assistant text first.`,
        description: "Tool choice",
      });
    }
  }

  if (responseFormat) {
    if (responseFormat.type === ResponseFormatTypes.JsonObject) {
      context.push({
        text: "Return ONLY a valid JSON object and no markdown.",
        description: "OpenAI response_format",
      });
    } else if (responseFormat.type === ResponseFormatTypes.JsonSchema) {
      context.push({
        text: "Return ONLY valid JSON that conforms to the provided JSON schema.",
        description: "OpenAI response_format",
      });
      if (responseFormat.jsonSchema) {
        context.push({
          text: JSON.stringify(responseFormat.jsonSchema),
          description: "JSON schema",
        });
      }
    }
  }

  if (reasoningEffort) {
    context.push({
      text: `Reasoning effort preference: ${reasoningEffort}.`,
      description: "Reasoning hint",
    });
  }
  if (temperature !== null) {
    context.push({
      text: `Sampling temperature preference: ${temperature}.`,
      description: "Generation hint",
    });
  }
}

function normalizeJsonLikeString(rawText: string): string | null {
  const parsed = tryParseJsonValueFromText(rawText);
  if (parsed === null) {
    return null;
  }
  return JSON.stringify(parsed);
}

function tryParseJsonValueFromText(rawText: string): JsonValue | null {
  for (const candidate of enumerateJsonCandidates(rawText)) {
    const parsed = tryParseJsonValue(candidate);
    if (parsed !== null) {
      return parsed;
    }
  }
  return null;
}

function tryParseJsonValue(rawText: string): JsonValue | null {
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

function extractImageReference(partObject: JsonObject): string | null {
  const imageUrl = partObject.image_url;
  if (typeof imageUrl === "string" && imageUrl.trim()) {
    return imageUrl.trim();
  }
  if (isJsonObject(imageUrl)) {
    const url = tryGetString(imageUrl, "url");
    if (url) {
      return url.trim();
    }
  }
  const directUrl = tryGetString(partObject, "url");
  return directUrl ? directUrl.trim() : null;
}

function convertToolCallsToContextText(toolCalls: JsonValue[]): string {
  const calls = toolCalls.filter(isJsonObject);
  return calls.length === 0
    ? ""
    : `assistant tool_calls: ${JSON.stringify(calls)}`;
}

function buildLocationHint(
  requestJson: JsonObject,
  defaultTimeZone: string,
): JsonObject {
  let locationHint: JsonObject = {};
  const explicitLocationHint = requestJson.m365_location_hint;
  if (isJsonObject(explicitLocationHint)) {
    locationHint = cloneJsonValue(explicitLocationHint);
  } else if (
    typeof explicitLocationHint === "string" &&
    explicitLocationHint.trim()
  ) {
    locationHint.timeZone = explicitLocationHint.trim();
  }

  const timeZoneOverride = tryGetString(requestJson, "m365_time_zone");
  if (timeZoneOverride) {
    locationHint.timeZone = timeZoneOverride;
  }

  if (!tryGetString(locationHint, "timeZone")) {
    locationHint.timeZone = defaultTimeZone?.trim()
      ? defaultTimeZone
      : "America/New_York";
  }

  const countryOrRegion = tryGetString(requestJson, "m365_country_or_region");
  if (countryOrRegion && !tryGetString(locationHint, "countryOrRegion")) {
    locationHint.countryOrRegion = countryOrRegion;
  }
  return locationHint;
}

function buildContextualResources(requestJson: JsonObject): JsonObject | null {
  const custom = requestJson.m365_contextual_resources;
  if (isJsonObject(custom)) {
    return cloneJsonValue(custom);
  }
  const direct = requestJson.contextualResources;
  if (isJsonObject(direct)) {
    return cloneJsonValue(direct);
  }
  return null;
}
