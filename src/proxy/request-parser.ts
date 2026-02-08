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

    if (!content && role === "assistant" && Array.isArray(message.tool_calls)) {
      content = convertToolCallsToContextText(message.tool_calls);
    }

    if (!content && role === "tool") {
      const toolName = tryGetString(message, "name");
      const toolCallId = tryGetString(message, "tool_call_id");
      let prefix = toolName ? `tool:${toolName}` : "tool";
      if (toolCallId) {
        prefix += `[${toolCallId}]`;
      }
      const toolPayload = tryGetString(message, "content");
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

function extractMessageContent(
  contentNode: JsonValue | undefined,
  role: string,
): string {
  if (contentNode === undefined || contentNode === null) {
    return "";
  }
  if (typeof contentNode === "string") {
    return contentNode;
  }
  if (isJsonObject(contentNode)) {
    const directText =
      tryGetString(contentNode, "text") ?? tryGetString(contentNode, "value");
    if (directText) {
      return directText;
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
      textParts.push(part.trim());
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
        textParts.push(partText.trim());
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
      text: 'If a tool is needed, respond ONLY as JSON with this shape: {"tool_calls":[{"name":"<tool-name>","arguments":{...}}]}. Do not wrap the JSON in markdown.',
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
        text: "At least one tool call is required before any natural-language answer.",
        description: "Tool choice",
      });
    } else if (
      tooling.toolChoiceMode === ToolChoiceModes.Function &&
      tooling.toolChoiceFunctionName
    ) {
      context.push({
        text: `Only call tool '${tooling.toolChoiceFunctionName}'.`,
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
