export type JsonPrimitive = string | number | boolean | null;
export type JsonValue =
  | JsonPrimitive
  | JsonValue[]
  | { [key: string]: JsonValue };
export type JsonObject = { [key: string]: JsonValue };

export const TransportNames = {
  Graph: "graph",
  Substrate: "substrate",
} as const;

export const ToolChoiceModes = {
  Auto: "auto",
  None: "none",
  Required: "required",
  Function: "function",
} as const;

export const ResponseFormatTypes = {
  JsonObject: "json_object",
  JsonSchema: "json_schema",
} as const;

export type ContextMessage = {
  text: string;
  description: string | null;
};

export type OpenAiToolDefinition = {
  name: string;
  description: string | null;
  parameters: JsonObject;
};

export type OpenAiTooling = {
  tools: OpenAiToolDefinition[];
  toolChoiceMode: string;
  toolChoiceFunctionName: string | null;
  parallelToolCalls: boolean;
};

export type OpenAiResponseFormat = {
  type: string;
  name: string | null;
  jsonSchema: JsonObject | null;
};

export type OpenAiAssistantToolCall = {
  id: string;
  name: string;
  argumentsJson: string;
};

export type OpenAiAssistantResponse = {
  content: string | null;
  toolCalls: OpenAiAssistantToolCall[];
  finishReason: string;
  strictToolErrorMessage?: string | null;
};

export type ParsedOpenAiRequest = {
  model: string;
  stream: boolean;
  promptText: string;
  userKey: string | null;
  locationHint: JsonObject;
  contextualResources: JsonObject | null;
  additionalContext: ContextMessage[];
  tooling: OpenAiTooling;
  responseFormat: OpenAiResponseFormat | null;
  reasoningEffort: string | null;
  temperature: number | null;
};

export type ParsedResponsesRequest = {
  base: ParsedOpenAiRequest;
  previousResponseId: string | null;
  inputItemsForStorage: JsonValue[];
  instructions: string | null;
};

export type JsonPayload = {
  json: JsonObject;
  rawText: string;
};

export type CreateConversationResult = {
  isSuccess: boolean;
  statusCode: number;
  conversationId: string | null;
  rawBody: string;
};

export type ChatResult = {
  isSuccess: boolean;
  statusCode: number;
  responseJson: JsonObject | null;
  rawBody: string;
  assistantText: string | null;
  conversationId: string | null;
};

export type StoredOpenAiResponseRecord = {
  responseId: string;
  createdAtUnix: number;
  response: JsonObject;
  conversationId: string | null;
  expiresAtUtc: number;
};

export type SubstrateStreamUpdate = {
  deltaText: string | null;
  conversationId: string | null;
};

export type SubstrateOptions = {
  hubPath: string;
  source: string;
  quoteSourceInQuery: boolean;
  scenario: string;
  origin: string;
  product: string | null;
  agentHost: string | null;
  licenseType: string | null;
  agent: string | null;
  variants: string | null;
  clientPlatform: string;
  productThreadType: string;
  invocationTimeoutSeconds: number;
  keepAliveSeconds: number;
  optionsSets: string[];
  allowedMessageTypes: string[];
  invocationTarget: string;
  invocationType: number;
  locale: string;
  experienceType: string;
  entityAnnotationTypes: string[];
};

export type WrapperOptions = {
  listenUrl: string;
  debugPath: string | null;
  transport: string;
  graphBaseUrl: string;
  createConversationPath: string;
  chatPathTemplate: string;
  chatOverStreamPathTemplate: string;
  substrate: SubstrateOptions;
  defaultModel: string;
  defaultTimeZone: string;
  conversationTtlMinutes: number;
  maxAdditionalContextMessages: number;
  includeConversationIdInResponseBody: boolean;
};
