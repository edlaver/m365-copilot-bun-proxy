import { createHash, randomUUID } from "node:crypto";
import { Hono } from "hono";
import {
  CopilotGraphClient,
  CopilotSubstrateClient,
  summarizeUpstreamFailure,
} from "./clients";
import { ConversationStore } from "./conversation-store";
import { DebugMarkdownLogger } from "./logger";
import {
  buildAssistantResponse,
  buildChatCompletion,
  buildChatCompletionChunk,
  buildToolCallsDelta,
  computeTrailingDelta,
  extractCopilotAssistantText,
  extractCopilotAssistantTextFromStreamData,
  extractCopilotConversationIdFromStream,
  requiresBufferedAssistantResponse,
  tryBuildAssistantResponseFromChatCompletionPayload,
  tryExtractIncrementalSimulatedChatContent,
  tryExtractSimulatedResponsePayload,
} from "./openai";
import { ResponseStore } from "./response-store";
import {
  buildFunctionCallOutputItems,
  buildMessageOutputItem,
  buildOpenAiResponseFromAssistant,
  buildOpenAiResponseObject,
  buildResponseContentPartAddedEvent,
  buildResponseContentPartDoneEvent,
  buildResponseCompletedEvent,
  buildResponseCreatedEvent,
  buildResponseInProgressEvent,
  buildResponseOutputItemAddedEvent,
  buildResponseOutputItemDoneEvent,
  buildResponseOutputTextDeltaEvent,
  buildResponseOutputTextDoneEvent,
  createOpenAiOutputItemId,
  createOpenAiResponseId,
} from "./responses-api";
import {
  buildCopilotRequestPayload,
  isSupportedOpenAiTransformMode,
  isSupportedTransport,
  normalizeOpenAiTransformMode,
  resolveTransport,
  scopeConversationKey,
  selectConversation,
  tryParseOpenAiRequest,
  tryParseResponsesRequest,
} from "./request-parser";
import {
  OpenAiTransformModes,
  ToolChoiceModes,
  TransportNames,
  type JsonValue,
  type JsonObject,
  type ChatResult,
  type OpenAiAssistantResponse,
  type ParsedOpenAiRequest,
  type ParsedResponsesRequest,
  type SubstrateStreamUpdate,
  type WrapperOptions,
} from "./types";
import { ProxyTokenProvider } from "./token-provider";
import { ProxyVizTraceStore } from "./viz-trace-store";
import {
  cloneJsonValue,
  extractGraphErrorMessage,
  isJsonObject,
  nowUnix,
  readSseEvents,
  tryGetString,
  tryParseJsonObject,
  tryReadJsonPayload,
} from "./utils";

type Services = {
  options: WrapperOptions;
  debugLogger: DebugMarkdownLogger;
  graphClient: CopilotGraphClient;
  substrateClient: CopilotSubstrateClient;
  conversationStore: ConversationStore;
  responseStore: ResponseStore;
  tokenProvider: ProxyTokenProvider;
  vizTraceStore?: ProxyVizTraceStore;
};

type TraceContext = {
  traceId: string;
  requestType: string;
  transformMode: string;
  transport: string;
};

const VizTraceHeaderName = "x-m365-viz-trace-id";
const TransformModeHeaderName = "x-m365-openai-transform-mode";

const AvailableModelIds = [
  "m365-copilot-quick",
  "m365-copilot-reasoning",
  "m365-copilot-gpt5.2-quick",
  "m365-copilot-gpt5.2-reasoning",
  "m365-copilot",
  "m365-copilot-auto",
  "m365-copilot-magic",
] as const;

export function createProxyApp(services: Services): Hono {
  const app = new Hono();

  app.get("/healthz", (c) => c.json({ status: "ok" }));
  app.get("/v1/models", (c) => c.json(buildModelsResponse()));
  app.get("/openai/v1/models", (c) => c.json(buildModelsResponse()));
  app.get("/__viz/traces/:traceId", (c) =>
    handleVizTraceRetrieve(services, c.req.param("traceId")),
  );

  app.post("/v1/chat/completions", (c) => handleChat(c.req.raw, services));
  app.post("/openai/v1/chat/completions", (c) =>
    handleChat(c.req.raw, services),
  );
  app.post("/v1/responses", (c) => handleResponsesCreate(c.req.raw, services));
  app.post("/openai/v1/responses", (c) =>
    handleResponsesCreate(c.req.raw, services),
  );
  app.get("/v1/responses", (c) => handleResponsesList(c.req.raw, services));
  app.get("/openai/v1/responses", (c) =>
    handleResponsesList(c.req.raw, services),
  );
  app.get("/v1/responses/:responseId", (c) =>
    handleResponsesRetrieve(c.req.raw, services, c.req.param("responseId")),
  );
  app.get("/openai/v1/responses/:responseId", (c) =>
    handleResponsesRetrieve(c.req.raw, services, c.req.param("responseId")),
  );
  app.delete("/v1/responses/:responseId", (c) =>
    handleResponsesDelete(c.req.raw, services, c.req.param("responseId")),
  );
  app.delete("/openai/v1/responses/:responseId", (c) =>
    handleResponsesDelete(c.req.raw, services, c.req.param("responseId")),
  );
  return app;
}

function buildModelsResponse(): JsonObject {
  return {
    object: "list",
    data: AvailableModelIds.map((id) => ({
      id,
      object: "model",
      created: 0,
      owned_by: "microsoft-365-copilot",
    })),
  };
}

function handleVizTraceRetrieve(
  services: Services,
  traceId: string,
): Response {
  const normalizedTraceId = traceId.trim();
  if (!normalizedTraceId) {
    return new Response(
      JSON.stringify({
        error: {
          message: "Trace id is required.",
          type: "invalid_request_error",
          param: "traceId",
          code: "missing_trace_id",
        },
      }),
      {
        status: 400,
        headers: { "content-type": "application/json" },
      },
    );
  }

  const record = services.vizTraceStore?.get(normalizedTraceId);
  if (!record) {
    return new Response(
      JSON.stringify({
        error: {
          message: `Unknown trace id '${normalizedTraceId}'.`,
          type: "invalid_request_error",
          param: "traceId",
          code: "unknown_trace_id",
        },
      }),
      {
        status: 404,
        headers: { "content-type": "application/json" },
      },
    );
  }

  return new Response(JSON.stringify(record), {
    status: 200,
    headers: { "content-type": "application/json" },
  });
}

function resolveRequestOptionsWithTransformModeOverride(
  request: Request,
  options: WrapperOptions,
):
  | { ok: true; options: WrapperOptions; transformMode: string }
  | { ok: false; error: string } {
  const requestedMode = request.headers.get(TransformModeHeaderName);
  if (!requestedMode || !requestedMode.trim()) {
    return {
      ok: true,
      options,
      transformMode: normalizeOpenAiTransformMode(options.openAiTransformMode),
    };
  }

  if (!isSupportedOpenAiTransformMode(requestedMode)) {
    return {
      ok: false,
      error: `Unsupported OpenAI transform mode '${requestedMode}'. Supported values: '${OpenAiTransformModes.Simulated}', '${OpenAiTransformModes.Mapped}'.`,
    };
  }

  const transformMode = normalizeOpenAiTransformMode(requestedMode);
  return {
    ok: true,
    options: {
      ...options,
      openAiTransformMode: transformMode,
    },
    transformMode,
  };
}

function resolveTraceContext(
  request: Request,
  requestType: string,
  transformMode: string,
  transport: string,
): TraceContext | null {
  const traceId = request.headers.get(VizTraceHeaderName)?.trim();
  if (!traceId) {
    return null;
  }
  return {
    traceId,
    requestType,
    transformMode,
    transport,
  };
}

function initializeTrace(
  services: Services,
  trace: TraceContext | null,
): void {
  if (!trace) {
    return;
  }
  services.vizTraceStore?.start(
    trace.traceId,
    trace.requestType,
    trace.transformMode,
    trace.transport,
  );
}

function tracePane2(
  services: Services,
  trace: TraceContext | null,
  pane2: JsonValue | null,
  proxyStatusCode: number | null = null,
): void {
  if (!trace) {
    return;
  }
  services.vizTraceStore?.setPane2(trace.traceId, pane2, proxyStatusCode);
}

function tracePane3(
  services: Services,
  trace: TraceContext | null,
  pane3: JsonValue | null,
): void {
  if (!trace) {
    return;
  }
  services.vizTraceStore?.setPane3(trace.traceId, pane3);
}

function tracePane4(
  services: Services,
  trace: TraceContext | null,
  pane4: JsonValue | null,
  upstreamStatusCode: number | null = null,
): void {
  if (!trace) {
    return;
  }
  services.vizTraceStore?.setPane4(trace.traceId, pane4, upstreamStatusCode);
}

function traceSubstrateStreamUpdate(
  services: Services,
  trace: TraceContext | null,
  update: SubstrateStreamUpdate,
): void {
  if (update.upstreamRequestPayload !== undefined) {
    tracePane3(services, trace, update.upstreamRequestPayload ?? null);
  }
  if (update.upstreamResponsePayload !== undefined) {
    tracePane4(services, trace, update.upstreamResponsePayload ?? null);
  }
}

function traceError(
  services: Services,
  trace: TraceContext | null,
  error: JsonValue | null,
  proxyStatusCode: number | null = null,
): void {
  if (!trace) {
    return;
  }
  services.vizTraceStore?.setError(trace.traceId, error, proxyStatusCode);
}

function traceComplete(
  services: Services,
  trace: TraceContext | null,
  proxyStatusCode: number | null = null,
): void {
  if (!trace) {
    return;
  }
  services.vizTraceStore?.complete(trace.traceId, proxyStatusCode);
}

function buildUpstreamStreamCapture(
  streamType: "sse" | "signalr",
  chunks: JsonValue[],
): JsonObject {
  return {
    streamType,
    itemCount: chunks.length,
    items: cloneJsonValue(chunks),
  };
}

async function handleChat(
  request: Request,
  services: Services,
): Promise<Response> {
  const {
    options: baseOptions,
    graphClient,
    substrateClient,
    conversationStore,
    debugLogger,
  } = services;
  const authorizationHeader = await resolveAuthorizationHeader(request, services);
  if (!authorizationHeader) {
    return writeOpenAiError(
      services,
      401,
      "Authorization header is missing/empty and automatic token acquisition failed.",
      "invalid_request_error",
      "missing_authorization",
    );
  }

  const payload = await tryReadJsonPayload(request.clone());
  await debugLogger.logIncomingRequest(request, payload?.rawText ?? null);
  if (!payload) {
    return writeOpenAiError(
      services,
      400,
      "Request body must be valid JSON.",
      "invalid_request_error",
      "invalid_json",
    );
  }

  const resolvedOptionsResult = resolveRequestOptionsWithTransformModeOverride(
    request,
    baseOptions,
  );
  if (!resolvedOptionsResult.ok) {
    return writeOpenAiError(
      services,
      400,
      resolvedOptionsResult.error,
      "invalid_request_error",
      "invalid_transform_mode",
    );
  }
  const options = resolvedOptionsResult.options;
  const selectedTransport = resolveTransport(request, payload.json, options);
  const trace = resolveTraceContext(
    request,
    "chat/completions",
    resolvedOptionsResult.transformMode,
    selectedTransport,
  );
  initializeTrace(services, trace);

  const parsed = tryParseOpenAiRequest(payload.json, options);
  if (!parsed.ok) {
    traceError(
      services,
      trace,
      {
        message: parsed.error,
        type: "invalid_request_error",
        param: null,
        code: "invalid_request",
      },
      400,
    );
    return writeOpenAiError(
      services,
      400,
      parsed.error,
      "invalid_request_error",
      "invalid_request",
    );
  }
  const parsedRequest = parsed.request;

  if (!isSupportedTransport(selectedTransport)) {
    traceError(
      services,
      trace,
      {
        message: `Unsupported transport '${selectedTransport}'.`,
        type: "invalid_request_error",
        param: null,
        code: "invalid_transport",
      },
      400,
    );
    return writeOpenAiError(
      services,
      400,
      `Unsupported transport '${selectedTransport}'. Supported values: '${TransportNames.Graph}', '${TransportNames.Substrate}'.`,
      "invalid_request_error",
      "invalid_transport",
    );
  }

  const responseHeaders = new Headers({
    "x-m365-transport": selectedTransport,
  });
  const conversationSelection = selectConversation(
    request,
    payload.json,
    parsedRequest.userKey,
  );
  const scopedConversationKey = scopeConversationKey(
    conversationSelection.conversationKey,
    selectedTransport,
  );

  let conversationId = conversationSelection.conversationId;
  let createdConversation = false;

  if (!conversationId) {
    if (!conversationSelection.forceNewConversation && scopedConversationKey) {
      const existing = conversationStore.tryGet(scopedConversationKey);
      if (existing) {
        conversationId = existing;
      }
    }

    if (!conversationId) {
      const createResult =
        selectedTransport === TransportNames.Substrate
          ? substrateClient.createConversation()
          : await graphClient.createConversation(authorizationHeader);

      if (!createResult.isSuccess || !createResult.conversationId) {
        const fallbackMessage =
          selectedTransport === TransportNames.Substrate
            ? "Unable to initialize Substrate conversation."
            : "Unable to create Microsoft 365 Copilot conversation.";
        const code =
          selectedTransport === TransportNames.Substrate
            ? "substrate_error"
            : "graph_error";
        return writeFromUpstreamFailure(
          services,
          createResult.statusCode,
          createResult.rawBody,
          fallbackMessage,
          code,
        );
      }

      conversationId = createResult.conversationId;
      createdConversation = true;
      if (scopedConversationKey) {
        conversationStore.set(scopedConversationKey, conversationId);
      }
    }
  }

  if (conversationId && scopedConversationKey) {
    conversationStore.set(scopedConversationKey, conversationId);
  }

  if (!conversationId) {
    return writeOpenAiError(
      services,
      500,
      "Conversation ID resolution failed.",
      "server_error",
      "conversation_id_missing",
    );
  }

  responseHeaders.set("x-m365-conversation-id", conversationId);
  if (createdConversation) {
    responseHeaders.set("x-m365-conversation-created", "true");
  }

  const graphPayload = buildCopilotRequestPayload(parsedRequest);
  if (selectedTransport === TransportNames.Graph) {
    tracePane3(services, trace, graphPayload);
  }
  const shouldBufferAssistant =
    requiresBufferedAssistantResponse(parsedRequest);

  const executeChatTurn = async (): Promise<ChatResult> => {
    if (selectedTransport === TransportNames.Substrate) {
      const result = await substrateClient.chat(
        authorizationHeader,
        conversationId!,
        parsedRequest,
        createdConversation,
        async (update) => {
          traceSubstrateStreamUpdate(services, trace, update);
        },
      );
      tracePane3(services, trace, result.upstreamRequestPayload ?? null);
      tracePane4(
        services,
        trace,
        result.upstreamResponsePayload ?? null,
        result.statusCode,
      );
      return result;
    }
    const result = await graphClient.chat(
      authorizationHeader,
      conversationId!,
      graphPayload,
    );
    tracePane4(
      services,
      trace,
      result.upstreamResponsePayload ?? null,
      result.statusCode,
    );
    return result;
  };
  const executeChatTurnWithRecovery = async (): Promise<ChatResult> => {
    let result = await executeChatTurn();
    if (
      shouldRetrySubstrateNoAssistantContent(
        selectedTransport,
        createdConversation,
        result,
      )
    ) {
      const createRetryConversation = substrateClient.createConversation();
      if (createRetryConversation.isSuccess && createRetryConversation.conversationId) {
        conversationId = createRetryConversation.conversationId;
        createdConversation = true;
        responseHeaders.set("x-m365-conversation-id", conversationId);
        responseHeaders.set("x-m365-conversation-created", "true");
        if (scopedConversationKey) {
          conversationStore.set(scopedConversationKey, conversationId);
        }
        result = await executeChatTurn();
      }
    }
    return result;
  };

  if (parsedRequest.stream) {
    if (
      parsedRequest.transformMode === OpenAiTransformModes.Simulated &&
      selectedTransport === TransportNames.Substrate
    ) {
      return streamSubstrateAsSimulatedOpenAi(
        services,
        authorizationHeader,
        conversationId,
        parsedRequest,
        createdConversation,
        scopedConversationKey,
        responseHeaders,
        trace,
      );
    }

    if (shouldBufferAssistant) {
      const buffered = await executeChatTurnWithRecovery();
      if (!buffered.isSuccess) {
        return writeFromUpstreamFailure(
          services,
          buffered.statusCode,
          buffered.rawBody,
          selectedTransport === TransportNames.Substrate
            ? "Substrate chat request failed."
            : "Microsoft Graph chat request failed.",
          selectedTransport === TransportNames.Substrate
            ? "substrate_error"
            : "graph_error",
        );
      }

      if (buffered.conversationId) {
        conversationId = buffered.conversationId;
        responseHeaders.set("x-m365-conversation-id", conversationId);
        if (scopedConversationKey) {
          conversationStore.set(scopedConversationKey, conversationId);
        }
      }

      const assistantText =
        buffered.assistantText ??
        extractCopilotAssistantText(
          buffered.responseJson,
          parsedRequest.promptText,
        ) ??
        "";
      if (parsedRequest.transformMode === OpenAiTransformModes.Simulated) {
        let simulatedPayload = tryExtractSimulatedResponsePayload(
          assistantText,
          "chat.completions",
        );
        let normalizedSimulatedPayload = simulatedPayload
          ? normalizeSimulatedChatCompletionPayload(
              simulatedPayload,
              parsedRequest.model,
              conversationId,
              options.includeConversationIdInResponseBody,
            )
          : null;

        if (
          !normalizedSimulatedPayload ||
          !hasUsableSimulatedChatCompletionPayload(normalizedSimulatedPayload) ||
          shouldRetrySimulatedInvalidChatToolPayload(normalizedSimulatedPayload) ||
          shouldRetrySimulatedToollessChatPayload(
            options,
            parsedRequest,
            normalizedSimulatedPayload,
          )
        ) {
          const retryResult = await executeChatTurnWithRecovery();
          if (!retryResult.isSuccess) {
            return writeFromUpstreamFailure(
              services,
              retryResult.statusCode,
              retryResult.rawBody,
              selectedTransport === TransportNames.Substrate
                ? "Substrate chat request failed."
                : "Microsoft Graph chat request failed.",
              selectedTransport === TransportNames.Substrate
                ? "substrate_error"
                : "graph_error",
            );
          }
          if (retryResult.conversationId) {
            conversationId = retryResult.conversationId;
            responseHeaders.set("x-m365-conversation-id", conversationId);
            if (scopedConversationKey) {
              conversationStore.set(scopedConversationKey, conversationId);
            }
          }
          const retryAssistantText =
            retryResult.assistantText ??
            extractCopilotAssistantText(
              retryResult.responseJson,
              parsedRequest.promptText,
            ) ??
            "";
          simulatedPayload = tryExtractSimulatedResponsePayload(
            retryAssistantText,
            "chat.completions",
          );
          normalizedSimulatedPayload = simulatedPayload
            ? normalizeSimulatedChatCompletionPayload(
                simulatedPayload,
                parsedRequest.model,
                conversationId,
                options.includeConversationIdInResponseBody,
              )
            : null;
        }

        if (
          !normalizedSimulatedPayload ||
          !hasUsableSimulatedChatCompletionPayload(normalizedSimulatedPayload)
        ) {
          return writeOpenAiError(
            services,
            502,
            "Simulated mode response did not include a usable assistant message or tool call payload.",
            "api_error",
            "invalid_simulated_payload",
          );
        }
        return buildSimulatedChatStreamResponse(
          services,
          parsedRequest.model,
          conversationId,
          normalizedSimulatedPayload,
          options.includeConversationIdInResponseBody,
          responseHeaders,
          trace,
        );
      }

      let assistantResponse = buildAssistantResponse(
        parsedRequest,
        assistantText,
      );
      if (
        shouldRetryStrictToolOutput(
          selectedTransport,
          parsedRequest,
          assistantResponse,
        )
      ) {
        const retryResult = await executeChatTurnWithRecovery();
        if (!retryResult.isSuccess) {
          return writeFromUpstreamFailure(
            services,
            retryResult.statusCode,
            retryResult.rawBody,
            selectedTransport === TransportNames.Substrate
              ? "Substrate chat request failed."
              : "Microsoft Graph chat request failed.",
            selectedTransport === TransportNames.Substrate
              ? "substrate_error"
              : "graph_error",
          );
        }
        if (retryResult.conversationId) {
          conversationId = retryResult.conversationId;
          responseHeaders.set("x-m365-conversation-id", conversationId);
          if (scopedConversationKey) {
            conversationStore.set(scopedConversationKey, conversationId);
          }
        }
        const retryAssistantText =
          retryResult.assistantText ??
          extractCopilotAssistantText(
            retryResult.responseJson,
            parsedRequest.promptText,
          ) ??
          "";
        assistantResponse = buildAssistantResponse(parsedRequest, retryAssistantText);
      }
      const strictToolError = await tryWriteStrictToolOutputError(
        services,
        assistantResponse,
      );
      if (strictToolError) {
        return strictToolError;
      }
      return buildAssistantStreamResponse(
        services,
        parsedRequest.model,
        conversationId,
        assistantResponse,
        options.includeConversationIdInResponseBody,
        responseHeaders,
        trace,
      );
    }

    if (selectedTransport === TransportNames.Graph) {
      const graphResponse = await graphClient.chatOverStream(
        authorizationHeader,
        conversationId,
        graphPayload,
      );
      if (!graphResponse.ok) {
        return writeFromUpstreamFailure(
          services,
          graphResponse.status,
          await graphResponse.text(),
          "Microsoft Graph chatOverStream request failed.",
          "graph_error",
        );
      }
      return transformGraphStreamToOpenAi(
        services,
        graphResponse,
        parsedRequest.model,
        conversationId,
        parsedRequest.promptText,
        options.includeConversationIdInResponseBody,
        responseHeaders,
        trace,
      );
    }

    return streamSubstrateAsOpenAi(
      services,
      authorizationHeader,
      conversationId,
      parsedRequest,
      createdConversation,
      scopedConversationKey,
      responseHeaders,
      trace,
    );
  }

  const chatResponse = await executeChatTurnWithRecovery();
  if (!chatResponse.isSuccess) {
    return writeFromUpstreamFailure(
      services,
      chatResponse.statusCode,
      chatResponse.rawBody,
      selectedTransport === TransportNames.Substrate
        ? "Substrate chat request failed."
        : "Microsoft Graph chat request failed.",
      selectedTransport === TransportNames.Substrate
        ? "substrate_error"
        : "graph_error",
    );
  }

  if (chatResponse.conversationId) {
    conversationId = chatResponse.conversationId;
    responseHeaders.set("x-m365-conversation-id", conversationId);
    if (scopedConversationKey) {
      conversationStore.set(scopedConversationKey, conversationId);
    }
  }

  const assistantText =
    chatResponse.assistantText ??
    extractCopilotAssistantText(
      chatResponse.responseJson,
      parsedRequest.promptText,
    ) ??
    "";
  if (parsedRequest.transformMode === OpenAiTransformModes.Simulated) {
    let simulatedPayload = tryExtractSimulatedResponsePayload(
      assistantText,
      "chat.completions",
    );
    let normalized = simulatedPayload
      ? normalizeSimulatedChatCompletionPayload(
          simulatedPayload,
          parsedRequest.model,
          conversationId,
          options.includeConversationIdInResponseBody,
        )
      : null;

    if (
      !normalized ||
      !hasUsableSimulatedChatCompletionPayload(normalized) ||
      shouldRetrySimulatedInvalidChatToolPayload(normalized) ||
      shouldRetrySimulatedToollessChatPayload(options, parsedRequest, normalized)
    ) {
      const retryResult = await executeChatTurnWithRecovery();
      if (!retryResult.isSuccess) {
        return writeFromUpstreamFailure(
          services,
          retryResult.statusCode,
          retryResult.rawBody,
          selectedTransport === TransportNames.Substrate
            ? "Substrate chat request failed."
            : "Microsoft Graph chat request failed.",
          selectedTransport === TransportNames.Substrate
            ? "substrate_error"
            : "graph_error",
        );
      }
      if (retryResult.conversationId) {
        conversationId = retryResult.conversationId;
        responseHeaders.set("x-m365-conversation-id", conversationId);
        if (scopedConversationKey) {
          conversationStore.set(scopedConversationKey, conversationId);
        }
      }
      const retryAssistantText =
        retryResult.assistantText ??
        extractCopilotAssistantText(
          retryResult.responseJson,
          parsedRequest.promptText,
        ) ??
        "";
      simulatedPayload = tryExtractSimulatedResponsePayload(
        retryAssistantText,
        "chat.completions",
      );
      normalized = simulatedPayload
        ? normalizeSimulatedChatCompletionPayload(
            simulatedPayload,
            parsedRequest.model,
            conversationId,
            options.includeConversationIdInResponseBody,
          )
        : null;
    }

    if (
      !normalized ||
      !hasUsableSimulatedChatCompletionPayload(normalized)
    ) {
      traceError(
        services,
        trace,
        {
          message:
            "Simulated mode response did not include a usable assistant message or tool call payload.",
          type: "api_error",
          param: null,
          code: "invalid_simulated_payload",
        },
        502,
      );
      return writeOpenAiError(
        services,
        502,
        "Simulated mode response did not include a usable assistant message or tool call payload.",
        "api_error",
        "invalid_simulated_payload",
      );
    }
    const body = JSON.stringify(normalized);
    tracePane2(services, trace, normalized, 200);
    traceComplete(services, trace, 200);
    responseHeaders.set("content-type", "application/json");
    await debugLogger.logOutgoingResponse(200, responseHeaders.entries(), body);
    return new Response(body, { status: 200, headers: responseHeaders });
  }

  let assistantResponse = buildAssistantResponse(
    parsedRequest,
    assistantText,
  );
  if (
    shouldRetryStrictToolOutput(
      selectedTransport,
      parsedRequest,
      assistantResponse,
    )
  ) {
    const retryResult = await executeChatTurnWithRecovery();
    if (!retryResult.isSuccess) {
      return writeFromUpstreamFailure(
        services,
        retryResult.statusCode,
        retryResult.rawBody,
        selectedTransport === TransportNames.Substrate
          ? "Substrate chat request failed."
          : "Microsoft Graph chat request failed.",
        selectedTransport === TransportNames.Substrate
          ? "substrate_error"
          : "graph_error",
      );
    }
    if (retryResult.conversationId) {
      conversationId = retryResult.conversationId;
      responseHeaders.set("x-m365-conversation-id", conversationId);
      if (scopedConversationKey) {
        conversationStore.set(scopedConversationKey, conversationId);
      }
    }
    const retryAssistantText =
      retryResult.assistantText ??
      extractCopilotAssistantText(
        retryResult.responseJson,
        parsedRequest.promptText,
      ) ??
      "";
    assistantResponse = buildAssistantResponse(parsedRequest, retryAssistantText);
  }
  const strictToolError = await tryWriteStrictToolOutputError(
    services,
    assistantResponse,
  );
  if (strictToolError) {
    return strictToolError;
  }
  const body = JSON.stringify(
    buildChatCompletion(
      parsedRequest.model,
      assistantResponse,
      conversationId,
      options.includeConversationIdInResponseBody,
    ),
  );
  tracePane2(services, trace, JSON.parse(body) as JsonValue, 200);
  traceComplete(services, trace, 200);

  responseHeaders.set("content-type", "application/json");
  await debugLogger.logOutgoingResponse(200, responseHeaders.entries(), body);
  return new Response(body, { status: 200, headers: responseHeaders });
}

async function handleResponsesCreate(
  request: Request,
  services: Services,
): Promise<Response> {
  const {
    options: baseOptions,
    graphClient,
    substrateClient,
    conversationStore,
    responseStore,
    debugLogger,
  } = services;
  const authorizationHeader = await resolveAuthorizationHeader(request, services);
  if (!authorizationHeader) {
    return writeOpenAiError(
      services,
      401,
      "Authorization header is missing/empty and automatic token acquisition failed.",
      "invalid_request_error",
      "missing_authorization",
    );
  }

  const payload = await tryReadJsonPayload(request.clone());
  await debugLogger.logIncomingRequest(request, payload?.rawText ?? null);
  if (!payload) {
    return writeOpenAiError(
      services,
      400,
      "Request body must be valid JSON.",
      "invalid_request_error",
      "invalid_json",
    );
  }

  const resolvedOptionsResult = resolveRequestOptionsWithTransformModeOverride(
    request,
    baseOptions,
  );
  if (!resolvedOptionsResult.ok) {
    return writeOpenAiError(
      services,
      400,
      resolvedOptionsResult.error,
      "invalid_request_error",
      "invalid_transform_mode",
    );
  }
  const options = resolvedOptionsResult.options;
  const selectedTransport = resolveTransport(request, payload.json, options);
  const trace = resolveTraceContext(
    request,
    "responses",
    resolvedOptionsResult.transformMode,
    selectedTransport,
  );
  initializeTrace(services, trace);

  const parsed = tryParseResponsesRequest(payload.json, options);
  if (!parsed.ok) {
    traceError(
      services,
      trace,
      {
        message: parsed.error,
        type: "invalid_request_error",
        param: null,
        code: "invalid_request",
      },
      400,
    );
    return writeOpenAiError(
      services,
      400,
      parsed.error,
      "invalid_request_error",
      "invalid_request",
    );
  }
  const parsedRequest = parsed.request;
  const baseRequest = parsedRequest.base;

  if (!isSupportedTransport(selectedTransport)) {
    traceError(
      services,
      trace,
      {
        message: `Unsupported transport '${selectedTransport}'.`,
        type: "invalid_request_error",
        param: null,
        code: "invalid_transport",
      },
      400,
    );
    return writeOpenAiError(
      services,
      400,
      `Unsupported transport '${selectedTransport}'. Supported values: '${TransportNames.Graph}', '${TransportNames.Substrate}'.`,
      "invalid_request_error",
      "invalid_transport",
    );
  }

  const responseHeaders = new Headers({
    "x-m365-transport": selectedTransport,
  });
  const requestHash = computeResponsesRequestHash(
    payload.rawText,
    payload.json,
    selectedTransport,
  );
  const replayLoopHash = computeResponsesReplayLoopHash(
    payload.json,
    selectedTransport,
  );
  const replayResponseId = buildReplayResponseIdFromHash(replayLoopHash);
  const suppressedConversationId = resolveSuppressedResponsesConversationId(
    request,
    payload.json,
    parsedRequest,
    selectedTransport,
    conversationStore,
    responseStore,
  );
  const trailingAssistantReplay = detectAssistantTailReplayWithoutNewUserTurn(
    parsedRequest.inputItemsForStorage,
  );
  if (responseStore.hasRecentRequestHash(requestHash)) {
    responseHeaders.set("x-m365-request-hash-replayed", "true");
    const replayConversationId =
      responseStore.getRecentRequestHashConversationId(requestHash) ??
      responseStore.getRecentRequestHashConversationId(replayLoopHash) ??
      suppressedConversationId;
    rememberResponsesReplayHashes(
      responseStore,
      requestHash,
      replayLoopHash,
      replayConversationId,
    );
    return buildSuppressedReplayResponsesResult(
      services,
      parsedRequest,
      responseHeaders,
      replayConversationId,
      trailingAssistantReplay,
      replayResponseId,
    );
  }
  if (trailingAssistantReplay) {
    const replayConversationId =
      responseStore.getRecentRequestHashConversationId(requestHash) ??
      responseStore.getRecentRequestHashConversationId(replayLoopHash) ??
      suppressedConversationId;
    rememberResponsesReplayHashes(
      responseStore,
      requestHash,
      replayLoopHash,
      replayConversationId,
    );
    return buildSuppressedReplayResponsesResult(
      services,
      parsedRequest,
      responseHeaders,
      replayConversationId,
      trailingAssistantReplay,
      replayResponseId,
    );
  }
  const conversationSelection = selectConversation(
    request,
    payload.json,
    baseRequest.userKey,
  );
  const scopedConversationKey = scopeConversationKey(
    conversationSelection.conversationKey,
    selectedTransport,
  );

  let conversationId = conversationSelection.conversationId;
  let createdConversation = false;

  if (!conversationId && parsedRequest.previousResponseId) {
    const previousConversationId = responseStore.tryGetConversationLink(
      parsedRequest.previousResponseId,
    );
    if (!previousConversationId) {
      return writeOpenAiError(
        services,
        400,
        `Unknown previous_response_id '${parsedRequest.previousResponseId}'.`,
        "invalid_request_error",
        "invalid_previous_response_id",
      );
    }
    conversationId = previousConversationId;
  }

  if (!conversationId) {
    if (!conversationSelection.forceNewConversation && scopedConversationKey) {
      const existing = conversationStore.tryGet(scopedConversationKey);
      if (existing) {
        conversationId = existing;
      }
    }

    if (!conversationId) {
      const createResult =
        selectedTransport === TransportNames.Substrate
          ? substrateClient.createConversation()
          : await graphClient.createConversation(authorizationHeader);

      if (!createResult.isSuccess || !createResult.conversationId) {
        const fallbackMessage =
          selectedTransport === TransportNames.Substrate
            ? "Unable to initialize Substrate conversation."
            : "Unable to create Microsoft 365 Copilot conversation.";
        const code =
          selectedTransport === TransportNames.Substrate
            ? "substrate_error"
            : "graph_error";
        return writeFromUpstreamFailure(
          services,
          createResult.statusCode,
          createResult.rawBody,
          fallbackMessage,
          code,
        );
      }

      conversationId = createResult.conversationId;
      createdConversation = true;
      if (scopedConversationKey) {
        conversationStore.set(scopedConversationKey, conversationId);
      }
    }
  }

  if (conversationId && scopedConversationKey) {
    conversationStore.set(scopedConversationKey, conversationId);
  }

  if (!conversationId) {
    return writeOpenAiError(
      services,
      500,
      "Conversation ID resolution failed.",
      "server_error",
      "conversation_id_missing",
    );
  }

  responseHeaders.set("x-m365-conversation-id", conversationId);
  if (createdConversation) {
    responseHeaders.set("x-m365-conversation-created", "true");
  }

  const graphPayload = buildCopilotRequestPayload(baseRequest);
  if (selectedTransport === TransportNames.Graph) {
    tracePane3(services, trace, graphPayload);
  }
  const shouldBufferAssistant = requiresBufferedAssistantResponse(baseRequest);

  const executeChatTurn = async (): Promise<ChatResult> => {
    if (selectedTransport === TransportNames.Substrate) {
      const result = await substrateClient.chat(
        authorizationHeader,
        conversationId!,
        baseRequest,
        createdConversation,
        async (update) => {
          traceSubstrateStreamUpdate(services, trace, update);
        },
      );
      tracePane3(services, trace, result.upstreamRequestPayload ?? null);
      tracePane4(
        services,
        trace,
        result.upstreamResponsePayload ?? null,
        result.statusCode,
      );
      return result;
    }
    const result = await graphClient.chat(
      authorizationHeader,
      conversationId!,
      graphPayload,
    );
    tracePane4(
      services,
      trace,
      result.upstreamResponsePayload ?? null,
      result.statusCode,
    );
    return result;
  };
  const executeChatTurnWithRecovery = async (): Promise<ChatResult> => {
    let result = await executeChatTurn();
    if (
      shouldRetrySubstrateNoAssistantContent(
        selectedTransport,
        createdConversation,
        result,
      )
    ) {
      const createRetryConversation = substrateClient.createConversation();
      if (createRetryConversation.isSuccess && createRetryConversation.conversationId) {
        conversationId = createRetryConversation.conversationId;
        createdConversation = true;
        responseHeaders.set("x-m365-conversation-id", conversationId);
        responseHeaders.set("x-m365-conversation-created", "true");
        if (scopedConversationKey) {
          conversationStore.set(scopedConversationKey, conversationId);
        }
        result = await executeChatTurn();
      }
    }
    return result;
  };

  if (baseRequest.stream) {
    if (shouldBufferAssistant) {
      const buffered = await executeChatTurnWithRecovery();
      if (!buffered.isSuccess) {
        return writeFromUpstreamFailure(
          services,
          buffered.statusCode,
          buffered.rawBody,
          selectedTransport === TransportNames.Substrate
            ? "Substrate chat request failed."
            : "Microsoft Graph chat request failed.",
          selectedTransport === TransportNames.Substrate
            ? "substrate_error"
            : "graph_error",
        );
      }

      if (buffered.conversationId) {
        conversationId = buffered.conversationId;
        responseHeaders.set("x-m365-conversation-id", conversationId);
        if (scopedConversationKey) {
          conversationStore.set(scopedConversationKey, conversationId);
        }
      }

      const assistantText =
        buffered.assistantText ??
        extractCopilotAssistantText(
          buffered.responseJson,
          baseRequest.promptText,
        ) ??
        "";
      if (baseRequest.transformMode === OpenAiTransformModes.Simulated) {
        let simulatedPayload = tryExtractSimulatedResponsePayload(
          assistantText,
          "responses",
        );
        let normalized = simulatedPayload
          ? normalizeSimulatedResponsesPayload(
              simulatedPayload,
              parsedRequest,
              conversationId,
              options.includeConversationIdInResponseBody,
            )
          : null;

        if (
          !normalized ||
          !hasUsableSimulatedResponsesPayload(normalized.responseBody) ||
          shouldRetrySimulatedInvalidResponsesToolPayload(
            normalized.responseBody,
          ) ||
          shouldRetrySimulatedToollessResponsesPayload(
            options,
            baseRequest,
            normalized.responseBody,
          )
        ) {
          const retryResult = await executeChatTurnWithRecovery();
          if (!retryResult.isSuccess) {
            return writeFromUpstreamFailure(
              services,
              retryResult.statusCode,
              retryResult.rawBody,
              selectedTransport === TransportNames.Substrate
                ? "Substrate chat request failed."
                : "Microsoft Graph chat request failed.",
              selectedTransport === TransportNames.Substrate
                ? "substrate_error"
                : "graph_error",
            );
          }
          if (retryResult.conversationId) {
            conversationId = retryResult.conversationId;
            responseHeaders.set("x-m365-conversation-id", conversationId);
            if (scopedConversationKey) {
              conversationStore.set(scopedConversationKey, conversationId);
            }
          }
          const retryAssistantText =
            retryResult.assistantText ??
            extractCopilotAssistantText(
              retryResult.responseJson,
              baseRequest.promptText,
            ) ??
            "";
          simulatedPayload = tryExtractSimulatedResponsePayload(
            retryAssistantText,
            "responses",
          );
          normalized = simulatedPayload
            ? normalizeSimulatedResponsesPayload(
                simulatedPayload,
                parsedRequest,
                conversationId,
                options.includeConversationIdInResponseBody,
              )
            : null;
        }

        if (
          !normalized ||
          !hasUsableSimulatedResponsesPayload(normalized.responseBody)
        ) {
          return writeOpenAiError(
            services,
            502,
            "Simulated mode response did not include a usable response output payload.",
            "api_error",
            "invalid_simulated_payload",
          );
        }
        rememberResponsesReplayHashes(
          responseStore,
          requestHash,
          replayLoopHash,
          conversationId,
        );
        return buildSimulatedResponsesStreamResponse(
          services,
          parsedRequest,
          conversationId,
          normalized.responseBody,
          responseHeaders,
          trace,
        );
      }

      let assistantResponse = buildAssistantResponse(baseRequest, assistantText);
      if (
        shouldRetryStrictToolOutput(
          selectedTransport,
          baseRequest,
          assistantResponse,
        )
      ) {
        const retryResult = await executeChatTurnWithRecovery();
        if (!retryResult.isSuccess) {
          return writeFromUpstreamFailure(
            services,
            retryResult.statusCode,
            retryResult.rawBody,
            selectedTransport === TransportNames.Substrate
              ? "Substrate chat request failed."
              : "Microsoft Graph chat request failed.",
            selectedTransport === TransportNames.Substrate
              ? "substrate_error"
              : "graph_error",
          );
        }
        if (retryResult.conversationId) {
          conversationId = retryResult.conversationId;
          responseHeaders.set("x-m365-conversation-id", conversationId);
          if (scopedConversationKey) {
            conversationStore.set(scopedConversationKey, conversationId);
          }
        }
        const retryAssistantText =
          retryResult.assistantText ??
          extractCopilotAssistantText(
            retryResult.responseJson,
            baseRequest.promptText,
          ) ??
          "";
        assistantResponse = buildAssistantResponse(baseRequest, retryAssistantText);
      }
      const strictToolError = await tryWriteStrictToolOutputError(
        services,
        assistantResponse,
      );
      if (strictToolError) {
        return strictToolError;
      }
      rememberResponsesReplayHashes(
        responseStore,
        requestHash,
        replayLoopHash,
        conversationId,
      );
      return buildBufferedResponsesStreamResponse(
        services,
        parsedRequest,
        conversationId,
        assistantResponse,
        responseHeaders,
        trace,
      );
    }

    if (selectedTransport === TransportNames.Graph) {
      const graphResponse = await graphClient.chatOverStream(
        authorizationHeader,
        conversationId,
        graphPayload,
      );
      if (!graphResponse.ok) {
        return writeFromUpstreamFailure(
          services,
          graphResponse.status,
          await graphResponse.text(),
          "Microsoft Graph chatOverStream request failed.",
          "graph_error",
        );
      }
      rememberResponsesReplayHashes(
        responseStore,
        requestHash,
        replayLoopHash,
        conversationId,
      );
      return transformGraphStreamToResponses(
        services,
        graphResponse,
        parsedRequest,
        conversationId,
        scopedConversationKey,
        responseHeaders,
        trace,
      );
    }

    rememberResponsesReplayHashes(
      responseStore,
      requestHash,
      replayLoopHash,
      conversationId,
    );
    return streamSubstrateAsResponses(
      services,
      authorizationHeader,
      conversationId,
      parsedRequest,
      createdConversation,
      scopedConversationKey,
      responseHeaders,
      trace,
    );
  }

  const chatResponse = await executeChatTurnWithRecovery();
  if (!chatResponse.isSuccess) {
    return writeFromUpstreamFailure(
      services,
      chatResponse.statusCode,
      chatResponse.rawBody,
      selectedTransport === TransportNames.Substrate
        ? "Substrate chat request failed."
        : "Microsoft Graph chat request failed.",
      selectedTransport === TransportNames.Substrate
        ? "substrate_error"
        : "graph_error",
    );
  }

  if (chatResponse.conversationId) {
    conversationId = chatResponse.conversationId;
    responseHeaders.set("x-m365-conversation-id", conversationId);
    if (scopedConversationKey) {
      conversationStore.set(scopedConversationKey, conversationId);
    }
  }

  const assistantText =
    chatResponse.assistantText ??
    extractCopilotAssistantText(chatResponse.responseJson, baseRequest.promptText) ??
    "";
  if (baseRequest.transformMode === OpenAiTransformModes.Simulated) {
    let simulatedPayload = tryExtractSimulatedResponsePayload(
      assistantText,
      "responses",
    );
    let normalized = simulatedPayload
      ? normalizeSimulatedResponsesPayload(
          simulatedPayload,
          parsedRequest,
          conversationId,
          options.includeConversationIdInResponseBody,
        )
      : null;

    if (
      !normalized ||
      !hasUsableSimulatedResponsesPayload(normalized.responseBody) ||
      shouldRetrySimulatedInvalidResponsesToolPayload(normalized.responseBody) ||
      shouldRetrySimulatedToollessResponsesPayload(
        options,
        baseRequest,
        normalized.responseBody,
      )
    ) {
      const retryResult = await executeChatTurnWithRecovery();
      if (!retryResult.isSuccess) {
        return writeFromUpstreamFailure(
          services,
          retryResult.statusCode,
          retryResult.rawBody,
          selectedTransport === TransportNames.Substrate
            ? "Substrate chat request failed."
            : "Microsoft Graph chat request failed.",
          selectedTransport === TransportNames.Substrate
            ? "substrate_error"
            : "graph_error",
        );
      }
      if (retryResult.conversationId) {
        conversationId = retryResult.conversationId;
        responseHeaders.set("x-m365-conversation-id", conversationId);
        if (scopedConversationKey) {
          conversationStore.set(scopedConversationKey, conversationId);
        }
      }
      const retryAssistantText =
        retryResult.assistantText ??
        extractCopilotAssistantText(
          retryResult.responseJson,
          baseRequest.promptText,
        ) ??
        "";
      simulatedPayload = tryExtractSimulatedResponsePayload(
        retryAssistantText,
        "responses",
      );
      normalized = simulatedPayload
        ? normalizeSimulatedResponsesPayload(
            simulatedPayload,
            parsedRequest,
            conversationId,
            options.includeConversationIdInResponseBody,
          )
        : null;
    }

    if (
      !normalized ||
      !hasUsableSimulatedResponsesPayload(normalized.responseBody)
    ) {
      traceError(
        services,
        trace,
        {
          message:
            "Simulated mode response did not include a usable response output payload.",
          type: "api_error",
          param: null,
          code: "invalid_simulated_payload",
        },
        502,
      );
      return writeOpenAiError(
        services,
        502,
        "Simulated mode response did not include a usable response output payload.",
        "api_error",
        "invalid_simulated_payload",
      );
    }

    responseStore.set(normalized.responseId, normalized.responseBody, conversationId);
    rememberResponsesReplayHashes(
      responseStore,
      requestHash,
      replayLoopHash,
      conversationId,
    );
    responseHeaders.set("content-type", "application/json");
    const body = JSON.stringify(normalized.responseBody);
    tracePane2(services, trace, normalized.responseBody, 200);
    traceComplete(services, trace, 200);
    await debugLogger.logOutgoingResponse(200, responseHeaders.entries(), body);
    return new Response(body, { status: 200, headers: responseHeaders });
  }

  let assistantResponse = buildAssistantResponse(baseRequest, assistantText);
  if (
    shouldRetryStrictToolOutput(
      selectedTransport,
      baseRequest,
      assistantResponse,
    )
  ) {
    const retryResult = await executeChatTurnWithRecovery();
    if (!retryResult.isSuccess) {
      return writeFromUpstreamFailure(
        services,
        retryResult.statusCode,
        retryResult.rawBody,
        selectedTransport === TransportNames.Substrate
          ? "Substrate chat request failed."
          : "Microsoft Graph chat request failed.",
        selectedTransport === TransportNames.Substrate
          ? "substrate_error"
          : "graph_error",
      );
    }
    if (retryResult.conversationId) {
      conversationId = retryResult.conversationId;
      responseHeaders.set("x-m365-conversation-id", conversationId);
      if (scopedConversationKey) {
        conversationStore.set(scopedConversationKey, conversationId);
      }
    }
    const retryAssistantText =
      retryResult.assistantText ??
      extractCopilotAssistantText(
        retryResult.responseJson,
        baseRequest.promptText,
      ) ??
      "";
    assistantResponse = buildAssistantResponse(baseRequest, retryAssistantText);
  }
  const strictToolError = await tryWriteStrictToolOutputError(
    services,
    assistantResponse,
  );
  if (strictToolError) {
    return strictToolError;
  }
  const responseId = createOpenAiResponseId();
  const createdAt = nowUnix();
  const responseBody = buildOpenAiResponseFromAssistant(
    responseId,
    createdAt,
    baseRequest.model,
    "completed",
    parsedRequest,
    assistantResponse,
    options.includeConversationIdInResponseBody,
    conversationId,
  );
  responseStore.set(responseId, responseBody, conversationId);
  rememberResponsesReplayHashes(
    responseStore,
    requestHash,
    replayLoopHash,
    conversationId,
  );

  responseHeaders.set("content-type", "application/json");
  const body = JSON.stringify(responseBody);
  tracePane2(services, trace, responseBody, 200);
  traceComplete(services, trace, 200);
  await debugLogger.logOutgoingResponse(200, responseHeaders.entries(), body);
  return new Response(body, { status: 200, headers: responseHeaders });
}

async function handleResponsesList(
  request: Request,
  services: Services,
): Promise<Response> {
  const authorizationHeader = await resolveAuthorizationHeader(request, services);
  await services.debugLogger.logIncomingRequest(request, null);
  if (!authorizationHeader) {
    return writeOpenAiError(
      services,
      401,
      "Authorization header is missing/empty and automatic token acquisition failed.",
      "invalid_request_error",
      "missing_authorization",
    );
  }

  const requestUrl = new URL(request.url);
  const limit = clampListLimit(requestUrl.searchParams.get("limit"));
  const listed = services.responseStore.list(limit);
  const body = JSON.stringify({
    object: "list",
    data: listed.data,
    has_more: listed.hasMore,
    first_id: listed.firstId,
    last_id: listed.lastId,
  });
  const headers = new Headers({ "content-type": "application/json" });
  await services.debugLogger.logOutgoingResponse(200, headers.entries(), body);
  return new Response(body, { status: 200, headers });
}

async function handleResponsesRetrieve(
  request: Request,
  services: Services,
  responseIdParam: string,
): Promise<Response> {
  const authorizationHeader = await resolveAuthorizationHeader(request, services);
  await services.debugLogger.logIncomingRequest(request, null);
  if (!authorizationHeader) {
    return writeOpenAiError(
      services,
      401,
      "Authorization header is missing/empty and automatic token acquisition failed.",
      "invalid_request_error",
      "missing_authorization",
    );
  }

  const responseId = responseIdParam.trim();
  if (!responseId) {
    return writeOpenAiError(
      services,
      400,
      "The response ID is required.",
      "invalid_request_error",
      "missing_response_id",
    );
  }

  const response = services.responseStore.tryGet(responseId);
  if (!response) {
    return writeOpenAiError(
      services,
      404,
      `Response '${responseId}' was not found.`,
      "invalid_request_error",
      "response_not_found",
    );
  }

  const headers = new Headers({ "content-type": "application/json" });
  const body = JSON.stringify(response);
  await services.debugLogger.logOutgoingResponse(200, headers.entries(), body);
  return new Response(body, { status: 200, headers });
}

async function handleResponsesDelete(
  request: Request,
  services: Services,
  responseIdParam: string,
): Promise<Response> {
  const authorizationHeader = await resolveAuthorizationHeader(request, services);
  await services.debugLogger.logIncomingRequest(request, null);
  if (!authorizationHeader) {
    return writeOpenAiError(
      services,
      401,
      "Authorization header is missing/empty and automatic token acquisition failed.",
      "invalid_request_error",
      "missing_authorization",
    );
  }

  const responseId = responseIdParam.trim();
  if (!responseId) {
    return writeOpenAiError(
      services,
      400,
      "The response ID is required.",
      "invalid_request_error",
      "missing_response_id",
    );
  }

  const deleted = services.responseStore.tryDelete(responseId);
  if (!deleted) {
    return writeOpenAiError(
      services,
      404,
      `Response '${responseId}' was not found.`,
      "invalid_request_error",
      "response_not_found",
    );
  }

  const headers = new Headers({ "content-type": "application/json" });
  const body = JSON.stringify({
    id: responseId,
    object: "response",
    deleted: true,
  });
  await services.debugLogger.logOutgoingResponse(200, headers.entries(), body);
  return new Response(body, { status: 200, headers });
}

async function buildSuppressedReplayResponsesResult(
  services: Services,
  parsedRequest: ParsedResponsesRequest,
  headers: Headers,
  conversationId: string | null,
  replayText: string | null,
  replayResponseId: string | null = null,
): Promise<Response> {
  const responseId = replayResponseId?.trim()
    ? replayResponseId.trim()
    : createOpenAiResponseId();
  const createdAt = nowUnix();
  const includeConversationId = services.options.includeConversationIdInResponseBody;
  const normalizedConversationId = conversationId?.trim()
    ? conversationId.trim()
    : null;
  const responseConversationId = includeConversationId
    ? normalizedConversationId
    : null;
  const replayTextValue = replayText?.trim() ? replayText : null;
  const outputItems = [
    buildMessageOutputItem(
      createOpenAiOutputItemId("msg"),
      replayTextValue ?? "",
      "completed",
    ),
  ];
  const inProgress = buildOpenAiResponseObject(
    responseId,
    createdAt,
    parsedRequest.base.model,
    "in_progress",
    [],
    parsedRequest,
    responseConversationId,
  );
  const completed = buildOpenAiResponseObject(
    responseId,
    createdAt,
    parsedRequest.base.model,
    "completed",
    outputItems,
    parsedRequest,
    responseConversationId,
  );
  services.responseStore.set(responseId, completed, normalizedConversationId);
  headers.set("x-m365-replay-suppressed", "true");
  if (normalizedConversationId) {
    headers.set("x-m365-conversation-id", normalizedConversationId);
  }

  if (!parsedRequest.base.stream) {
    headers.set("content-type", "application/json");
    const body = JSON.stringify(completed);
    await services.debugLogger.logOutgoingResponse(200, headers.entries(), body);
    return new Response(body, { status: 200, headers });
  }

  const stream = new ReadableStream<Uint8Array>({
    start(controller) {
      const encoder = new TextEncoder();
      const writeDataEvent = (event: JsonObject) => {
        enqueueSseJsonEvent(controller, encoder, event);
      };

      writeDataEvent(buildResponseCreatedEvent(inProgress));
      writeDataEvent(buildResponseInProgressEvent(inProgress));
      for (let index = 0; index < outputItems.length; index += 1) {
        const outputItem = outputItems[index];
        const outputItemId = String(outputItem.id ?? createOpenAiOutputItemId("msg"));
        writeDataEvent(
          buildResponseOutputItemAddedEvent(
            responseId,
            index,
            buildMessageOutputItem(outputItemId, "", "in_progress"),
          ),
        );
        writeDataEvent(
          buildResponseContentPartAddedEvent(
            responseId,
            index,
            outputItemId,
            { type: "output_text", text: "" },
          ),
        );
        if (replayTextValue) {
          writeDataEvent(
            buildResponseOutputTextDeltaEvent(
              responseId,
              index,
              outputItemId,
              replayTextValue,
            ),
          );
        }
        writeDataEvent(
          buildResponseOutputTextDoneEvent(
            responseId,
            index,
            outputItemId,
            replayTextValue ?? "",
          ),
        );
        writeDataEvent(
          buildResponseContentPartDoneEvent(
            responseId,
            index,
            outputItemId,
            { type: "output_text", text: replayTextValue ?? "" },
          ),
        );
        writeDataEvent(
          buildResponseOutputItemDoneEvent(responseId, index, outputItem),
        );
      }
      writeDataEvent(buildResponseCompletedEvent(completed));
      enqueueSseDoneEvent(controller, encoder);
      controller.close();
    },
  });

  headers.set("content-type", "text/event-stream");
  headers.set("cache-control", "no-cache");
  headers.set("connection", "keep-alive");
  await services.debugLogger.logOutgoingResponse(200, headers.entries(), null);
  return finalizeOutgoingStreamResponse(services, stream, headers);
}

async function buildBufferedResponsesStreamResponse(
  services: Services,
  parsedRequest: ParsedResponsesRequest,
  conversationId: string,
  assistantResponse: ReturnType<typeof buildAssistantResponse>,
  headers: Headers,
  trace: TraceContext | null,
): Promise<Response> {
  const responseId = createOpenAiResponseId();
  const createdAt = nowUnix();
  const includeConversationId = services.options.includeConversationIdInResponseBody;
  const stream = new ReadableStream<Uint8Array>({
    start(controller) {
      const encoder = new TextEncoder();
      const writeDataEvent = (event: JsonObject) => {
        enqueueSseJsonEvent(controller, encoder, event);
      };
      const writeError = (message: string, code: string) => {
        controller.enqueue(
          encoder.encode(
            `event: error\ndata: ${JSON.stringify({
              error: {
                message,
                type: "api_error",
                param: null,
                code,
              },
            })}\n\n`,
          ),
        );
      };

      try {
        const inProgress = buildOpenAiResponseObject(
          responseId,
          createdAt,
          parsedRequest.base.model,
          "in_progress",
          [],
          parsedRequest,
          includeConversationId ? conversationId : null,
        );
        writeDataEvent(buildResponseCreatedEvent(inProgress));
        writeDataEvent(buildResponseInProgressEvent(inProgress));

        const outputItems =
          assistantResponse.toolCalls.length > 0
            ? buildFunctionCallOutputItems(assistantResponse.toolCalls, "completed")
            : [
                buildMessageOutputItem(
                  createOpenAiOutputItemId("msg"),
                  assistantResponse.content ?? "",
                  "completed",
                ),
              ];

        for (let index = 0; index < outputItems.length; index++) {
          const item = outputItems[index];
          writeDataEvent(
            buildResponseOutputItemAddedEvent(
              responseId,
              index,
              item.type === "message"
                ? buildMessageOutputItem(String(item.id ?? ""), "", "in_progress")
                : item,
            ),
          );
          if (item.type === "message") {
            const content = assistantResponse.content ?? "";
            writeDataEvent(
              buildResponseContentPartAddedEvent(
                responseId,
                index,
                String(item.id ?? ""),
                { type: "output_text", text: "" },
              ),
            );
            if (content) {
              writeDataEvent(
                buildResponseOutputTextDeltaEvent(
                  responseId,
                  index,
                  String(item.id ?? ""),
                  content,
                ),
              );
            }
            writeDataEvent(
              buildResponseOutputTextDoneEvent(
                responseId,
                index,
                String(item.id ?? ""),
                content,
              ),
            );
            writeDataEvent(
              buildResponseContentPartDoneEvent(
                responseId,
                index,
                String(item.id ?? ""),
                { type: "output_text", text: content },
              ),
            );
          }
          writeDataEvent(buildResponseOutputItemDoneEvent(responseId, index, item));
        }

        const completed = buildOpenAiResponseObject(
          responseId,
          createdAt,
          parsedRequest.base.model,
          "completed",
          outputItems,
          parsedRequest,
          includeConversationId ? conversationId : null,
        );
        writeDataEvent(buildResponseCompletedEvent(completed));
        services.responseStore.set(responseId, completed, conversationId);
        tracePane2(services, trace, completed, 200);
        traceComplete(services, trace, 200);
      } catch (error) {
        traceError(
          services,
          trace,
          {
            message: `Failed to build streaming response. ${String(error)}`,
            type: "api_error",
            param: null,
            code: "response_stream_error",
          },
          500,
        );
        writeError(
          `Failed to build streaming response. ${String(error)}`,
          "response_stream_error",
        );
      } finally {
        enqueueSseDoneEvent(controller, encoder);
        controller.close();
      }
    },
  });

  headers.set("content-type", "text/event-stream");
  headers.set("cache-control", "no-cache");
  headers.set("connection", "keep-alive");
  headers.set("x-m365-conversation-id", conversationId);
  await services.debugLogger.logOutgoingResponse(200, headers.entries(), null);
  return finalizeOutgoingStreamResponse(services, stream, headers);
}

function normalizeSimulatedChatCompletionPayload(
  payload: JsonObject,
  model: string,
  conversationId: string,
  includeConversationId: boolean,
): JsonObject {
  const normalized: JsonObject = { ...payload };
  const choices = normalizeSimulatedChatChoices(normalized);
  normalized.choices = choices;

  if (!tryGetString(normalized, "id")) {
    normalized.id = `chatcmpl-${randomUUID().replaceAll("-", "")}`;
  }
  if (!tryGetString(normalized, "object")) {
    normalized.object = "chat.completion";
  }
  if (normalized.created === undefined) {
    normalized.created = nowUnix();
  }
  if (!tryGetString(normalized, "model")) {
    normalized.model = model;
  }
  if (includeConversationId) {
    normalized.conversation_id = conversationId;
  }
  return normalized;
}

function hasUsableSimulatedChatCompletionPayload(payload: JsonObject): boolean {
  const assistantResponse = tryBuildAssistantResponseFromChatCompletionPayload(payload);
  if (!assistantResponse) {
    return false;
  }
  if (assistantResponse.toolCalls.length > 0) {
    return true;
  }
  return Boolean(assistantResponse.content?.trim());
}

function shouldRetrySimulatedToollessChatPayload(
  options: WrapperOptions,
  request: ParsedOpenAiRequest,
  payload: JsonObject | null,
): boolean {
  if (!payload || !options.retrySimulatedToollessResponses) {
    return false;
  }
  if (request.tooling.tools.length === 0) {
    return false;
  }
  if (request.tooling.toolChoiceMode === ToolChoiceModes.None) {
    return false;
  }

  const assistantResponse =
    tryBuildAssistantResponseFromChatCompletionPayload(payload);
  if (!assistantResponse || assistantResponse.toolCalls.length > 0) {
    return false;
  }
  if (!assistantResponse.content?.trim()) {
    return false;
  }

  const choices = payload.choices;
  const firstChoice =
    Array.isArray(choices) && choices.length > 0 && isJsonObject(choices[0])
      ? choices[0]
      : null;
  const finishReason = firstChoice ? tryGetString(firstChoice, "finish_reason") : null;
  return !finishReason || finishReason.toLowerCase() === "stop";
}

function shouldRetrySimulatedInvalidChatToolPayload(
  payload: JsonObject | null,
): boolean {
  if (!payload) {
    return false;
  }
  const assistantResponse =
    tryBuildAssistantResponseFromChatCompletionPayload(payload);
  if (!assistantResponse || assistantResponse.toolCalls.length === 0) {
    return false;
  }
  for (const toolCall of assistantResponse.toolCalls) {
    if (isKnownInvalidSimulatedToolCall(toolCall.name, toolCall.argumentsJson)) {
      return true;
    }
  }
  return false;
}

function normalizeSimulatedChatChoices(payload: JsonObject): JsonObject[] {
  const explicitChoices = payload.choices;
  if (Array.isArray(explicitChoices) && explicitChoices.length > 0) {
    const normalized = explicitChoices.filter(isJsonObject).map(normalizeSimulatedChatChoice);
    if (normalized.length > 0) {
      return normalized;
    }
  }

  if (isJsonObject(payload.choice)) {
    return [normalizeSimulatedChatChoice(payload.choice)];
  }

  if (looksLikeChatChoice(payload)) {
    return [normalizeSimulatedChatChoice(payload)];
  }

  const outputBasedChoice = buildChoiceFromResponsesShape(payload);
  if (outputBasedChoice) {
    return [outputBasedChoice];
  }

  return [
    normalizeSimulatedChatChoice({
      index: 0,
      finish_reason: "stop",
      message: {
        role: "assistant",
        content: tryGetString(payload, "output_text") ?? tryGetString(payload, "text") ?? "",
      },
    }),
  ];
}

function looksLikeChatChoice(node: JsonObject): boolean {
  return (
    isJsonObject(node.message) ||
    isJsonObject(node.delta) ||
    node.finish_reason !== undefined ||
    node.index !== undefined
  );
}

function buildChoiceFromResponsesShape(payload: JsonObject): JsonObject | null {
  const output = getSimulatedResponsesOutputItems(payload);
  if (!Array.isArray(output) || output.length === 0) {
    return null;
  }

  const toolCalls: JsonObject[] = [];
  let messageText: string | null = null;

  for (const item of output) {
    if (!isJsonObject(item)) {
      continue;
    }
    if (isSimulatedResponsesFunctionCallOutputItem(item)) {
      const normalizedFunctionCallItem =
        normalizeSimulatedResponsesFunctionCallItem(item);
      const toolCall = normalizedFunctionCallItem
        ? normalizeSimulatedToolCall({
            id:
              tryGetString(normalizedFunctionCallItem, "call_id") ??
              tryGetString(normalizedFunctionCallItem, "id") ??
              `call_${randomUUID().replaceAll("-", "")}`,
            type: "function",
            function: {
              name:
                tryGetString(normalizedFunctionCallItem, "name") ?? "unknown_tool",
              arguments: normalizedFunctionCallItem.arguments,
            },
          })
        : null;
      if (toolCall) {
        toolCalls.push(toolCall);
      }
      continue;
    }
    const text = extractMessageOutputText(item);
    if (text) {
      messageText = text;
    }
  }

  if (toolCalls.length > 0) {
    return normalizeSimulatedChatChoice({
      index: 0,
      finish_reason: "tool_calls",
      message: {
        role: "assistant",
        content: null,
        tool_calls: toolCalls,
      },
    });
  }

  return normalizeSimulatedChatChoice({
    index: 0,
    finish_reason: "stop",
    message: {
      role: "assistant",
      content:
        messageText ??
        tryGetString(payload, "output_text") ??
        tryGetString(payload, "text") ??
        "",
    },
  });
}

function normalizeSimulatedChatChoice(choice: JsonObject): JsonObject {
  const normalized: JsonObject = {
    index: typeof choice.index === "number" ? choice.index : 0,
  };

  const messageNode = isJsonObject(choice.message)
    ? choice.message
    : isJsonObject(choice.delta)
      ? choice.delta
      : {};
  normalized.message = normalizeSimulatedAssistantMessage(messageNode);

  const finishReason = tryGetString(choice, "finish_reason");
  if (finishReason) {
    normalized.finish_reason = finishReason;
  } else {
    const toolCalls = (normalized.message as JsonObject).tool_calls;
    normalized.finish_reason =
      Array.isArray(toolCalls) && toolCalls.length > 0 ? "tool_calls" : "stop";
  }

  return normalized;
}

function normalizeSimulatedAssistantMessage(message: JsonObject): JsonObject {
  const normalized: JsonObject = { ...message };
  if (!tryGetString(normalized, "role")) {
    normalized.role = "assistant";
  }

  const rawToolCalls = Array.isArray(normalized.tool_calls)
    ? normalized.tool_calls
    : [];
  const toolCalls = rawToolCalls
    .filter(isJsonObject)
    .map(normalizeSimulatedToolCall)
    .filter(isJsonObject);
  if (toolCalls.length > 0) {
    normalized.tool_calls = toolCalls;
    if (normalized.content === undefined) {
      normalized.content = null;
    }
  } else if (normalized.tool_calls !== undefined) {
    delete normalized.tool_calls;
  }

  if (normalized.content === undefined || normalized.content === null) {
    if (toolCalls.length === 0) {
      normalized.content = "";
    }
  } else if (
    typeof normalized.content !== "string" &&
    !Array.isArray(normalized.content)
  ) {
    normalized.content = JSON.stringify(normalized.content);
  }

  return normalized;
}

function normalizeSimulatedToolCall(toolCall: JsonObject): JsonObject | null {
  const functionNode = isJsonObject(toolCall.function) ? toolCall.function : null;
  const functionName =
    tryGetString(functionNode, "name") ?? tryGetString(toolCall, "name");
  if (!functionName) {
    return null;
  }

  const argumentsNode =
    functionNode?.arguments !== undefined
      ? functionNode.arguments
      : toolCall.arguments;
  const argumentsText = normalizeSimulatedToolArguments(argumentsNode);

  return {
    id: tryGetString(toolCall, "id") ?? `call_${randomUUID().replaceAll("-", "")}`,
    type: "function",
    function: {
      name: functionName,
      arguments: argumentsText,
    },
  };
}

function normalizeSimulatedToolArguments(argumentsNode: unknown): string {
  if (typeof argumentsNode !== "string") {
    return JSON.stringify(argumentsNode ?? {});
  }

  const raw = argumentsNode.trim();
  if (!raw) {
    return "{}";
  }

  const repaired = sanitizeJsonControlCharsInsideStringLiterals(raw);
  for (const candidate of [raw, repaired]) {
    try {
      const parsed = JSON.parse(candidate) as unknown;
      return JSON.stringify(parsed);
    } catch {
      // try next candidate
    }
  }

  return JSON.stringify({ input: raw });
}

function sanitizeJsonControlCharsInsideStringLiterals(raw: string): string {
  let output = "";
  let inString = false;
  let escaped = false;

  for (let index = 0; index < raw.length; index++) {
    const ch = raw[index];
    if (!ch) {
      continue;
    }

    if (inString) {
      if (escaped) {
        output += ch;
        escaped = false;
        continue;
      }
      if (ch === "\\") {
        output += ch;
        escaped = true;
        continue;
      }
      if (ch === "\"") {
        output += ch;
        inString = false;
        continue;
      }
      if (ch === "\n") {
        output += "\\n";
        continue;
      }
      if (ch === "\r") {
        output += "\\r";
        continue;
      }
      if (ch === "\t") {
        output += "\\t";
        continue;
      }

      output += ch;
      continue;
    }

    if (ch === "\"") {
      inString = true;
      output += ch;
      continue;
    }

    output += ch;
  }

  return output;
}

function normalizeSimulatedResponsesPayload(
  payload: JsonObject,
  parsedRequest: ParsedResponsesRequest,
  conversationId: string,
  includeConversationId: boolean,
): { responseId: string; responseBody: JsonObject } {
  const responseBody: JsonObject = { ...payload };
  const responseId = tryGetString(responseBody, "id") ?? createOpenAiResponseId();
  responseBody.id = responseId;
  if (!tryGetString(responseBody, "object")) {
    responseBody.object = "response";
  }
  if (responseBody.created_at === undefined) {
    responseBody.created_at = nowUnix();
  }
  if (!tryGetString(responseBody, "status")) {
    responseBody.status = "completed";
  }
  if (!tryGetString(responseBody, "model")) {
    responseBody.model = parsedRequest.base.model;
  }
  const rawOutputItems = getSimulatedResponsesOutputItems(responseBody);
  if (!Array.isArray(rawOutputItems)) {
    const outputText =
      tryGetString(responseBody, "output_text") ??
      tryGetString(responseBody, "text") ??
      extractTextFromSimulatedChoices(responseBody) ??
      "";
    responseBody.output = [
      buildMessageOutputItem(createOpenAiOutputItemId("msg"), outputText, "completed"),
    ];
  } else {
    responseBody.output = normalizeSimulatedResponseOutputItems(rawOutputItems);
  }
  if (responseBody.outputs !== undefined) {
    delete responseBody.outputs;
  }
  if (!tryGetString(responseBody, "output_text")?.trim()) {
    responseBody.output_text = extractResponseOutputText(
      Array.isArray(responseBody.output) ? responseBody.output : [],
    );
  }
  if (includeConversationId) {
    responseBody.conversation = conversationId;
    responseBody.conversation_id = conversationId;
  }

  return { responseId, responseBody };
}

function hasUsableSimulatedResponsesPayload(responseBody: JsonObject): boolean {
  const outputText = tryGetString(responseBody, "output_text");
  if (outputText?.trim()) {
    return true;
  }

  const output = responseBody.output;
  if (!Array.isArray(output)) {
    return false;
  }
  for (const item of output) {
    if (!isJsonObject(item)) {
      continue;
    }
    const type = (tryGetString(item, "type") ?? "").toLowerCase();
    if (type === "function_call") {
      return true;
    }
    if (extractMessageOutputText(item).trim()) {
      return true;
    }
  }
  return false;
}

function shouldRetrySimulatedToollessResponsesPayload(
  options: WrapperOptions,
  request: ParsedOpenAiRequest,
  responseBody: JsonObject | null,
): boolean {
  if (!responseBody || !options.retrySimulatedToollessResponses) {
    return false;
  }
  if (request.tooling.tools.length === 0) {
    return false;
  }
  if (request.tooling.toolChoiceMode === ToolChoiceModes.None) {
    return false;
  }

  const output = responseBody.output;
  if (!Array.isArray(output)) {
    return false;
  }

  let hasFunctionCall = false;
  let hasMessageText = false;
  for (const item of output) {
    if (!isJsonObject(item)) {
      continue;
    }
    const type = (tryGetString(item, "type") ?? "").toLowerCase();
    if (type === "function_call") {
      hasFunctionCall = true;
      break;
    }
    if (extractMessageOutputText(item).trim()) {
      hasMessageText = true;
    }
  }

  if (hasFunctionCall) {
    return false;
  }
  if (hasMessageText) {
    return true;
  }

  const outputText = tryGetString(responseBody, "output_text");
  return Boolean(outputText?.trim());
}

function shouldRetrySimulatedInvalidResponsesToolPayload(
  responseBody: JsonObject | null,
): boolean {
  if (!responseBody) {
    return false;
  }
  const output = responseBody.output;
  if (!Array.isArray(output)) {
    return false;
  }

  for (const item of output) {
    if (!isJsonObject(item)) {
      continue;
    }
    const type = (tryGetString(item, "type") ?? "").toLowerCase();
    if (type !== "function_call") {
      continue;
    }
    const name = tryGetString(item, "name") ?? "";
    const argsNode = item.arguments;
    const argumentsJson =
      typeof argsNode === "string" ? argsNode : JSON.stringify(argsNode ?? {});
    if (isKnownInvalidSimulatedToolCall(name, argumentsJson)) {
      return true;
    }
  }

  return false;
}

function isKnownInvalidSimulatedToolCall(
  name: string,
  argumentsJson: string,
): boolean {
  const normalizedName = name.trim().toLowerCase();
  if (normalizedName !== "apply_diff") {
    return false;
  }

  const args = tryParseJsonObject(argumentsJson);
  if (!args) {
    return false;
  }
  const diff = tryGetString(args, "diff");
  if (!diff) {
    return false;
  }
  return hasEmptySearchBlock(diff);
}

function hasEmptySearchBlock(diff: string): boolean {
  const normalized = diff.replace(/\r\n/g, "\n");
  const blockPattern = /<<<<<<< SEARCH[\s\S]*?-------\n([\s\S]*?)\n=======/g;
  for (const match of normalized.matchAll(blockPattern)) {
    const searchBody = match[1] ?? "";
    if (!searchBody.trim()) {
      return true;
    }
  }
  return false;
}

async function buildSimulatedChatStreamResponse(
  services: Services,
  model: string,
  conversationId: string,
  payload: JsonObject,
  includeConversationId: boolean,
  headers: Headers,
  trace: TraceContext | null,
): Promise<Response> {
  const normalized = normalizeSimulatedChatCompletionPayload(
    payload,
    model,
    conversationId,
    includeConversationId,
  );
  const assistantResponse =
    tryBuildAssistantResponseFromChatCompletionPayload(normalized);
  if (!assistantResponse) {
    traceError(
      services,
      trace,
      {
        message: "Simulated chat payload was not a valid chat completion object.",
        type: "api_error",
        param: null,
        code: "invalid_simulated_payload",
      },
      502,
    );
    return writeOpenAiError(
      services,
      502,
      "Simulated chat payload was not a valid chat completion object.",
      "api_error",
      "invalid_simulated_payload",
    );
  }

  return buildAssistantStreamResponse(
    services,
    model,
    conversationId,
    assistantResponse,
    includeConversationId,
    headers,
    trace,
  );
}

async function buildSimulatedResponsesStreamResponse(
  services: Services,
  parsedRequest: ParsedResponsesRequest,
  conversationId: string,
  payload: JsonObject,
  headers: Headers,
  trace: TraceContext | null,
): Promise<Response> {
  const includeConversationId = services.options.includeConversationIdInResponseBody;
  const normalized = normalizeSimulatedResponsesPayload(
    payload,
    parsedRequest,
    conversationId,
    includeConversationId,
  );
  const responseBody = normalized.responseBody;
  const responseId = createOpenAiResponseId();
  responseBody.id = responseId;
  const rawOutputItems = Array.isArray(responseBody.output)
    ? responseBody.output.filter(isJsonObject)
    : [];
  const resolvedModel =
    tryGetString(responseBody, "model") ?? parsedRequest.base.model;
  const createdAt = resolveResponseCreatedAt(responseBody.created_at);
  const finalizedOutputItems: JsonObject[] = rawOutputItems.map((item, index) => {
    const itemId = tryGetString(item, "id") ?? createOpenAiOutputItemId("out");
    const itemType = (tryGetString(item, "type") ?? "").toLowerCase();
    if (itemType === "message") {
      return buildMessageOutputItem(
        itemId,
        extractMessageOutputText(item),
        "completed",
      );
    }
    const normalizedItem: JsonObject = { ...item };
    normalizedItem.id = itemId;
    if (!tryGetString(normalizedItem, "status")) {
      normalizedItem.status = "completed";
    }
    if (!tryGetString(normalizedItem, "type")) {
      normalizedItem.type = "output";
    }
    return normalizedItem;
  });
  const responseConversationId = includeConversationId ? conversationId : null;

  const stream = new ReadableStream<Uint8Array>({
    start(controller) {
      const encoder = new TextEncoder();
      const writeDataEvent = (event: JsonObject) => {
        enqueueSseJsonEvent(controller, encoder, event);
      };

      const inProgress = buildOpenAiResponseObject(
        responseId,
        createdAt,
        resolvedModel,
        "in_progress",
        [],
        parsedRequest,
        responseConversationId,
      );
      writeDataEvent(buildResponseCreatedEvent(inProgress));
      writeDataEvent(buildResponseInProgressEvent(inProgress));

      for (let index = 0; index < finalizedOutputItems.length; index++) {
        const item = finalizedOutputItems[index];
        const itemId = tryGetString(item, "id") ?? createOpenAiOutputItemId("out");
        const itemType = (tryGetString(item, "type") ?? "").toLowerCase();
        if (itemType === "message") {
          const text = extractMessageOutputText(item);
          writeDataEvent(
            buildResponseOutputItemAddedEvent(
              responseId,
              index,
              buildMessageOutputItem(itemId, "", "in_progress"),
            ),
          );
          writeDataEvent(
            buildResponseContentPartAddedEvent(
              responseId,
              index,
              itemId,
              { type: "output_text", text: "" },
            ),
          );
          if (text) {
            writeDataEvent(
              buildResponseOutputTextDeltaEvent(responseId, index, itemId, text),
            );
          }
          writeDataEvent(
            buildResponseOutputTextDoneEvent(responseId, index, itemId, text),
          );
          writeDataEvent(
            buildResponseContentPartDoneEvent(
              responseId,
              index,
              itemId,
              { type: "output_text", text },
            ),
          );
          writeDataEvent(
            buildResponseOutputItemDoneEvent(
              responseId,
              index,
              buildMessageOutputItem(itemId, text, "completed"),
            ),
          );
          continue;
        }

        writeDataEvent(buildResponseOutputItemAddedEvent(responseId, index, item));
        writeDataEvent(buildResponseOutputItemDoneEvent(responseId, index, item));
      }

      const completed = buildOpenAiResponseObject(
        responseId,
        createdAt,
        resolvedModel,
        "completed",
        finalizedOutputItems,
        parsedRequest,
        responseConversationId,
      );
      writeDataEvent(buildResponseCompletedEvent(completed));
      services.responseStore.set(responseId, completed, conversationId);
      tracePane2(services, trace, completed, 200);
      traceComplete(services, trace, 200);
      enqueueSseDoneEvent(controller, encoder);
      controller.close();
    },
  });

  headers.set("content-type", "text/event-stream");
  headers.set("cache-control", "no-cache");
  headers.set("connection", "keep-alive");
  headers.set("x-m365-conversation-id", conversationId);
  await services.debugLogger.logOutgoingResponse(200, headers.entries(), null);
  return finalizeOutgoingStreamResponse(services, stream, headers);
}

function resolveResponseCreatedAt(createdAt: unknown): number {
  if (typeof createdAt === "number" && Number.isFinite(createdAt)) {
    return Math.trunc(createdAt);
  }
  if (typeof createdAt === "string") {
    const parsed = Number.parseInt(createdAt, 10);
    if (Number.isFinite(parsed)) {
      return parsed;
    }
  }
  return nowUnix();
}

function extractResponseOutputText(outputItems: unknown[]): string {
  const textParts: string[] = [];
  for (const item of outputItems) {
    if (!isJsonObject(item)) {
      continue;
    }
    const text = extractMessageOutputText(item);
    if (text) {
      textParts.push(text);
    }
  }
  return textParts.join("");
}

function computeResponsesRequestHash(
  rawRequestText: string | null,
  requestJson: JsonObject,
  transport: string,
): string {
  const payloadText =
    rawRequestText?.trim() && rawRequestText.trim().length > 0
      ? rawRequestText
      : JSON.stringify(requestJson);
  return createHash("sha256")
    .update(transport)
    .update("\n")
    .update(payloadText)
    .digest("hex");
}

function computeResponsesReplayLoopHash(
  requestJson: JsonObject,
  transport: string,
): string {
  const canonicalRequestJson = buildReplayLoopCanonicalRequestJson(requestJson);
  return computeResponsesRequestHash(null, canonicalRequestJson, transport);
}

function buildReplayLoopCanonicalRequestJson(requestJson: JsonObject): JsonObject {
  const inputItems = Array.isArray(requestJson.input) ? requestJson.input : null;
  if (!inputItems || inputItems.length === 0) {
    return requestJson;
  }

  let lastNonAssistantIndex = inputItems.length - 1;
  while (
    lastNonAssistantIndex >= 0 &&
    isAssistantInputItem(inputItems[lastNonAssistantIndex])
  ) {
    lastNonAssistantIndex -= 1;
  }

  if (
    lastNonAssistantIndex < 0 ||
    lastNonAssistantIndex >= inputItems.length - 1
  ) {
    return requestJson;
  }

  return {
    ...requestJson,
    input: inputItems.slice(0, lastNonAssistantIndex + 1),
  };
}

function isAssistantInputItem(inputItem: unknown): boolean {
  if (!isJsonObject(inputItem)) {
    return false;
  }
  return (tryGetString(inputItem, "role") ?? "").toLowerCase() === "assistant";
}

function rememberResponsesReplayHashes(
  responseStore: ResponseStore,
  requestHash: string,
  replayLoopHash: string,
  conversationId: string | null,
): void {
  responseStore.rememberRequestHash(requestHash, conversationId);
  if (replayLoopHash !== requestHash) {
    responseStore.rememberRequestHash(replayLoopHash, conversationId);
  }
}

function buildReplayResponseIdFromHash(requestHash: string): string {
  return `resp_replay_${requestHash.slice(0, 24)}`;
}

function resolveSuppressedResponsesConversationId(
  request: Request,
  requestJson: JsonObject,
  parsedRequest: ParsedResponsesRequest,
  selectedTransport: string,
  conversationStore: ConversationStore,
  responseStore: ResponseStore,
): string | null {
  const selection = selectConversation(
    request,
    requestJson,
    parsedRequest.base.userKey,
  );
  if (selection.conversationId) {
    return selection.conversationId;
  }
  if (parsedRequest.previousResponseId) {
    const linked = responseStore.tryGetConversationLink(
      parsedRequest.previousResponseId,
    );
    if (linked) {
      return linked;
    }
  }
  const scopedConversationKey = scopeConversationKey(
    selection.conversationKey,
    selectedTransport,
  );
  if (!scopedConversationKey) {
    return null;
  }
  return conversationStore.tryGet(scopedConversationKey);
}

function detectAssistantTailReplayWithoutNewUserTurn(
  inputItems: unknown[],
): string | null {
  if (!Array.isArray(inputItems) || inputItems.length === 0) {
    return null;
  }

  const latestAssistantText = extractAssistantInputItemText(
    inputItems[inputItems.length - 1],
  );
  if (!latestAssistantText) {
    return null;
  }

  for (let index = inputItems.length - 2; index >= 0; index -= 1) {
    const item = inputItems[index];
    if (!isJsonObject(item)) {
      continue;
    }
    const role = (tryGetString(item, "role") ?? "").toLowerCase();
    if (role === "user") {
      return latestAssistantText;
    }
  }
  return null;
}

function extractAssistantInputItemText(inputItem: unknown): string | null {
  if (!isJsonObject(inputItem)) {
    return null;
  }

  const role = (tryGetString(inputItem, "role") ?? "").toLowerCase();
  if (role !== "assistant") {
    return null;
  }

  const content = inputItem.content;
  if (typeof content === "string") {
    const text = content.trim();
    return text ? text : null;
  }

  if (Array.isArray(content)) {
    const textParts: string[] = [];
    for (const part of content) {
      if (typeof part === "string") {
        if (part.trim()) {
          textParts.push(part);
        }
        continue;
      }
      if (!isJsonObject(part)) {
        continue;
      }
      const partType = (tryGetString(part, "type") ?? "").toLowerCase();
      if (
        partType &&
        partType !== "output_text" &&
        partType !== "text" &&
        partType !== "input_text"
      ) {
        continue;
      }
      const text =
        tryGetString(part, "text") ??
        tryGetString(part, "output_text") ??
        tryGetString(part, "input_text");
      if (text?.trim()) {
        textParts.push(text);
      }
    }
    const combined = textParts.join("").trim();
    return combined ? combined : null;
  }

  const fallback =
    tryGetString(inputItem, "output_text") ?? tryGetString(inputItem, "text");
  if (!fallback) {
    return null;
  }
  const normalized = fallback.trim();
  return normalized ? normalized : null;
}

function normalizeSimulatedResponseOutputItems(outputItems: unknown[]): JsonObject[] {
  const normalized: JsonObject[] = [];

  for (const item of outputItems) {
    if (typeof item === "string") {
      const text = item.trim();
      if (!text) {
        continue;
      }
      normalized.push(
        buildMessageOutputItem(createOpenAiOutputItemId("msg"), text, "completed"),
      );
      continue;
    }

    if (!isJsonObject(item)) {
      continue;
    }

    if (isSimulatedResponsesFunctionCallOutputItem(item)) {
      normalized.push(
        normalizeSimulatedResponsesFunctionCallItem(item) ?? item,
      );
      continue;
    }

    const text = extractMessageOutputText(item).trim();
    if (!text) {
      normalized.push(item);
      continue;
    }

    const id = tryGetString(item, "id") ?? createOpenAiOutputItemId("msg");
    const status = tryGetString(item, "status") ?? "completed";
    normalized.push(buildMessageOutputItem(id, text, status));
  }

  return normalized;
}

function getSimulatedResponsesOutputItems(payload: JsonObject): unknown[] | null {
  if (Array.isArray(payload.output)) {
    return payload.output;
  }
  if (Array.isArray(payload.outputs)) {
    return payload.outputs;
  }
  return null;
}

function isSimulatedResponsesFunctionCallOutputItem(item: JsonObject): boolean {
  return (
    (tryGetString(item, "type") ?? "").toLowerCase() === "function_call" ||
    isJsonObject(item.function_call)
  );
}

function normalizeSimulatedResponsesFunctionCallItem(
  item: JsonObject,
): JsonObject | null {
  const functionCallNode = isJsonObject(item.function_call) ? item.function_call : null;
  const name =
    tryGetString(item, "name") ?? tryGetString(functionCallNode, "name");
  if (!name) {
    return null;
  }

  const itemId = tryGetString(item, "id") ?? createOpenAiOutputItemId("fc");
  return {
    id: itemId,
    type: "function_call",
    status: tryGetString(item, "status") ?? "completed",
    call_id:
      tryGetString(item, "call_id") ??
      tryGetString(item, "tool_call_id") ??
      tryGetString(functionCallNode, "call_id") ??
      itemId,
    name,
    arguments: normalizeSimulatedToolArguments(
      functionCallNode?.arguments !== undefined
        ? functionCallNode.arguments
        : item.arguments,
    ),
  };
}

function extractTextFromSimulatedChoices(payload: JsonObject): string | null {
  const choices = payload.choices;
  if (!Array.isArray(choices)) {
    return null;
  }

  for (const choice of choices) {
    if (!isJsonObject(choice)) {
      continue;
    }
    const messageNode = choice.message;
    if (!isJsonObject(messageNode)) {
      continue;
    }
    const text = extractMessageOutputText(messageNode).trim();
    if (text) {
      return text;
    }
  }

  return null;
}

function extractMessageOutputText(outputItem: JsonObject): string {
  const type = (tryGetString(outputItem, "type") ?? "").toLowerCase();
  if (type === "function_call" || isJsonObject(outputItem.function_call)) {
    return "";
  }

  if (type === "output_text" || type === "text") {
    return (
      tryGetString(outputItem, "text") ?? tryGetString(outputItem, "output_text") ?? ""
    );
  }

  const directText = tryGetString(outputItem, "output_text");
  if (directText) {
    return directText;
  }

  const content = outputItem.content;
  if (typeof content === "string") {
    return content;
  }

  if (!Array.isArray(content)) {
    return tryGetString(outputItem, "text") ?? "";
  }

  const textParts: string[] = [];
  for (const part of content) {
    if (typeof part === "string") {
      if (part) {
        textParts.push(part);
      }
      continue;
    }
    if (!isJsonObject(part)) {
      continue;
    }
    const partType = (tryGetString(part, "type") ?? "").toLowerCase();
    if (partType && partType !== "output_text" && partType !== "text") {
      continue;
    }
    const text = tryGetString(part, "text") ?? tryGetString(part, "output_text");
    if (text) {
      textParts.push(text);
    }
  }
  return textParts.join("");
}

async function transformGraphStreamToResponses(
  services: Services,
  graphResponse: Response,
  parsedRequest: ParsedResponsesRequest,
  initialConversationId: string,
  scopedConversationKey: string | null,
  headers: Headers,
  trace: TraceContext | null,
): Promise<Response> {
  const includeConversationId = services.options.includeConversationIdInResponseBody;
  const responseId = createOpenAiResponseId();
  const createdAt = nowUnix();
  const messageItemId = createOpenAiOutputItemId("msg");
  let conversationId = initialConversationId;
  let emittedContent = "";
  const upstreamItems: JsonValue[] = [];

  const stream = new ReadableStream<Uint8Array>({
    start: async (controller) => {
      const encoder = new TextEncoder();
      const writeDataEvent = (event: JsonObject) => {
        enqueueSseJsonEvent(controller, encoder, event);
      };
      const writeError = (message: string, code: string) => {
        controller.enqueue(
          encoder.encode(
            `event: error\ndata: ${JSON.stringify({
              error: {
                message,
                type: "api_error",
                param: null,
                code,
              },
            })}\n\n`,
          ),
        );
      };

      const inProgress = buildOpenAiResponseObject(
        responseId,
        createdAt,
        parsedRequest.base.model,
        "in_progress",
        [],
        parsedRequest,
        includeConversationId ? conversationId : null,
      );
      writeDataEvent(buildResponseCreatedEvent(inProgress));
      writeDataEvent(buildResponseInProgressEvent(inProgress));
      writeDataEvent(
        buildResponseOutputItemAddedEvent(
          responseId,
          0,
          buildMessageOutputItem(messageItemId, "", "in_progress"),
        ),
      );
      writeDataEvent(
        buildResponseContentPartAddedEvent(
          responseId,
          0,
          messageItemId,
          { type: "output_text", text: "" },
        ),
      );

      try {
        if (graphResponse.body) {
          for await (const event of readSseEvents(graphResponse.body)) {
            const data = event.data.trim();
            if (!data) {
              continue;
            }
            if (data.toLowerCase() === "[done]") {
              break;
            }
            upstreamItems.push(tryParseJsonObject(data) ?? { rawText: data });
            tracePane4(
              services,
              trace,
              buildUpstreamStreamCapture("sse", upstreamItems),
              graphResponse.status,
            );

            const streamConversationId =
              extractCopilotConversationIdFromStream(data);
            if (streamConversationId) {
              conversationId = streamConversationId;
              if (scopedConversationKey) {
                services.conversationStore.set(
                  scopedConversationKey,
                  conversationId,
                );
              }
            }

            const latestAssistantText = extractCopilotAssistantTextFromStreamData(
              data,
              parsedRequest.base.promptText,
            );
            if (!latestAssistantText) {
              continue;
            }

            const delta = computeTrailingDelta(emittedContent, latestAssistantText);
            if (!delta) {
              continue;
            }

            emittedContent += delta;
            writeDataEvent(
              buildResponseOutputTextDeltaEvent(responseId, 0, messageItemId, delta),
            );
          }
        }

        writeDataEvent(
          buildResponseOutputTextDoneEvent(
            responseId,
            0,
            messageItemId,
            emittedContent,
          ),
        );
        writeDataEvent(
          buildResponseContentPartDoneEvent(
            responseId,
            0,
            messageItemId,
            { type: "output_text", text: emittedContent },
          ),
        );
        const outputItem = buildMessageOutputItem(
          messageItemId,
          emittedContent,
          "completed",
        );
        writeDataEvent(buildResponseOutputItemDoneEvent(responseId, 0, outputItem));

        const completed = buildOpenAiResponseObject(
          responseId,
          createdAt,
          parsedRequest.base.model,
          "completed",
          [outputItem],
          parsedRequest,
          includeConversationId ? conversationId : null,
        );
        writeDataEvent(buildResponseCompletedEvent(completed));
        services.responseStore.set(responseId, completed, conversationId);
        tracePane4(
          services,
          trace,
          buildUpstreamStreamCapture("sse", upstreamItems),
          graphResponse.status,
        );
        tracePane2(services, trace, completed, 200);
        traceComplete(services, trace, 200);
      } catch (error) {
        tracePane4(
          services,
          trace,
          buildUpstreamStreamCapture("sse", upstreamItems),
          graphResponse.status,
        );
        traceError(
          services,
          trace,
          {
            message: `Microsoft Graph chatOverStream request failed. ${String(error)}`,
            type: "api_error",
            param: null,
            code: "graph_error",
          },
          500,
        );
        writeError(
          `Microsoft Graph chatOverStream request failed. ${String(error)}`,
          "graph_error",
        );
      } finally {
        enqueueSseDoneEvent(controller, encoder);
        controller.close();
      }
    },
  });

  headers.set("content-type", "text/event-stream");
  headers.set("cache-control", "no-cache");
  headers.set("connection", "keep-alive");
  headers.set("x-m365-conversation-id", initialConversationId);
  await services.debugLogger.logOutgoingResponse(200, headers.entries(), null);
  return finalizeOutgoingStreamResponse(services, stream, headers);
}

async function streamSubstrateAsResponses(
  services: Services,
  authorizationHeader: string,
  initialConversationId: string,
  parsedRequest: ParsedResponsesRequest,
  createdConversation: boolean,
  scopedConversationKey: string | null,
  headers: Headers,
  trace: TraceContext | null,
): Promise<Response> {
  const includeConversationId = services.options.includeConversationIdInResponseBody;
  const responseId = createOpenAiResponseId();
  const createdAt = nowUnix();
  const messageItemId = createOpenAiOutputItemId("msg");
  let conversationId = initialConversationId;
  let emitted = "";

  const stream = new ReadableStream<Uint8Array>({
    start: async (controller) => {
      const encoder = new TextEncoder();
      const writeDataEvent = (event: JsonObject) => {
        enqueueSseJsonEvent(controller, encoder, event);
      };
      const writeError = (message: string, code: string) => {
        controller.enqueue(
          encoder.encode(
            `event: error\ndata: ${JSON.stringify({
              error: {
                message,
                type: "api_error",
                param: null,
                code,
              },
            })}\n\n`,
          ),
        );
      };

      const inProgress = buildOpenAiResponseObject(
        responseId,
        createdAt,
        parsedRequest.base.model,
        "in_progress",
        [],
        parsedRequest,
        includeConversationId ? conversationId : null,
      );
      writeDataEvent(buildResponseCreatedEvent(inProgress));
      writeDataEvent(buildResponseInProgressEvent(inProgress));
      writeDataEvent(
        buildResponseOutputItemAddedEvent(
          responseId,
          0,
          buildMessageOutputItem(messageItemId, "", "in_progress"),
        ),
      );
      writeDataEvent(
        buildResponseContentPartAddedEvent(
          responseId,
          0,
          messageItemId,
          { type: "output_text", text: "" },
        ),
      );

      const substrateResponse = await services.substrateClient.chatStream(
        authorizationHeader,
        conversationId,
        parsedRequest.base,
        createdConversation,
        async (update) => {
          traceSubstrateStreamUpdate(services, trace, update);
          if (update.conversationId) {
            conversationId = update.conversationId;
            if (scopedConversationKey) {
              services.conversationStore.set(scopedConversationKey, conversationId);
            }
          }
          if (!update.deltaText) {
            return;
          }
          emitted += update.deltaText;
          writeDataEvent(
            buildResponseOutputTextDeltaEvent(
              responseId,
              0,
              messageItemId,
              update.deltaText,
            ),
          );
        },
      );
      tracePane3(services, trace, substrateResponse.upstreamRequestPayload ?? null);
      tracePane4(
        services,
        trace,
        substrateResponse.upstreamResponsePayload ?? null,
        substrateResponse.statusCode,
      );

      if (!substrateResponse.isSuccess) {
        const details = extractGraphErrorMessage(substrateResponse.rawBody);
        writeError(
          details
            ? `Substrate chat request failed. ${details}`
            : "Substrate chat request failed.",
          "substrate_error",
        );
        traceError(
          services,
          trace,
          {
            message: details
              ? `Substrate chat request failed. ${details}`
              : "Substrate chat request failed.",
            type: "api_error",
            param: null,
            code: "substrate_error",
          },
          substrateResponse.statusCode,
        );
        enqueueSseDoneEvent(controller, encoder);
        controller.close();
        return;
      }

      if (substrateResponse.conversationId) {
        conversationId = substrateResponse.conversationId;
      }
      const assistantText =
        substrateResponse.assistantText ??
        extractCopilotAssistantText(
          substrateResponse.responseJson,
          parsedRequest.base.promptText,
        ) ??
        "";

      const trailing = computeTrailingDelta(emitted, assistantText);
      if (trailing) {
        emitted += trailing;
        writeDataEvent(
          buildResponseOutputTextDeltaEvent(responseId, 0, messageItemId, trailing),
        );
      }

      writeDataEvent(
        buildResponseOutputTextDoneEvent(responseId, 0, messageItemId, emitted),
      );
      writeDataEvent(
        buildResponseContentPartDoneEvent(
          responseId,
          0,
          messageItemId,
          { type: "output_text", text: emitted },
        ),
      );
      const outputItem = buildMessageOutputItem(messageItemId, emitted, "completed");
      writeDataEvent(buildResponseOutputItemDoneEvent(responseId, 0, outputItem));

      const completed = buildOpenAiResponseObject(
        responseId,
        createdAt,
        parsedRequest.base.model,
        "completed",
        [outputItem],
        parsedRequest,
        includeConversationId ? conversationId : null,
      );
      writeDataEvent(buildResponseCompletedEvent(completed));
      services.responseStore.set(responseId, completed, conversationId);
      tracePane2(services, trace, completed, 200);
      traceComplete(services, trace, 200);
      enqueueSseDoneEvent(controller, encoder);
      controller.close();
    },
  });

  headers.set("content-type", "text/event-stream");
  headers.set("cache-control", "no-cache");
  headers.set("connection", "keep-alive");
  headers.set("x-m365-conversation-id", initialConversationId);
  await services.debugLogger.logOutgoingResponse(200, headers.entries(), null);
  return finalizeOutgoingStreamResponse(services, stream, headers);
}

async function resolveAuthorizationHeader(
  request: Request,
  services: Services,
): Promise<string | null> {
  return services.tokenProvider.resolveAuthorizationHeader(
    request.headers.get("authorization"),
  );
}

async function tryWriteStrictToolOutputError(
  services: Services,
  assistantResponse: OpenAiAssistantResponse,
): Promise<Response | null> {
  const message = assistantResponse.strictToolErrorMessage;
  if (!message) {
    return null;
  }
  return writeOpenAiError(
    services,
    400,
    message,
    "invalid_request_error",
    "invalid_tool_output",
  );
}

function shouldRetrySubstrateNoAssistantContent(
  transport: string,
  createdConversation: boolean,
  chatResult: ChatResult,
): boolean {
  if (transport !== TransportNames.Substrate || !createdConversation) {
    return false;
  }
  if (chatResult.isSuccess || chatResult.statusCode !== 502) {
    return false;
  }
  const message = extractGraphErrorMessage(chatResult.rawBody);
  if (!message) {
    return false;
  }
  return message
    .toLowerCase()
    .includes("substrate chat returned no assistant content");
}

function shouldRetryStrictToolOutput(
  transport: string,
  request: ParsedOpenAiRequest,
  assistantResponse: OpenAiAssistantResponse,
): boolean {
  if (transport !== TransportNames.Substrate) {
    return false;
  }
  if (!assistantResponse.strictToolErrorMessage) {
    return false;
  }
  return (
    request.tooling.toolChoiceMode === ToolChoiceModes.Required ||
    request.tooling.toolChoiceMode === ToolChoiceModes.Function
  );
}

function clampListLimit(raw: string | null): number {
  if (!raw || !raw.trim()) {
    return 20;
  }
  const parsed = Number.parseInt(raw, 10);
  if (!Number.isFinite(parsed) || parsed <= 0) {
    return 20;
  }
  return parsed > 100 ? 100 : parsed;
}

async function writeOpenAiError(
  services: Services,
  statusCode: number,
  message: string,
  type: string,
  code: string,
): Promise<Response> {
  const body = JSON.stringify({
    error: {
      message,
      type,
      param: null,
      code,
    },
  });
  const headers = new Headers({ "content-type": "application/json" });
  await services.debugLogger.logOutgoingResponse(
    statusCode,
    headers.entries(),
    body,
  );
  return new Response(body, { status: statusCode, headers });
}

async function writeFromUpstreamFailure(
  services: Services,
  statusCode: number,
  responseBody: string | null,
  fallbackMessage: string,
  errorCode: string,
): Promise<Response> {
  const { statusCode: normalizedStatusCode, message } =
    summarizeUpstreamFailure(statusCode, responseBody, fallbackMessage);
  return writeOpenAiError(
    services,
    normalizedStatusCode,
    message,
    "api_error",
    errorCode,
  );
}

function enqueueSseJsonEvent(
  controller: ReadableStreamDefaultController<Uint8Array>,
  encoder: TextEncoder,
  event: JsonObject,
): void {
  const eventName = tryGetString(event, "type");
  if (eventName) {
    controller.enqueue(encoder.encode(`event: ${eventName}\n`));
  }
  controller.enqueue(encoder.encode(`data: ${JSON.stringify(event)}\n\n`));
}

function enqueueSseDoneEvent(
  controller: ReadableStreamDefaultController<Uint8Array>,
  encoder: TextEncoder,
): void {
  controller.enqueue(encoder.encode("event: done\ndata: [DONE]\n\n"));
}

const MaxLoggedOutgoingStreamBodyChars = 1_000_000;

function finalizeOutgoingStreamResponse(
  services: Services,
  stream: ReadableStream<Uint8Array>,
  headers: Headers,
): Response {
  const response = new Response(stream, { status: 200, headers });
  if (services.options.logStreamingResponseBody !== true || !response.body) {
    return response;
  }

  const contentType = (response.headers.get("content-type") ?? "").toLowerCase();
  if (!contentType.includes("text/event-stream")) {
    return response;
  }

  const [clientBody, loggerBody] = response.body.tee();
  const responseHeaders = new Headers(response.headers);
  void captureOutgoingStreamBodyForDebug(
    services,
    response.status,
    responseHeaders,
    loggerBody,
  );

  return new Response(clientBody, {
    status: response.status,
    statusText: response.statusText,
    headers: responseHeaders,
  });
}

async function captureOutgoingStreamBodyForDebug(
  services: Services,
  statusCode: number,
  headers: Headers,
  body: ReadableStream<Uint8Array>,
): Promise<void> {
  try {
    const captured = await readStreamTextWithLimit(
      body,
      MaxLoggedOutgoingStreamBodyChars,
    );
    await services.debugLogger.logOutgoingStreamBody(
      statusCode,
      headers.entries(),
      captured,
    );
  } catch {
    // Best-effort debug capture should never impact the live stream.
  }
}

async function readStreamTextWithLimit(
  stream: ReadableStream<Uint8Array>,
  maxChars: number,
): Promise<string> {
  const reader = stream.getReader();
  const decoder = new TextDecoder();
  let text = "";
  let truncated = false;

  try {
    while (true) {
      const { value, done } = await reader.read();
      if (done) {
        break;
      }
      if (!value) {
        continue;
      }

      const chunk = decoder.decode(value, { stream: true });
      if (truncated || !chunk) {
        continue;
      }

      if (text.length + chunk.length <= maxChars) {
        text += chunk;
        continue;
      }

      const remaining = Math.max(maxChars - text.length, 0);
      if (remaining > 0) {
        text += chunk.slice(0, remaining);
      }
      truncated = true;
    }

    const finalChunk = decoder.decode();
    if (!truncated && finalChunk) {
      if (text.length + finalChunk.length <= maxChars) {
        text += finalChunk;
      } else {
        const remaining = Math.max(maxChars - text.length, 0);
        if (remaining > 0) {
          text += finalChunk.slice(0, remaining);
        }
        truncated = true;
      }
    }
  } finally {
    reader.releaseLock();
  }

  if (truncated) {
    return `${text}\n\n[stream body truncated in debug log]`;
  }
  return text;
}

async function buildAssistantStreamResponse(
  services: Services,
  model: string,
  conversationId: string,
  assistantResponse: ReturnType<typeof buildAssistantResponse>,
  includeConversationId: boolean,
  headers: Headers,
  trace: TraceContext | null,
): Promise<Response> {
  const completionId = `chatcmpl-${randomUUID().replaceAll("-", "")}`;
  const created = nowUnix();
  const responseConversationId = includeConversationId ? conversationId : null;
  const stream = new ReadableStream<Uint8Array>({
    start(controller) {
      const encoder = new TextEncoder();
      const writeChunk = (
        role: string | null,
        content: string | null,
        finishReason: string | null,
        toolCalls?: unknown,
      ) => {
        const chunk = buildChatCompletionChunk(
          completionId,
          created,
          model,
          role,
          content,
          finishReason,
          responseConversationId,
          toolCalls as never,
        );
        controller.enqueue(
          encoder.encode(`data: ${JSON.stringify(chunk)}\n\n`),
        );
      };

      writeChunk("assistant", null, null);
      if (assistantResponse.toolCalls.length > 0) {
        writeChunk(
          null,
          null,
          null,
          buildToolCallsDelta(assistantResponse.toolCalls),
        );
      } else if (assistantResponse.content?.trim()) {
        writeChunk(null, assistantResponse.content, null);
      }
      writeChunk(null, null, assistantResponse.finishReason);
      controller.enqueue(encoder.encode("data: [DONE]\n\n"));
      tracePane2(
        services,
        trace,
        buildChatCompletion(
          model,
          assistantResponse,
          conversationId,
          includeConversationId,
        ),
        200,
      );
      traceComplete(services, trace, 200);
      controller.close();
    },
  });

  headers.set("content-type", "text/event-stream");
  headers.set("cache-control", "no-cache");
  headers.set("connection", "keep-alive");
  headers.set("x-m365-conversation-id", conversationId);
  await services.debugLogger.logOutgoingResponse(200, headers.entries(), null);
  return finalizeOutgoingStreamResponse(services, stream, headers);
}

async function transformGraphStreamToOpenAi(
  services: Services,
  graphResponse: Response,
  model: string,
  initialConversationId: string,
  promptText: string,
  includeConversationId: boolean,
  headers: Headers,
  trace: TraceContext | null,
): Promise<Response> {
  const completionId = `chatcmpl-${randomUUID().replaceAll("-", "")}`;
  const created = nowUnix();
  let conversationId = initialConversationId;
  let emittedContent = "";
  const upstreamItems: JsonValue[] = [];

  const stream = new ReadableStream<Uint8Array>({
    start: async (controller) => {
      const encoder = new TextEncoder();
      const writeChunk = (
        role: string | null,
        content: string | null,
        finishReason: string | null,
      ) => {
        const chunk = buildChatCompletionChunk(
          completionId,
          created,
          model,
          role,
          content,
          finishReason,
          includeConversationId ? conversationId : null,
        );
        controller.enqueue(
          encoder.encode(`data: ${JSON.stringify(chunk)}\n\n`),
        );
      };

      writeChunk("assistant", null, null);
      if (graphResponse.body) {
        for await (const event of readSseEvents(graphResponse.body)) {
          const data = event.data.trim();
          if (!data) {
            continue;
          }
          if (data.toLowerCase() === "[done]") {
            break;
          }
          upstreamItems.push(tryParseJsonObject(data) ?? { rawText: data });
          tracePane4(
            services,
            trace,
            buildUpstreamStreamCapture("sse", upstreamItems),
            graphResponse.status,
          );

          const streamConversationId =
            extractCopilotConversationIdFromStream(data);
          if (streamConversationId) {
            conversationId = streamConversationId;
          }

          const latestAssistantText = extractCopilotAssistantTextFromStreamData(
            data,
            promptText,
          );
          if (!latestAssistantText) {
            continue;
          }

          const delta = computeTrailingDelta(
            emittedContent,
            latestAssistantText,
          );
          if (!delta) {
            continue;
          }
          emittedContent += delta;
          writeChunk(null, delta, null);
        }
      }
      writeChunk(null, null, "stop");
      controller.enqueue(encoder.encode("data: [DONE]\n\n"));
      tracePane4(
        services,
        trace,
        buildUpstreamStreamCapture("sse", upstreamItems),
        graphResponse.status,
      );
      tracePane2(
        services,
        trace,
        buildChatCompletion(
          model,
          {
            content: emittedContent,
            toolCalls: [],
            finishReason: "stop",
          },
          conversationId,
          includeConversationId,
        ),
        200,
      );
      traceComplete(services, trace, 200);
      controller.close();
    },
  });

  headers.set("content-type", "text/event-stream");
  headers.set("cache-control", "no-cache");
  headers.set("connection", "keep-alive");
  await services.debugLogger.logOutgoingResponse(200, headers.entries(), null);
  return finalizeOutgoingStreamResponse(services, stream, headers);
}

async function streamSubstrateAsSimulatedOpenAi(
  services: Services,
  authorizationHeader: string,
  initialConversationId: string,
  parsedRequest: ParsedOpenAiRequest,
  createdConversation: boolean,
  scopedConversationKey: string | null,
  headers: Headers,
  trace: TraceContext | null,
): Promise<Response> {
  const completionId = `chatcmpl-${randomUUID().replaceAll("-", "")}`;
  const created = nowUnix();
  let conversationId = initialConversationId;
  let isStartOfSession = createdConversation;
  let streamInitialized = false;
  let emittedContent = "";
  let emittedToolCalls = false;
  let accumulatedAssistantText = "";
  const shouldHoldForToolValidation =
    parsedRequest.tooling.tools.length > 0 &&
    (parsedRequest.tooling.toolChoiceMode === ToolChoiceModes.Required ||
      parsedRequest.tooling.toolChoiceMode === ToolChoiceModes.Function);
  const incrementalContentStreamingEnabled =
    services.options.substrate.incrementalSimulatedContentStreaming === true &&
    !shouldHoldForToolValidation &&
    parsedRequest.responseFormat === null;
  let incrementalContentStreamingSuppressed = false;
  let incrementalSuppressionReason: string | null = null;
  let incrementalExtractionAttemptCount = 0;
  let incrementalEmissionCount = 0;
  const streamStartedAtMs = Date.now();
  let firstDeltaAtMs: number | null = null;
  let firstParseableAtMs: number | null = null;
  let firstOpenAiChunkAtMs: number | null = null;
  let firstOpenAiPayloadChunkAtMs: number | null = null;
  let deltaChunkCount = 0;
  let deltaCharCount = 0;
  let parseAttemptCount = 0;
  let parseSuccessCount = 0;
  let openAiChunkCount = 0;
  let openAiInitChunkCount = 0;
  let openAiPayloadChunkCount = 0;
  let openAiContentChunkCount = 0;
  let openAiToolChunkCount = 0;
  let openAiFinishChunkCount = 0;
  let openAiErrorChunkCount = 0;
  let bufferedRetryCount = 0;
  const retryReasons: string[] = [];
  let usedBufferedFallbackAfterStreamFailure = false;

  const stream = new ReadableStream<Uint8Array>({
    start: async (controller) => {
      const encoder = new TextEncoder();
      const toOffsetMs = (value: number | null): number | null =>
        value === null ? null : value - streamStartedAtMs;
      const logStreamingDiagnostics = async (
        outcome: string,
        details: JsonObject = {},
      ): Promise<void> => {
        const diagnosticsLogger = services.debugLogger as DebugMarkdownLogger & {
          logSimulatedStreamingDiagnostics?: (payload: JsonObject) => Promise<void>;
        };
        if (
          typeof diagnosticsLogger.logSimulatedStreamingDiagnostics !==
          "function"
        ) {
          return;
        }
        await diagnosticsLogger.logSimulatedStreamingDiagnostics({
          completionId,
          conversationIdInitial: initialConversationId,
          conversationIdFinal: conversationId,
          requestModel: parsedRequest.model,
          toolChoiceMode: parsedRequest.tooling.toolChoiceMode,
          holdForToolValidation: shouldHoldForToolValidation,
          incrementalContentStreamingEnabled,
          incrementalContentStreamingSuppressed,
          incrementalSuppressionReason,
          incrementalExtractionAttemptCount,
          incrementalEmissionCount,
          outcome,
          streamInitialized,
          usedBufferedFallbackAfterStreamFailure,
          bufferedRetryCount,
          retryReasons,
          accumulatedAssistantTextLength: accumulatedAssistantText.length,
          emittedContentLength: emittedContent.length,
          emittedToolCalls,
          deltaChunkCount,
          deltaCharCount,
          parseAttemptCount,
          parseSuccessCount,
          openAiChunkCount,
          openAiInitChunkCount,
          openAiPayloadChunkCount,
          openAiContentChunkCount,
          openAiToolChunkCount,
          openAiFinishChunkCount,
          openAiErrorChunkCount,
          streamDurationMs: Date.now() - streamStartedAtMs,
          firstDeltaOffsetMs: toOffsetMs(firstDeltaAtMs),
          firstParseableOffsetMs: toOffsetMs(firstParseableAtMs),
          firstOpenAiChunkOffsetMs: toOffsetMs(firstOpenAiChunkAtMs),
          firstOpenAiPayloadChunkOffsetMs: toOffsetMs(
            firstOpenAiPayloadChunkAtMs,
          ),
          parseableBeforeFirstPayloadChunk:
            firstParseableAtMs !== null &&
            (firstOpenAiPayloadChunkAtMs === null ||
              firstParseableAtMs <= firstOpenAiPayloadChunkAtMs),
          ...details,
        });
      };
      const writeChunk = (
        role: string | null,
        content: string | null,
        finishReason: string | null,
        toolCalls?: unknown,
      ) => {
        const nowMs = Date.now();
        if (firstOpenAiChunkAtMs === null) {
          firstOpenAiChunkAtMs = nowMs;
        }
        openAiChunkCount += 1;
        if (role === "assistant" && content === null && finishReason === null && !toolCalls) {
          openAiInitChunkCount += 1;
        } else {
          openAiPayloadChunkCount += 1;
          if (firstOpenAiPayloadChunkAtMs === null) {
            firstOpenAiPayloadChunkAtMs = nowMs;
          }
          if (toolCalls) {
            openAiToolChunkCount += 1;
          } else if (finishReason === "error") {
            openAiErrorChunkCount += 1;
          } else if (finishReason !== null) {
            openAiFinishChunkCount += 1;
          } else if (content !== null) {
            openAiContentChunkCount += 1;
          }
        }
        const chunk = buildChatCompletionChunk(
          completionId,
          created,
          parsedRequest.model,
          role,
          content,
          finishReason,
          services.options.includeConversationIdInResponseBody
            ? conversationId
            : null,
          toolCalls as never,
        );
        controller.enqueue(
          encoder.encode(`data: ${JSON.stringify(chunk)}\n\n`),
        );
      };
      const ensureInitialized = () => {
        if (streamInitialized) {
          return;
        }
        streamInitialized = true;
        writeChunk("assistant", null, null);
      };
      const emitAssistantResponse = (assistantResponse: OpenAiAssistantResponse) => {
        ensureInitialized();
        if (assistantResponse.toolCalls.length > 0) {
          if (!emittedToolCalls) {
            emittedToolCalls = true;
            writeChunk(
              null,
              null,
              null,
              buildToolCallsDelta(assistantResponse.toolCalls),
            );
          }
          return;
        }

        const content = assistantResponse.content ?? "";
        const delta = computeTrailingDelta(emittedContent, content);
        if (!delta) {
          return;
        }
        emittedContent += delta;
        writeChunk(null, delta, null);
      };
      const tryParseFromAccumulatedText = (): {
        payload: JsonObject;
        assistantResponse: OpenAiAssistantResponse;
      } | null => {
        parseAttemptCount += 1;
        const simulatedPayload = tryExtractSimulatedResponsePayload(
          accumulatedAssistantText,
          "chat.completions",
        );
        if (!simulatedPayload) {
          return null;
        }
        const payload = normalizeSimulatedChatCompletionPayload(
          simulatedPayload,
          parsedRequest.model,
          conversationId,
          services.options.includeConversationIdInResponseBody,
        );
        const assistantResponse =
          tryBuildAssistantResponseFromChatCompletionPayload(payload);
        if (!assistantResponse) {
          return null;
        }
        parseSuccessCount += 1;
        if (firstParseableAtMs === null) {
          firstParseableAtMs = Date.now();
        }
        return { payload, assistantResponse };
      };
      const tryEmitFromAccumulatedText = (): {
        payload: JsonObject;
        assistantResponse: OpenAiAssistantResponse;
      } | null => {
        const parsed = tryParseFromAccumulatedText();
        if (!parsed) {
          return null;
        }
        if (shouldHoldForToolValidation) {
          return parsed;
        }
        const { assistantResponse } = parsed;
        emitAssistantResponse(assistantResponse);
        return parsed;
      };
      const tryEmitIncrementalContentFromAccumulatedText = (): void => {
        if (
          !incrementalContentStreamingEnabled ||
          incrementalContentStreamingSuppressed ||
          emittedToolCalls
        ) {
          return;
        }
        incrementalExtractionAttemptCount += 1;
        const incrementalExtraction =
          tryExtractIncrementalSimulatedChatContent(accumulatedAssistantText);
        if (incrementalExtraction.hasToolCalls) {
          incrementalContentStreamingSuppressed = true;
          incrementalSuppressionReason = "tool_calls_detected";
          return;
        }
        if (incrementalExtraction.content === null) {
          return;
        }
        ensureInitialized();
        const delta = computeTrailingDelta(
          emittedContent,
          incrementalExtraction.content,
        );
        if (!delta) {
          return;
        }
        emittedContent += delta;
        incrementalEmissionCount += 1;
        writeChunk(null, delta, null);
      };
      const executeBufferedTurnWithRecovery = async (): Promise<ChatResult> => {
        let result = await services.substrateClient.chat(
          authorizationHeader,
          conversationId,
          parsedRequest,
          isStartOfSession,
          async (update) => {
            traceSubstrateStreamUpdate(services, trace, update);
          },
        );
        if (
          shouldRetrySubstrateNoAssistantContent(
            TransportNames.Substrate,
            isStartOfSession,
            result,
          )
        ) {
          retryReasons.push("substrate_no_assistant_content");
          const createRetryConversation = services.substrateClient.createConversation();
          if (createRetryConversation.isSuccess && createRetryConversation.conversationId) {
            conversationId = createRetryConversation.conversationId;
            isStartOfSession = true;
            bufferedRetryCount += 1;
            if (scopedConversationKey) {
              services.conversationStore.set(scopedConversationKey, conversationId);
            }
            result = await services.substrateClient.chat(
              authorizationHeader,
              conversationId,
              parsedRequest,
              isStartOfSession,
              async (update) => {
                traceSubstrateStreamUpdate(services, trace, update);
              },
            );
          }
        }
        if (result.conversationId) {
          conversationId = result.conversationId;
        }
        return result;
      };
      const appendAssistantText = (
        chatResult: ChatResult,
        replace: boolean = false,
      ): void => {
        const latestAssistantText =
          chatResult.assistantText ??
          extractCopilotAssistantText(
            chatResult.responseJson,
            parsedRequest.promptText,
          ) ??
          "";
        if (replace) {
          accumulatedAssistantText = latestAssistantText;
          return;
        }
        const trailingAssistantTextDelta = computeTrailingDelta(
          accumulatedAssistantText,
          latestAssistantText,
        );
        if (trailingAssistantTextDelta) {
          accumulatedAssistantText += trailingAssistantTextDelta;
        }
      };

      let substrateResponse = await services.substrateClient.chatStream(
        authorizationHeader,
        conversationId,
        parsedRequest,
        isStartOfSession,
        async (update) => {
          traceSubstrateStreamUpdate(services, trace, update);
          if (update.conversationId) {
            conversationId = update.conversationId;
            if (scopedConversationKey) {
              services.conversationStore.set(
                scopedConversationKey,
                conversationId,
              );
            }
          }
          if (!update.deltaText) {
            return;
          }
          deltaChunkCount += 1;
          deltaCharCount += update.deltaText.length;
          if (firstDeltaAtMs === null) {
            firstDeltaAtMs = Date.now();
          }
          accumulatedAssistantText += update.deltaText;
          const streamedParsed = tryEmitFromAccumulatedText();
          if (!streamedParsed) {
            tryEmitIncrementalContentFromAccumulatedText();
          }
        },
      );
      tracePane3(services, trace, substrateResponse.upstreamRequestPayload ?? null);
      tracePane4(
        services,
        trace,
        substrateResponse.upstreamResponsePayload ?? null,
        substrateResponse.statusCode,
      );

      if (!substrateResponse.isSuccess) {
        if (!streamInitialized) {
          usedBufferedFallbackAfterStreamFailure = true;
          retryReasons.push("chat_stream_failed_then_buffered_retry");
          const bufferedRetry = await executeBufferedTurnWithRecovery();
          if (bufferedRetry.isSuccess) {
            substrateResponse = bufferedRetry;
          }
        }
      }

      if (!substrateResponse.isSuccess) {
        if (streamInitialized) {
          const details = extractGraphErrorMessage(substrateResponse.rawBody);
          const message = details
            ? `Substrate chat request failed. ${details}`
            : "Substrate chat request failed.";
          writeChunk(null, null, "error");
          controller.enqueue(
            encoder.encode(
              `event: error\ndata: ${JSON.stringify({
                error: {
                  message,
                  type: "api_error",
                  param: null,
                  code: "substrate_error",
                },
              })}\n\n`,
            ),
          );
          controller.enqueue(encoder.encode("data: [DONE]\n\n"));
          await logStreamingDiagnostics("upstream_failure_after_partial_emit", {
            statusCode: substrateResponse.statusCode,
          });
          traceError(
            services,
            trace,
            {
              message,
              type: "api_error",
              param: null,
              code: "substrate_error",
            },
            substrateResponse.statusCode,
          );
          controller.close();
          return;
        }
        const failure = await writeFromUpstreamFailure(
          services,
          substrateResponse.statusCode,
          substrateResponse.rawBody,
          "Substrate chat request failed.",
          "substrate_error",
        );
        controller.enqueue(
          encoder.encode(`event: error\ndata: ${await failure.text()}\n\n`),
        );
        controller.enqueue(encoder.encode("data: [DONE]\n\n"));
        await logStreamingDiagnostics("upstream_failure_before_emit", {
          statusCode: substrateResponse.statusCode,
        });
        traceError(
          services,
          trace,
          {
            message: "Substrate chat request failed.",
            type: "api_error",
            param: null,
            code: "substrate_error",
          },
          substrateResponse.statusCode,
        );
        controller.close();
        return;
      }

      if (substrateResponse.conversationId) {
        conversationId = substrateResponse.conversationId;
      }
      appendAssistantText(substrateResponse);

      let finalParsed = tryEmitFromAccumulatedText();
      const shouldRetryInvalidToolPayload =
        finalParsed &&
        shouldRetrySimulatedInvalidChatToolPayload(finalParsed.payload);
      const shouldRetryToollessPayload =
        finalParsed &&
        shouldRetrySimulatedToollessChatPayload(
          services.options,
          parsedRequest,
          finalParsed.payload,
        );
      const shouldRetrySimulatedPayload =
        Boolean(shouldRetryInvalidToolPayload || shouldRetryToollessPayload);
      if (shouldRetrySimulatedPayload && !streamInitialized) {
        if (shouldRetryInvalidToolPayload) {
          retryReasons.push("invalid_chat_tool_payload");
        }
        if (shouldRetryToollessPayload) {
          retryReasons.push("toolless_chat_payload");
        }
        bufferedRetryCount += 1;
        const bufferedRetry = await executeBufferedTurnWithRecovery();
        if (!bufferedRetry.isSuccess) {
          const failure = await writeFromUpstreamFailure(
            services,
            bufferedRetry.statusCode,
            bufferedRetry.rawBody,
            "Substrate chat request failed.",
            "substrate_error",
          );
          controller.enqueue(
            encoder.encode(`event: error\ndata: ${await failure.text()}\n\n`),
          );
          controller.enqueue(encoder.encode("data: [DONE]\n\n"));
          await logStreamingDiagnostics("buffered_retry_failed", {
            statusCode: bufferedRetry.statusCode,
          });
          controller.close();
          return;
        }
        appendAssistantText(bufferedRetry, true);
        finalParsed = tryParseFromAccumulatedText();
      }

      if (!finalParsed) {
        const message =
          "Simulated mode response did not include a usable assistant message or tool call payload.";
        if (streamInitialized) {
          writeChunk(null, null, "error");
        }
        controller.enqueue(
          encoder.encode(
            `event: error\ndata: ${JSON.stringify({
              error: {
                message,
                type: "api_error",
                param: null,
                code: "invalid_simulated_payload",
              },
            })}\n\n`,
          ),
        );
        controller.enqueue(encoder.encode("data: [DONE]\n\n"));
        await logStreamingDiagnostics("invalid_simulated_payload");
        traceError(
          services,
          trace,
          {
            message,
            type: "api_error",
            param: null,
            code: "invalid_simulated_payload",
          },
          502,
        );
        controller.close();
        return;
      }

      emitAssistantResponse(finalParsed.assistantResponse);
      ensureInitialized();
      writeChunk(null, null, finalParsed.assistantResponse.finishReason);
      controller.enqueue(encoder.encode("data: [DONE]\n\n"));
      tracePane2(services, trace, finalParsed.payload, 200);
      traceComplete(services, trace, 200);
      await logStreamingDiagnostics("completed", {
        finalFinishReason: finalParsed.assistantResponse.finishReason,
        finalHasToolCalls: finalParsed.assistantResponse.toolCalls.length > 0,
      });
      controller.close();
    },
  });

  headers.set("content-type", "text/event-stream");
  headers.set("cache-control", "no-cache");
  headers.set("connection", "keep-alive");
  headers.set("x-m365-conversation-id", initialConversationId);
  await services.debugLogger.logOutgoingResponse(200, headers.entries(), null);
  return finalizeOutgoingStreamResponse(services, stream, headers);
}

async function streamSubstrateAsOpenAi(
  services: Services,
  authorizationHeader: string,
  initialConversationId: string,
  parsedRequest: ParsedOpenAiRequest,
  createdConversation: boolean,
  scopedConversationKey: string | null,
  headers: Headers,
  trace: TraceContext | null,
): Promise<Response> {
  const completionId = `chatcmpl-${randomUUID().replaceAll("-", "")}`;
  const created = nowUnix();
  let conversationId = initialConversationId;
  let streamInitialized = false;
  let emitted = "";

  const stream = new ReadableStream<Uint8Array>({
    start: async (controller) => {
      const encoder = new TextEncoder();
      const writeChunk = (
        role: string | null,
        content: string | null,
        finishReason: string | null,
      ) => {
        const chunk = buildChatCompletionChunk(
          completionId,
          created,
          parsedRequest.model,
          role,
          content,
          finishReason,
          services.options.includeConversationIdInResponseBody
            ? conversationId
            : null,
        );
        controller.enqueue(
          encoder.encode(`data: ${JSON.stringify(chunk)}\n\n`),
        );
      };
      const ensureInitialized = () => {
        if (streamInitialized) {
          return;
        }
        streamInitialized = true;
        writeChunk("assistant", null, null);
      };

      const substrateResponse = await services.substrateClient.chatStream(
        authorizationHeader,
        conversationId,
        parsedRequest,
        createdConversation,
        async (update) => {
          traceSubstrateStreamUpdate(services, trace, update);
          if (update.conversationId) {
            conversationId = update.conversationId;
            if (scopedConversationKey) {
              services.conversationStore.set(
                scopedConversationKey,
                conversationId,
              );
            }
          }
          if (!update.deltaText) {
            return;
          }
          ensureInitialized();
          emitted += update.deltaText;
          writeChunk(null, update.deltaText, null);
        },
      );
      tracePane3(services, trace, substrateResponse.upstreamRequestPayload ?? null);
      tracePane4(
        services,
        trace,
        substrateResponse.upstreamResponsePayload ?? null,
        substrateResponse.statusCode,
      );

      if (!substrateResponse.isSuccess) {
        if (streamInitialized) {
          const details = extractGraphErrorMessage(substrateResponse.rawBody);
          const message = details
            ? `Substrate chat request failed. ${details}`
            : "Substrate chat request failed.";
          writeChunk(null, null, "error");
          controller.enqueue(
            encoder.encode(
              `event: error\ndata: ${JSON.stringify({
                error: {
                  message,
                  type: "api_error",
                  param: null,
                  code: "substrate_error",
                },
              })}\n\n`,
            ),
          );
          controller.enqueue(encoder.encode("data: [DONE]\n\n"));
          traceError(
            services,
            trace,
            {
              message,
              type: "api_error",
              param: null,
              code: "substrate_error",
            },
            substrateResponse.statusCode,
          );
          controller.close();
          return;
        }
        const failure = await writeFromUpstreamFailure(
          services,
          substrateResponse.statusCode,
          substrateResponse.rawBody,
          "Substrate chat request failed.",
          "substrate_error",
        );
        controller.enqueue(
          encoder.encode(`event: error\ndata: ${await failure.text()}\n\n`),
        );
        controller.enqueue(encoder.encode("data: [DONE]\n\n"));
        traceError(
          services,
          trace,
          {
            message: "Substrate chat request failed.",
            type: "api_error",
            param: null,
            code: "substrate_error",
          },
          substrateResponse.statusCode,
        );
        controller.close();
        return;
      }

      if (substrateResponse.conversationId) {
        conversationId = substrateResponse.conversationId;
      }
      const assistantText =
        substrateResponse.assistantText ??
        extractCopilotAssistantText(
          substrateResponse.responseJson,
          parsedRequest.promptText,
        ) ??
        "";

      ensureInitialized();
      const trailing = computeTrailingDelta(emitted, assistantText);
      if (trailing) {
        writeChunk(null, trailing, null);
      }
      writeChunk(null, null, "stop");
      controller.enqueue(encoder.encode("data: [DONE]\n\n"));
      tracePane2(
        services,
        trace,
        buildChatCompletion(
          parsedRequest.model,
          {
            content: `${emitted}${trailing ?? ""}`.trim() ? `${emitted}${trailing ?? ""}` : assistantText,
            toolCalls: [],
            finishReason: "stop",
          },
          conversationId,
          services.options.includeConversationIdInResponseBody,
        ),
        200,
      );
      traceComplete(services, trace, 200);
      controller.close();
    },
  });

  headers.set("content-type", "text/event-stream");
  headers.set("cache-control", "no-cache");
  headers.set("connection", "keep-alive");
  headers.set("x-m365-conversation-id", initialConversationId);
  await services.debugLogger.logOutgoingResponse(200, headers.entries(), null);
  return finalizeOutgoingStreamResponse(services, stream, headers);
}
