import { randomUUID } from "node:crypto";
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
  tryExtractSimulatedResponsePayload,
} from "./openai";
import { ResponseStore } from "./response-store";
import {
  buildFunctionCallOutputItems,
  buildMessageOutputItem,
  buildOpenAiResponseFromAssistant,
  buildOpenAiResponseObject,
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
  isSupportedTransport,
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
  type JsonObject,
  type ChatResult,
  type OpenAiAssistantResponse,
  type ParsedOpenAiRequest,
  type ParsedResponsesRequest,
  type WrapperOptions,
} from "./types";
import { ProxyTokenProvider } from "./token-provider";
import {
  extractGraphErrorMessage,
  isJsonObject,
  nowUnix,
  readSseEvents,
  tryGetString,
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
};

export function createProxyApp(services: Services): Hono {
  const app = new Hono();
  const { options } = services;

  app.get("/healthz", (c) => c.json({ status: "ok" }));
  app.get("/v1/models", (c) =>
    c.json({
      object: "list",
      data: [
        {
          id: options.defaultModel?.trim()
            ? options.defaultModel
            : "m365-copilot",
          object: "model",
          created: 0,
          owned_by: "microsoft-365-copilot",
        },
      ],
    }),
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

async function handleChat(
  request: Request,
  services: Services,
): Promise<Response> {
  const {
    options,
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

  const parsed = tryParseOpenAiRequest(payload.json, options);
  if (!parsed.ok) {
    return writeOpenAiError(
      services,
      400,
      parsed.error,
      "invalid_request_error",
      "invalid_request",
    );
  }
  const parsedRequest = parsed.request;

  const selectedTransport = resolveTransport(request, payload.json, options);
  if (!isSupportedTransport(selectedTransport)) {
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
  const shouldBufferAssistant =
    requiresBufferedAssistantResponse(parsedRequest);

  const executeChatTurn = async (): Promise<ChatResult> => {
    if (selectedTransport === TransportNames.Substrate) {
      return substrateClient.chat(
        authorizationHeader,
        conversationId!,
        parsedRequest,
        createdConversation,
      );
    }
    return graphClient.chat(authorizationHeader, conversationId!, graphPayload);
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
          !hasUsableSimulatedChatCompletionPayload(normalizedSimulatedPayload)
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

    if (!normalized || !hasUsableSimulatedChatCompletionPayload(normalized)) {
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

    if (!normalized || !hasUsableSimulatedChatCompletionPayload(normalized)) {
      return writeOpenAiError(
        services,
        502,
        "Simulated mode response did not include a usable assistant message or tool call payload.",
        "api_error",
        "invalid_simulated_payload",
      );
    }
    const body = JSON.stringify(normalized);
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

  responseHeaders.set("content-type", "application/json");
  await debugLogger.logOutgoingResponse(200, responseHeaders.entries(), body);
  return new Response(body, { status: 200, headers: responseHeaders });
}

async function handleResponsesCreate(
  request: Request,
  services: Services,
): Promise<Response> {
  const {
    options,
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

  const parsed = tryParseResponsesRequest(payload.json, options);
  if (!parsed.ok) {
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

  const selectedTransport = resolveTransport(request, payload.json, options);
  if (!isSupportedTransport(selectedTransport)) {
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
  const shouldBufferAssistant = requiresBufferedAssistantResponse(baseRequest);

  const executeChatTurn = async (): Promise<ChatResult> => {
    if (selectedTransport === TransportNames.Substrate) {
      return substrateClient.chat(
        authorizationHeader,
        conversationId!,
        baseRequest,
        createdConversation,
      );
    }
    return graphClient.chat(authorizationHeader, conversationId!, graphPayload);
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

        if (!normalized || !hasUsableSimulatedResponsesPayload(normalized.responseBody)) {
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

        if (!normalized || !hasUsableSimulatedResponsesPayload(normalized.responseBody)) {
          return writeOpenAiError(
            services,
            502,
            "Simulated mode response did not include a usable response output payload.",
            "api_error",
            "invalid_simulated_payload",
          );
        }
        return buildSimulatedResponsesStreamResponse(
          services,
          parsedRequest,
          conversationId,
          normalized.responseBody,
          responseHeaders,
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
      return buildBufferedResponsesStreamResponse(
        services,
        parsedRequest,
        conversationId,
        assistantResponse,
        responseHeaders,
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
      return transformGraphStreamToResponses(
        services,
        graphResponse,
        parsedRequest,
        conversationId,
        scopedConversationKey,
        responseHeaders,
      );
    }

    return streamSubstrateAsResponses(
      services,
      authorizationHeader,
      conversationId,
      parsedRequest,
      createdConversation,
      scopedConversationKey,
      responseHeaders,
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

    if (!normalized || !hasUsableSimulatedResponsesPayload(normalized.responseBody)) {
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

    if (!normalized || !hasUsableSimulatedResponsesPayload(normalized.responseBody)) {
      return writeOpenAiError(
        services,
        502,
        "Simulated mode response did not include a usable response output payload.",
        "api_error",
        "invalid_simulated_payload",
      );
    }

    responseStore.set(normalized.responseId, normalized.responseBody, conversationId);
    responseHeaders.set("content-type", "application/json");
    const body = JSON.stringify(normalized.responseBody);
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

  responseHeaders.set("content-type", "application/json");
  const body = JSON.stringify(responseBody);
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

async function buildBufferedResponsesStreamResponse(
  services: Services,
  parsedRequest: ParsedResponsesRequest,
  conversationId: string,
  assistantResponse: ReturnType<typeof buildAssistantResponse>,
  headers: Headers,
): Promise<Response> {
  const responseId = createOpenAiResponseId();
  const createdAt = nowUnix();
  const includeConversationId = services.options.includeConversationIdInResponseBody;
  const stream = new ReadableStream<Uint8Array>({
    start(controller) {
      const encoder = new TextEncoder();
      const writeDataEvent = (event: JsonObject) => {
        controller.enqueue(encoder.encode(`data: ${JSON.stringify(event)}\n\n`));
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
      } catch (error) {
        writeError(
          `Failed to build streaming response. ${String(error)}`,
          "response_stream_error",
        );
      } finally {
        controller.close();
      }
    },
  });

  headers.set("content-type", "text/event-stream");
  headers.set("cache-control", "no-cache");
  headers.set("connection", "keep-alive");
  headers.set("x-m365-conversation-id", conversationId);
  await services.debugLogger.logOutgoingResponse(200, headers.entries(), null);
  return new Response(stream, { status: 200, headers });
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
  const output = payload.output;
  if (!Array.isArray(output) || output.length === 0) {
    return null;
  }

  const toolCalls: JsonObject[] = [];
  let messageText: string | null = null;

  for (const item of output) {
    if (!isJsonObject(item)) {
      continue;
    }
    const type = (tryGetString(item, "type") ?? "").toLowerCase();
    if (type === "function_call") {
      const toolCall = normalizeSimulatedToolCall({
        id:
          tryGetString(item, "call_id") ??
          tryGetString(item, "id") ??
          `call_${randomUUID().replaceAll("-", "")}`,
        type: "function",
        function: {
          name: tryGetString(item, "name") ?? "unknown_tool",
          arguments: item.arguments,
        },
      });
      if (toolCall) {
        toolCalls.push(toolCall);
      }
      continue;
    }
    if (type !== "message") {
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
  if (!Array.isArray(responseBody.output)) {
    const outputText =
      tryGetString(responseBody, "output_text") ??
      tryGetString(responseBody, "text") ??
      "";
    responseBody.output = [
      buildMessageOutputItem(createOpenAiOutputItemId("msg"), outputText, "completed"),
    ];
  }
  if (responseBody.output_text === undefined) {
    responseBody.output_text = extractResponseOutputText(
      Array.isArray(responseBody.output) ? responseBody.output : [],
    );
  }
  if (includeConversationId) {
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
    if (type === "message" && extractMessageOutputText(item).trim()) {
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
  );
}

async function buildSimulatedResponsesStreamResponse(
  services: Services,
  parsedRequest: ParsedResponsesRequest,
  conversationId: string,
  payload: JsonObject,
  headers: Headers,
): Promise<Response> {
  const includeConversationId = services.options.includeConversationIdInResponseBody;
  const normalized = normalizeSimulatedResponsesPayload(
    payload,
    parsedRequest,
    conversationId,
    includeConversationId,
  );
  const responseBody = normalized.responseBody;
  const responseId = normalized.responseId;
  const outputItems = Array.isArray(responseBody.output)
    ? responseBody.output.filter(isJsonObject)
    : [];

  const stream = new ReadableStream<Uint8Array>({
    start(controller) {
      const encoder = new TextEncoder();
      const writeDataEvent = (event: JsonObject) => {
        controller.enqueue(encoder.encode(`data: ${JSON.stringify(event)}\n\n`));
      };

      const inProgress: JsonObject = {
        ...responseBody,
        status: "in_progress",
      };
      writeDataEvent(buildResponseCreatedEvent(inProgress));
      writeDataEvent(buildResponseInProgressEvent(inProgress));

      for (let index = 0; index < outputItems.length; index++) {
        const item = outputItems[index];
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
          if (text) {
            writeDataEvent(
              buildResponseOutputTextDeltaEvent(responseId, index, itemId, text),
            );
          }
          writeDataEvent(
            buildResponseOutputTextDoneEvent(responseId, index, itemId, text),
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

      const completed: JsonObject = {
        ...responseBody,
        status: "completed",
      };
      writeDataEvent(buildResponseCompletedEvent(completed));
      services.responseStore.set(responseId, completed, conversationId);
      controller.close();
    },
  });

  headers.set("content-type", "text/event-stream");
  headers.set("cache-control", "no-cache");
  headers.set("connection", "keep-alive");
  headers.set("x-m365-conversation-id", conversationId);
  await services.debugLogger.logOutgoingResponse(200, headers.entries(), null);
  return new Response(stream, { status: 200, headers });
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

function extractMessageOutputText(outputItem: JsonObject): string {
  if ((tryGetString(outputItem, "type") ?? "").toLowerCase() !== "message") {
    return "";
  }
  const content = outputItem.content;
  if (!Array.isArray(content)) {
    return "";
  }

  const textParts: string[] = [];
  for (const part of content) {
    if (!isJsonObject(part)) {
      continue;
    }
    const type = (tryGetString(part, "type") ?? "").toLowerCase();
    if (type !== "output_text" && type !== "text") {
      continue;
    }
    const text = tryGetString(part, "text");
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
): Promise<Response> {
  const includeConversationId = services.options.includeConversationIdInResponseBody;
  const responseId = createOpenAiResponseId();
  const createdAt = nowUnix();
  const messageItemId = createOpenAiOutputItemId("msg");
  let conversationId = initialConversationId;
  let emittedContent = "";

  const stream = new ReadableStream<Uint8Array>({
    start: async (controller) => {
      const encoder = new TextEncoder();
      const writeDataEvent = (event: JsonObject) => {
        controller.enqueue(encoder.encode(`data: ${JSON.stringify(event)}\n\n`));
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
      } catch (error) {
        writeError(
          `Microsoft Graph chatOverStream request failed. ${String(error)}`,
          "graph_error",
        );
      } finally {
        controller.close();
      }
    },
  });

  headers.set("content-type", "text/event-stream");
  headers.set("cache-control", "no-cache");
  headers.set("connection", "keep-alive");
  headers.set("x-m365-conversation-id", initialConversationId);
  await services.debugLogger.logOutgoingResponse(200, headers.entries(), null);
  return new Response(stream, { status: 200, headers });
}

async function streamSubstrateAsResponses(
  services: Services,
  authorizationHeader: string,
  initialConversationId: string,
  parsedRequest: ParsedResponsesRequest,
  createdConversation: boolean,
  scopedConversationKey: string | null,
  headers: Headers,
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
        controller.enqueue(encoder.encode(`data: ${JSON.stringify(event)}\n\n`));
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

      const substrateResponse = await services.substrateClient.chatStream(
        authorizationHeader,
        conversationId,
        parsedRequest.base,
        createdConversation,
        async (update) => {
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

      if (!substrateResponse.isSuccess) {
        const details = extractGraphErrorMessage(substrateResponse.rawBody);
        writeError(
          details
            ? `Substrate chat request failed. ${details}`
            : "Substrate chat request failed.",
          "substrate_error",
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
      controller.close();
    },
  });

  headers.set("content-type", "text/event-stream");
  headers.set("cache-control", "no-cache");
  headers.set("connection", "keep-alive");
  headers.set("x-m365-conversation-id", initialConversationId);
  await services.debugLogger.logOutgoingResponse(200, headers.entries(), null);
  return new Response(stream, { status: 200, headers });
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

async function buildAssistantStreamResponse(
  services: Services,
  model: string,
  conversationId: string,
  assistantResponse: ReturnType<typeof buildAssistantResponse>,
  includeConversationId: boolean,
  headers: Headers,
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
      controller.close();
    },
  });

  headers.set("content-type", "text/event-stream");
  headers.set("cache-control", "no-cache");
  headers.set("connection", "keep-alive");
  headers.set("x-m365-conversation-id", conversationId);
  await services.debugLogger.logOutgoingResponse(200, headers.entries(), null);
  return new Response(stream, { status: 200, headers });
}

async function transformGraphStreamToOpenAi(
  services: Services,
  graphResponse: Response,
  model: string,
  initialConversationId: string,
  promptText: string,
  includeConversationId: boolean,
  headers: Headers,
): Promise<Response> {
  const completionId = `chatcmpl-${randomUUID().replaceAll("-", "")}`;
  const created = nowUnix();
  let conversationId = initialConversationId;
  let emittedContent = "";

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
      controller.close();
    },
  });

  headers.set("content-type", "text/event-stream");
  headers.set("cache-control", "no-cache");
  headers.set("connection", "keep-alive");
  await services.debugLogger.logOutgoingResponse(200, headers.entries(), null);
  return new Response(stream, { status: 200, headers });
}

async function streamSubstrateAsOpenAi(
  services: Services,
  authorizationHeader: string,
  initialConversationId: string,
  parsedRequest: ParsedOpenAiRequest,
  createdConversation: boolean,
  scopedConversationKey: string | null,
  headers: Headers,
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
      controller.close();
    },
  });

  headers.set("content-type", "text/event-stream");
  headers.set("cache-control", "no-cache");
  headers.set("connection", "keep-alive");
  headers.set("x-m365-conversation-id", initialConversationId);
  await services.debugLogger.logOutgoingResponse(200, headers.entries(), null);
  return new Response(stream, { status: 200, headers });
}
