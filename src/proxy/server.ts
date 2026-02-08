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
} from "./openai";
import {
  buildCopilotRequestPayload,
  isSupportedTransport,
  resolveTransport,
  scopeConversationKey,
  selectConversation,
  tryParseOpenAiRequest,
} from "./request-parser";
import {
  TransportNames,
  type ChatResult,
  type ParsedOpenAiRequest,
  type WrapperOptions,
} from "./types";
import {
  extractGraphErrorMessage,
  normalizeBearerToken,
  nowUnix,
  readSseEvents,
  tryReadJsonPayload,
} from "./utils";

type Services = {
  options: WrapperOptions;
  debugLogger: DebugMarkdownLogger;
  graphClient: CopilotGraphClient;
  substrateClient: CopilotSubstrateClient;
  conversationStore: ConversationStore;
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
  const authorizationHeader = normalizeBearerToken(
    request.headers.get("authorization"),
  );
  if (!authorizationHeader) {
    return writeOpenAiError(
      services,
      401,
      "Missing Authorization Bearer token.",
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
  } else if (scopedConversationKey) {
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

  if (parsedRequest.stream) {
    if (shouldBufferAssistant) {
      const buffered = await executeChatTurn();
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
      const assistantResponse = buildAssistantResponse(
        parsedRequest,
        assistantText,
      );
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

  const chatResponse = await executeChatTurn();
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
  const assistantResponse = buildAssistantResponse(
    parsedRequest,
    assistantText,
  );
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
