import { describe, expect, test } from "bun:test";
import { createProxyApp } from "../src/proxy/server";
import { ConversationStore } from "../src/proxy/conversation-store";
import { ResponseStore } from "../src/proxy/response-store";
import { ProxyTokenProvider } from "../src/proxy/token-provider";
import { ProxyVizTraceStore } from "../src/proxy/viz-trace-store";
import { DebugMarkdownLogger } from "../src/proxy/logger";
import type {
  ChatResult,
  CreateConversationResult,
  JsonObject,
  ParsedOpenAiRequest,
  WrapperOptions,
} from "../src/proxy/types";
import {
  LogLevels,
  OpenAiTransformModes,
  TransportNames,
} from "../src/proxy/types";
import { readSseEvents, tryParseJsonObject } from "../src/proxy/utils";
import type { CopilotGraphClient, CopilotSubstrateClient } from "../src/proxy/clients";

describe("proxy viz trace capture", () => {
  test("rejects invalid transform override header", async () => {
    const app = createTestApp();

    const response = await app.fetch(
      new Request("http://localhost/v1/chat/completions", {
        method: "POST",
        headers: {
          authorization: "Bearer test",
          "content-type": "application/json",
          "x-m365-openai-transform-mode": "broken",
        },
        body: JSON.stringify({
          model: "m365-copilot",
          messages: [{ role: "user", content: "hello" }],
        }),
      }),
    );

    expect(response.status).toBe(400);
    const body = (await response.json()) as JsonObject;
    expect(body.error).toBeDefined();
  });

  test("captures chat/completions non-stream trace with simulated override", async () => {
    const simulatedPayload: JsonObject = {
      id: "chatcmpl_trace_simulated",
      object: "chat.completion",
      created: 1700000000,
      model: "simulated-model",
      choices: [
        {
          index: 0,
          message: {
            role: "assistant",
            content: "hello from trace",
          },
          finish_reason: "stop",
        },
      ],
    };
    const app = createTestApp({
      graphClient: {
        chat: async (
          _authorizationHeader: string,
          conversationId: string,
          payload: JsonObject,
        ): Promise<ChatResult> => ({
          isSuccess: true,
          statusCode: 200,
          responseJson: {
            id: conversationId,
            messages: [
              { text: "Say hello." },
              { text: "```json\n" + JSON.stringify(simulatedPayload, null, 2) + "\n```" },
            ],
          },
          rawBody: JSON.stringify({
            id: conversationId,
            messages: [{ text: "Say hello." }, { text: "assistant" }],
          }),
          assistantText: "```json\n" + JSON.stringify(simulatedPayload, null, 2) + "\n```",
          conversationId,
          upstreamRequestPayload: payload,
          upstreamResponsePayload: {
            id: conversationId,
            messages: [{ text: "Say hello." }, { text: "assistant" }],
          },
        }),
      },
    });

    const traceId = "trace-chat-non-stream";
    const response = await app.fetch(
      new Request("http://localhost/v1/chat/completions", {
        method: "POST",
        headers: {
          authorization: "Bearer test",
          "content-type": "application/json",
          "x-m365-viz-trace-id": traceId,
          "x-m365-openai-transform-mode": OpenAiTransformModes.Simulated,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          messages: [{ role: "user", content: "Say hello." }],
        }),
      }),
    );

    expect(response.status).toBe(200);
    const body = (await response.json()) as JsonObject;
    expect(body.id).toBe("chatcmpl_trace_simulated");

    const trace = await getTrace(app, traceId);
    expect(trace.status).toBe("completed");
    expect(trace.transformMode).toBe("simulated");
    expect(
      String(((trace.pane3 as JsonObject).message as JsonObject).text ?? ""),
    ).toContain('"messages"');
    expect((trace.pane4 as JsonObject).id).toBe("conv-1");
    expect((trace.pane2 as JsonObject).id).toBe("chatcmpl_trace_simulated");
  });

  test("captures graph chat/completions stream trace as buffered json", async () => {
    const app = createTestApp({
      graphClient: {
        chatOverStream: async (): Promise<Response> =>
          buildGraphStreamResponse("conv-1", ["Hello", "Hello world"]),
      },
    });

    const traceId = "trace-chat-stream";
    const response = await app.fetch(
      new Request("http://localhost/v1/chat/completions", {
        method: "POST",
        headers: {
          authorization: "Bearer test",
          "content-type": "application/json",
          "x-m365-viz-trace-id": traceId,
          "x-m365-openai-transform-mode": OpenAiTransformModes.Mapped,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: true,
          messages: [{ role: "user", content: "Say hello." }],
        }),
      }),
    );

    expect(response.status).toBe(200);
    await consumeResponse(response);

    const trace = await getTrace(app, traceId);
    expect(trace.status).toBe("completed");
    expect((trace.pane4 as JsonObject).streamType).toBe("sse");
    expect(Array.isArray((trace.pane4 as JsonObject).items)).toBeTrue();
    const pane2 = trace.pane2 as JsonObject;
    const choices = pane2.choices as JsonObject[];
    const message = choices[0]?.message as JsonObject;
    expect(message.content).toBe("Hello world");
  });

  test("captures responses non-stream trace", async () => {
    const app = createTestApp({
      graphClient: {
        chat: async (
          _authorizationHeader: string,
          conversationId: string,
          payload: JsonObject,
        ): Promise<ChatResult> => ({
          isSuccess: true,
          statusCode: 200,
          responseJson: {
            id: conversationId,
            messages: [{ text: "Say hello." }, { text: "assistant hello" }],
          },
          rawBody: JSON.stringify({
            id: conversationId,
            messages: [{ text: "Say hello." }, { text: "assistant hello" }],
          }),
          assistantText: "assistant hello",
          conversationId,
          upstreamRequestPayload: payload,
          upstreamResponsePayload: {
            id: conversationId,
            messages: [{ text: "Say hello." }, { text: "assistant hello" }],
          },
        }),
      },
    });

    const traceId = "trace-responses-non-stream";
    const response = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          authorization: "Bearer test",
          "content-type": "application/json",
          "x-m365-viz-trace-id": traceId,
          "x-m365-openai-transform-mode": OpenAiTransformModes.Mapped,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          input: "Say hello.",
        }),
      }),
    );

    expect(response.status).toBe(200);
    const body = (await response.json()) as JsonObject;
    expect(body.object).toBe("response");

    const trace = await getTrace(app, traceId);
    expect(trace.status).toBe("completed");
    expect((trace.pane2 as JsonObject).object).toBe("response");
    expect((trace.pane3 as JsonObject).message).toEqual({ text: "Say hello." });
    expect((trace.pane4 as JsonObject).id).toBe("conv-1");
  });

  test("captures graph responses stream trace", async () => {
    const app = createTestApp({
      graphClient: {
        chatOverStream: async (): Promise<Response> =>
          buildGraphStreamResponse("conv-1", ["A", "AB"]),
      },
    });

    const traceId = "trace-responses-stream";
    const response = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          authorization: "Bearer test",
          "content-type": "application/json",
          "x-m365-viz-trace-id": traceId,
          "x-m365-openai-transform-mode": OpenAiTransformModes.Mapped,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: true,
          input: [{ role: "user", content: [{ type: "input_text", text: "Hi" }] }],
        }),
      }),
    );

    expect(response.status).toBe(200);
    await consumeResponse(response);

    const trace = await getTrace(app, traceId);
    expect(trace.status).toBe("completed");
    expect((trace.pane4 as JsonObject).streamType).toBe("sse");
    expect((trace.pane2 as JsonObject).object).toBe("response");
    expect((trace.pane2 as JsonObject).output_text).toBe("AB");
  });

  test("captures substrate chat stream trace", async () => {
    const app = createTestApp({
      options: createOptions({ transport: TransportNames.Substrate }),
      substrateClient: {
        chatStream: async (
          _authorizationHeader: string,
          conversationId: string,
          _request: ParsedOpenAiRequest,
          _isStartOfSession: boolean,
          onStreamUpdate: (update: {
            deltaText: string | null;
            conversationId: string | null;
          }) => Promise<void>,
        ): Promise<ChatResult> => {
          await onStreamUpdate({
            deltaText: "hello ",
            conversationId,
          });
          await onStreamUpdate({
            deltaText: "world",
            conversationId,
          });
          return {
            isSuccess: true,
            statusCode: 200,
            responseJson: {
              id: conversationId,
              messages: [{ text: "Say hello." }, { text: "hello world" }],
            },
            rawBody: "{}",
            assistantText: "hello world",
            conversationId,
            upstreamRequestPayload: {
              type: "substrate-request",
              conversationId,
            },
            upstreamResponsePayload: {
              streamType: "signalr",
              frames: [{ text: "hello world" }],
            },
          };
        },
      },
    });

    const traceId = "trace-substrate-chat-stream";
    const response = await app.fetch(
      new Request("http://localhost/v1/chat/completions", {
        method: "POST",
        headers: {
          authorization: "Bearer test",
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Substrate,
          "x-m365-viz-trace-id": traceId,
          "x-m365-openai-transform-mode": OpenAiTransformModes.Mapped,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: true,
          messages: [{ role: "user", content: "Say hello." }],
        }),
      }),
    );

    expect(response.status).toBe(200);
    await consumeResponse(response);

    const trace = await getTrace(app, traceId);
    expect(trace.status).toBe("completed");
    expect(trace.transport).toBe(TransportNames.Substrate);
    expect((trace.pane3 as JsonObject).type).toBe("substrate-request");
    expect((trace.pane4 as JsonObject).streamType).toBe("signalr");
  });

  test("publishes live substrate trace updates before the request completes", async () => {
    const releaseResponse = createDeferred<void>();
    const liveUpdateReady = createDeferred<void>();
    const conversationId = "substrate-live-1";
    const app = createTestApp({
      options: createOptions({ transport: TransportNames.Substrate }),
      substrateClient: {
        chat: async (
          _authorizationHeader: string,
          currentConversationId: string,
          _request: ParsedOpenAiRequest,
          _isStartOfSession: boolean,
          onStreamUpdate?: (update: {
            deltaText: string | null;
            conversationId: string | null;
            upstreamRequestPayload?: JsonObject | null;
            upstreamResponsePayload?: JsonObject | null;
          }) => Promise<void>,
        ): Promise<ChatResult> => {
          await onStreamUpdate?.({
            deltaText: null,
            conversationId: currentConversationId,
            upstreamRequestPayload: {
              type: "substrate-request",
              conversationId: currentConversationId,
            },
            upstreamResponsePayload: {
              streamType: "signalr",
              frameCount: 1,
              frames: [{ type: 1, payload: "handshake" }],
            },
          });
          liveUpdateReady.resolve();
          await releaseResponse.promise;
          return {
            isSuccess: true,
            statusCode: 200,
            responseJson: {
              id: currentConversationId,
              messages: [{ text: "Say hello." }, { text: "hello world" }],
            },
            rawBody: "{}",
            assistantText: "hello world",
            conversationId: currentConversationId,
            upstreamRequestPayload: {
              type: "substrate-request",
              conversationId: currentConversationId,
            },
            upstreamResponsePayload: {
              streamType: "signalr",
              frameCount: 2,
              frames: [
                { type: 1, payload: "handshake" },
                { type: 2, payload: "hello world" },
              ],
            },
          };
        },
      },
    });

    const traceId = "trace-substrate-live";
    const responsePromise = app.fetch(
      new Request("http://localhost/v1/chat/completions", {
        method: "POST",
        headers: {
          authorization: "Bearer test",
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Substrate,
          "x-m365-viz-trace-id": traceId,
          "x-m365-openai-transform-mode": OpenAiTransformModes.Mapped,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          messages: [{ role: "user", content: "Say hello." }],
        }),
      }),
    );

    await liveUpdateReady.promise;

    const pendingTrace = await getTrace(app, traceId);
    expect(pendingTrace.status).toBe("pending");
    expect((pendingTrace.pane3 as JsonObject).type).toBe("substrate-request");
    expect((pendingTrace.pane4 as JsonObject).frameCount).toBe(1);

    releaseResponse.resolve();

    const response = await responsePromise;
    expect(response.status).toBe(200);
    await response.json();

    const completedTrace = await getTrace(app, traceId);
    expect(completedTrace.status).toBe("completed");
    expect((completedTrace.pane4 as JsonObject).frameCount).toBe(2);
    expect((completedTrace.pane2 as JsonObject).object).toBe("chat.completion");
    expect(((completedTrace.pane2 as JsonObject).choices as JsonObject[])[0]).toBeDefined();
    expect(conversationId).toBe("substrate-live-1");
  });
});

function createTestApp(overrides: {
  options?: WrapperOptions;
  graphClient?: Partial<CopilotGraphClient>;
  substrateClient?: Partial<CopilotSubstrateClient>;
} = {}) {
  const options = overrides.options ?? createOptions();
  const conversationStore = new ConversationStore(options);
  const responseStore = new ResponseStore(options);
  const vizTraceStore = new ProxyVizTraceStore(3600);

  let createConversationCount = 0;
  const graphClient = {
    createConversation: async (): Promise<CreateConversationResult> => ({
      isSuccess: true,
      statusCode: 200,
      conversationId: `conv-${++createConversationCount}`,
      rawBody: "{}",
    }),
    chat: async (
      _authorizationHeader: string,
      conversationId: string,
      payload: JsonObject,
    ): Promise<ChatResult> => ({
      isSuccess: true,
      statusCode: 200,
      responseJson: {
        id: conversationId,
        messages: [{ text: "user" }, { text: "assistant" }],
      },
      rawBody: JSON.stringify({
        id: conversationId,
        messages: [{ text: "user" }, { text: "assistant" }],
      }),
      assistantText: "assistant",
      conversationId,
      upstreamRequestPayload: payload,
      upstreamResponsePayload: {
        id: conversationId,
        messages: [{ text: "user" }, { text: "assistant" }],
      },
    }),
    chatOverStream: async (): Promise<Response> =>
      buildGraphStreamResponse("conv-1", ["assistant"]),
    ...overrides.graphClient,
  } as unknown as CopilotGraphClient;

  const substrateClient = {
    createConversation: (): CreateConversationResult => ({
      isSuccess: true,
      statusCode: 200,
      conversationId: `substrate-${++createConversationCount}`,
      rawBody: "{}",
    }),
    chat: async (): Promise<ChatResult> => ({
      isSuccess: true,
      statusCode: 200,
      responseJson: {
        id: "substrate-1",
        messages: [{ text: "user" }, { text: "assistant" }],
      },
      rawBody: "{}",
      assistantText: "assistant",
      conversationId: "substrate-1",
      upstreamRequestPayload: { type: "substrate-request" },
      upstreamResponsePayload: { streamType: "signalr", frames: [] },
    }),
    chatStream: async (): Promise<ChatResult> => ({
      isSuccess: true,
      statusCode: 200,
      responseJson: {
        id: "substrate-1",
        messages: [{ text: "user" }, { text: "assistant" }],
      },
      rawBody: "{}",
      assistantText: "assistant",
      conversationId: "substrate-1",
      upstreamRequestPayload: { type: "substrate-request" },
      upstreamResponsePayload: { streamType: "signalr", frames: [] },
    }),
    ...overrides.substrateClient,
  } as unknown as CopilotSubstrateClient;

  const debugLogger = {
    logIncomingRequest: async () => {},
    logOutgoingResponse: async () => {},
    logUpstreamRequest: async () => {},
    logUpstreamResponse: async () => {},
    logSubstrateFrame: async () => {},
  } as unknown as DebugMarkdownLogger;

  const tokenProvider = {
    resolveAuthorizationHeader: async () => "Bearer unit-test-token",
  } as unknown as ProxyTokenProvider;

  return createProxyApp({
    options,
    debugLogger,
    graphClient,
    substrateClient,
    conversationStore,
    responseStore,
    tokenProvider,
    vizTraceStore,
  });
}

function createOptions(overrides: Partial<WrapperOptions> = {}): WrapperOptions {
  return {
    listenUrl: "http://localhost:4000",
    debugPath: null,
    logLevel: LogLevels.Info,
    openAiTransformMode: OpenAiTransformModes.Mapped,
    temporaryChat: true,
    ignoreIncomingAuthorizationHeader: true,
    playwrightBrowser: "edge",
    transport: TransportNames.Graph,
    graphBaseUrl: "https://graph.microsoft.com",
    createConversationPath: "/beta/copilot/conversations",
    chatPathTemplate: "/beta/copilot/conversations/{conversationId}/chat",
    chatOverStreamPathTemplate:
      "/beta/copilot/conversations/{conversationId}/chatOverStream",
    substrate: {
      hubPath: "wss://substrate.office.com/m365Copilot/Chathub",
      source: "officeweb",
      quoteSourceInQuery: true,
      scenario: "OfficeWebIncludedCopilot",
      origin: "https://m365.cloud.microsoft",
      product: "Office",
      agentHost: "Bizchat.FullScreen",
      licenseType: "Starter",
      agent: "web",
      variants: null,
      clientPlatform: "web",
      productThreadType: "Office",
      invocationTimeoutSeconds: 120,
      keepAliveSeconds: 15,
      optionsSets: [],
      allowedMessageTypes: [],
      invocationTarget: "chat",
      invocationType: 4,
      locale: "en-US",
      experienceType: "Default",
      earlyCompleteOnSimulatedPayload: false,
      entityAnnotationTypes: [],
      incrementalSimulatedContentStreaming: false,
    },
    defaultModel: "m365-copilot",
    defaultTimeZone: "America/New_York",
    conversationTtlMinutes: 180,
    maxAdditionalContextMessages: 16,
    includeConversationIdInResponseBody: true,
    retrySimulatedToollessResponses: true,
    ...overrides,
  };
}

async function getTrace(app: ReturnType<typeof createTestApp>, traceId: string) {
  const response = await app.fetch(
    new Request(`http://localhost/__viz/traces/${traceId}`),
  );
  expect(response.status).toBe(200);
  return (await response.json()) as JsonObject;
}

async function consumeResponse(response: Response): Promise<void> {
  expect(response.body).not.toBeNull();
  for await (const event of readSseEvents(response.body!)) {
    const data = event.data.trim();
    if (!data || data.toLowerCase() === "[done]") {
      continue;
    }
    tryParseJsonObject(data);
  }
}

function buildGraphStreamResponse(
  conversationId: string,
  snapshots: string[],
): Response {
  const encoder = new TextEncoder();
  const stream = new ReadableStream<Uint8Array>({
    start(controller) {
      for (const snapshot of snapshots) {
        controller.enqueue(
          encoder.encode(
            `data: ${JSON.stringify({
              id: conversationId,
              messages: [{ text: "prompt" }, { text: snapshot }],
            })}\n\n`,
          ),
        );
      }
      controller.enqueue(encoder.encode("data: [DONE]\n\n"));
      controller.close();
    },
  });

  return new Response(stream, {
    status: 200,
    headers: {
      "content-type": "text/event-stream",
    },
  });
}

function createDeferred<T>() {
  let resolve!: (value: T | PromiseLike<T>) => void;
  let reject!: (reason?: unknown) => void;
  const promise = new Promise<T>((innerResolve, innerReject) => {
    resolve = innerResolve;
    reject = innerReject;
  });
  return {
    promise,
    resolve,
    reject,
  };
}
