import { describe, expect, test } from "bun:test";
import { CopilotGraphClient, CopilotSubstrateClient } from "../src/proxy/clients";
import { ConversationStore } from "../src/proxy/conversation-store";
import { DebugMarkdownLogger } from "../src/proxy/logger";
import { createProxyApp } from "../src/proxy/server";
import { ResponseStore } from "../src/proxy/response-store";
import { ProxyTokenProvider } from "../src/proxy/token-provider";
import {
  LogLevels,
  OpenAiTransformModes,
  TransportNames,
  type ChatResult,
  type CreateConversationResult,
  type JsonObject,
  type WrapperOptions,
} from "../src/proxy/types";
import { readSseEvents, tryParseJsonObject } from "../src/proxy/utils";

describe("simulated transform mode proxy flow", () => {
  test("chat/completions non-stream wraps incoming JSON and returns parsed JSON block", async () => {
    const simulatedCompletion: JsonObject = {
      id: "chatcmpl_simulated_1",
      object: "chat.completion",
      created: 1700000000,
      model: "simulated-model",
      choices: [
        {
          index: 0,
          message: {
            role: "assistant",
            content: "hello from simulated mode",
          },
          finish_reason: "stop",
        },
      ],
    };

    let capturedPrompt = "";
    const app = createProxyApp(
      createServices((conversationId, payload) => {
        capturedPrompt = readPrompt(payload);
        return buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson(simulatedCompletion),
        );
      }),
    );

    const requestBody: JsonObject = {
      model: "m365-copilot",
      stream: false,
      messages: [{ role: "user", content: "Say hello." }],
    };
    const response = await app.fetch(
      new Request("http://localhost/v1/chat/completions", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify(requestBody),
      }),
    );

    expect(response.status).toBe(200);
    const body = (await response.json()) as JsonObject;
    expect(body.id).toBe("chatcmpl_simulated_1");
    expect((body.choices as JsonObject[])[0]?.message).toEqual(
      (simulatedCompletion.choices as JsonObject[])[0]?.message,
    );
    expect(typeof capturedPrompt).toBe("string");
    expect(capturedPrompt).toContain("simulating the OpenAI chat.completions");
    expect(capturedPrompt).toContain("```json");
    expect(capturedPrompt).toContain("\"messages\"");
  });

  test("chat/completions stream uses simulated JSON payload", async () => {
    const simulatedCompletion: JsonObject = {
      id: "chatcmpl_simulated_stream",
      object: "chat.completion",
      created: 1700000000,
      model: "simulated-model",
      choices: [
        {
          index: 0,
          message: {
            role: "assistant",
            content: null,
            tool_calls: [
              {
                id: "call_sim_1",
                type: "function",
                function: {
                  name: "get_time",
                  arguments: "{\"zone\":\"UTC\"}",
                },
              },
            ],
          },
          finish_reason: "tool_calls",
        },
      ],
    };

    const app = createProxyApp(
      createServices((conversationId, payload) =>
        buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson(simulatedCompletion),
        ),
      ),
    );

    const response = await app.fetch(
      new Request("http://localhost/v1/chat/completions", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: true,
          messages: [{ role: "user", content: "Call get_time for UTC." }],
        }),
      }),
    );

    expect(response.status).toBe(200);
    expect(response.body).not.toBeNull();

    let sawToolDelta = false;
    let sawDone = false;
    let finishReason: string | null = null;
    for await (const event of readSseEvents(response.body!)) {
      const data = event.data.trim();
      if (!data) {
        continue;
      }
      if (data.toLowerCase() === "[done]") {
        sawDone = true;
        break;
      }
      const chunk = tryParseJsonObject(data);
      const choices = chunk?.choices;
      if (!Array.isArray(choices) || choices.length === 0) {
        continue;
      }
      const first = choices[0];
      if (!first || typeof first !== "object" || Array.isArray(first)) {
        continue;
      }
      const typed = first as Record<string, unknown>;
      if (typeof typed.finish_reason === "string") {
        finishReason = typed.finish_reason;
      }
      const delta = typed.delta;
      if (!delta || typeof delta !== "object" || Array.isArray(delta)) {
        continue;
      }
      const toolCalls = (delta as Record<string, unknown>).tool_calls;
      if (Array.isArray(toolCalls) && toolCalls.length > 0) {
        sawToolDelta = true;
      }
    }

    expect(sawToolDelta).toBeTrue();
    expect(finishReason).toBe("tool_calls");
    expect(sawDone).toBeTrue();
  });

  test("responses non-stream returns simulated response payload object", async () => {
    const simulatedResponse: JsonObject = {
      id: "resp_simulated_1",
      object: "response",
      created_at: 1700000000,
      status: "completed",
      model: "simulated-model",
      output: [
        {
          id: "msg_sim_1",
          type: "message",
          status: "completed",
          role: "assistant",
          content: [{ type: "output_text", text: "hello from responses mode" }],
        },
      ],
      output_text: "hello from responses mode",
    };

    const app = createProxyApp(
      createServices((conversationId, payload) =>
        buildGraphChatResult(
          conversationId,
          payload,
          toMarkdownJson(simulatedResponse),
        ),
      ),
    );

    const createResponse = await app.fetch(
      new Request("http://localhost/v1/responses", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-transport": TransportNames.Graph,
        },
        body: JSON.stringify({
          model: "m365-copilot",
          stream: false,
          input: "Say hello.",
        }),
      }),
    );

    expect(createResponse.status).toBe(200);
    const body = (await createResponse.json()) as JsonObject;
    expect(body.id).toBe("resp_simulated_1");
    expect(body.output_text).toBe("hello from responses mode");
  });
});

function createServices(
  onChat: (conversationId: string, payload: JsonObject) => ChatResult,
): Parameters<typeof createProxyApp>[0] {
  const options = createOptions();
  const conversationStore = new ConversationStore(options);
  const responseStore = new ResponseStore(options);

  const graphClient = {
    createConversation: async (): Promise<CreateConversationResult> => ({
      isSuccess: true,
      statusCode: 200,
      conversationId: "conv_simulated_1",
      rawBody: "{}",
    }),
    chat: async (
      _authorizationHeader: string,
      conversationId: string,
      payload: JsonObject,
    ): Promise<ChatResult> => onChat(conversationId, payload),
    chatOverStream: async (): Promise<Response> => {
      throw new Error("chatOverStream is not used in simulated mode tests.");
    },
  } as unknown as CopilotGraphClient;

  const substrateClient = {
    createConversation: (): CreateConversationResult => ({
      isSuccess: true,
      statusCode: 200,
      conversationId: "conv_substrate_unused",
      rawBody: "{}",
    }),
    chat: async (): Promise<ChatResult> => {
      throw new Error("substrate chat is not used in this test.");
    },
    chatStream: async (): Promise<ChatResult> => {
      throw new Error("substrate stream is not used in this test.");
    },
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

  return {
    options,
    debugLogger,
    graphClient,
    substrateClient,
    conversationStore,
    responseStore,
    tokenProvider,
  };
}

function createOptions(): WrapperOptions {
  return {
    listenUrl: "http://localhost:4000",
    debugPath: null,
    logLevel: LogLevels.Info,
    openAiTransformMode: OpenAiTransformModes.Simulated,
    ignoreIncomingAuthorizationHeader: true,
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
      entityAnnotationTypes: [],
    },
    defaultModel: "m365-copilot",
    defaultTimeZone: "America/New_York",
    conversationTtlMinutes: 180,
    maxAdditionalContextMessages: 16,
    includeConversationIdInResponseBody: true,
  };
}

function buildGraphChatResult(
  conversationId: string,
  payload: JsonObject,
  assistantText: string,
): ChatResult {
  return {
    isSuccess: true,
    statusCode: 200,
    responseJson: {
      id: conversationId,
      messages: [{ text: readPrompt(payload) }, { text: assistantText }],
    },
    rawBody: "{}",
    assistantText: null,
    conversationId: conversationId,
  };
}

function readPrompt(payload: JsonObject): string {
  const message = payload.message;
  if (!message || typeof message !== "object" || Array.isArray(message)) {
    return "";
  }
  const text = (message as Record<string, unknown>).text;
  return typeof text === "string" ? text : "";
}

function toMarkdownJson(payload: JsonObject): string {
  return `\`\`\`json\n${JSON.stringify(payload, null, 2)}\n\`\`\``;
}
