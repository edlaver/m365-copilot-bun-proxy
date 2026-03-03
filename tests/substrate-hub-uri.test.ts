import { describe, expect, test } from "bun:test";
import { buildSubstrateHubUri } from "../src/proxy/clients";
import {
  LogLevels,
  OpenAiTransformModes,
  TransportNames,
  type WrapperOptions,
} from "../src/proxy/types";

describe("buildSubstrateHubUri", () => {
  test("adds disableMemory=1 when temporaryChat is enabled", () => {
    const uri = buildSubstrateHubUri(
      createOptions(true),
      "oid-1",
      "tid-1",
      "token-1",
      "request-1",
      "session-1",
      "conversation-1",
    );

    expect(uri.searchParams.get("disableMemory")).toBe("1");
  });

  test("does not include disableMemory when temporaryChat is disabled", () => {
    const uri = buildSubstrateHubUri(
      createOptions(false),
      "oid-1",
      "tid-1",
      "token-1",
      "request-1",
      "session-1",
      "conversation-1",
    );

    expect(uri.searchParams.has("disableMemory")).toBeFalse();
  });
});

function createOptions(temporaryChat: boolean): WrapperOptions {
  return {
    listenUrl: "http://localhost:4000",
    debugPath: null,
    logLevel: LogLevels.Info,
    openAiTransformMode: OpenAiTransformModes.Simulated,
    temporaryChat,
    ignoreIncomingAuthorizationHeader: true,
    playwrightBrowser: "edge",
    transport: TransportNames.Substrate,
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
      earlyCompleteOnSimulatedPayload: false,
    },
    defaultModel: "m365-copilot",
    defaultTimeZone: "America/New_York",
    conversationTtlMinutes: 180,
    maxAdditionalContextMessages: 16,
    includeConversationIdInResponseBody: true,
    retrySimulatedToollessResponses: true,
  };
}
