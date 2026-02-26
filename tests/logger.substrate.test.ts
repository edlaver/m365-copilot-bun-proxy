import { afterEach, describe, expect, test } from "bun:test";
import { mkdtempSync, readFileSync, readdirSync, rmSync } from "node:fs";
import path from "node:path";
import { tmpdir } from "node:os";
import { DebugMarkdownLogger } from "../src/proxy/logger";
import {
  LogLevels,
  OpenAiTransformModes,
  TransportNames,
  type JsonObject,
  type WrapperOptions,
} from "../src/proxy/types";

const tempDirs: string[] = [];

afterEach(() => {
  while (tempDirs.length > 0) {
    const dir = tempDirs.pop();
    if (!dir) {
      continue;
    }
    rmSync(dir, { recursive: true, force: true });
  }
});

describe("DebugMarkdownLogger substrate response logging", () => {
  test("debug mode emits structured frame JSON and omits metadata-only frames", async () => {
    const debugPath = mkdtempSync(path.join(tmpdir(), "proxy-logger-"));
    tempDirs.push(debugPath);
    const logger = new DebugMarkdownLogger(createOptions(debugPath), true);

    const metadataOnlyFrame = {
      type: 1,
      target: "update",
      arguments: [{ requestId: "req-1", nonce: "n1" }],
    };
    const firstTextFrame = {
      type: 1,
      target: "update",
      arguments: [
        {
          requestId: "req-1",
          messages: [{ author: "bot", messageId: "msg-1", text: "```" }],
        },
      ],
    };
    const completeJsonFrame = {
      type: 1,
      target: "update",
      arguments: [
        {
          requestId: "req-1",
          messages: [
            {
              author: "bot",
              messageId: "msg-1",
              text:
                "```json\n" +
                JSON.stringify(
                  {
                    id: "chatcmpl-sim-1",
                    object: "chat.completion",
                    choices: [
                      {
                        index: 0,
                        finish_reason: "stop",
                        message: { role: "assistant", content: "ok" },
                      },
                    ],
                  },
                  null,
                  2,
                ) +
                "\n```",
            },
          ],
        },
      ],
    };
    const terminalFrame = { type: 3, invocationId: "0" };

    const payload = [
      JSON.stringify(metadataOnlyFrame),
      JSON.stringify(firstTextFrame),
      JSON.stringify(completeJsonFrame),
      JSON.stringify(terminalFrame),
    ].join("\u001e");

    await logger.logSubstrateFrame(
      "wss://substrate.office.com/m365Copilot/Chathub",
      "response",
      `${payload}\u001e`,
    );

    const files = readdirSync(debugPath).filter((name) =>
      name.endsWith("-substrate-response.md"),
    );
    expect(files.length).toBe(1);

    const content = readFileSync(path.join(debugPath, files[0]), "utf8");
    const startMarker = "```json\n";
    const startIndex = content.indexOf(startMarker);
    const endIndex = content.lastIndexOf("\n```");
    expect(startIndex).toBeGreaterThanOrEqual(0);
    expect(endIndex).toBeGreaterThan(startIndex);
    const jsonText = content
      .slice(startIndex + startMarker.length, endIndex)
      .trim();
    const parsed = JSON.parse(jsonText) as JsonObject;

    expect(parsed.format).toBe("signalr-json-v1");
    expect(parsed.frameCount).toBe(4);
    expect(parsed.includedFrameCount).toBe(3);
    expect(parsed.omittedFrameCount).toBe(1);
    expect(Array.isArray(parsed.frames)).toBeTrue();

    const frames = parsed.frames as JsonObject[];
    const reasons = frames
      .map((frame) => String(frame.reason ?? ""))
      .filter((value) => value.length > 0);
    expect(reasons.includes("first_text")).toBeTrue();
    expect(reasons.includes("complete_markdown_json")).toBeTrue();
    expect(reasons.includes("terminal")).toBeTrue();
  });

  test("trace mode writes simulated streaming diagnostics", async () => {
    const debugPath = mkdtempSync(path.join(tmpdir(), "proxy-logger-"));
    tempDirs.push(debugPath);
    const logger = new DebugMarkdownLogger(
      createOptions(debugPath, LogLevels.Trace),
      true,
    );

    await logger.logSimulatedStreamingDiagnostics({
      completionId: "chatcmpl-test",
      outcome: "completed",
      parseAttemptCount: 3,
      parseSuccessCount: 1,
    });

    const files = readdirSync(debugPath).filter((name) =>
      name.endsWith("-simulated-streaming.md"),
    );
    expect(files.length).toBe(1);
    const content = readFileSync(path.join(debugPath, files[0]), "utf8");
    expect(content.includes("Simulated Streaming Diagnostics")).toBeTrue();
    expect(content.includes("\"parseAttemptCount\": 3")).toBeTrue();
  });

  test("debug mode does not write simulated streaming diagnostics", async () => {
    const debugPath = mkdtempSync(path.join(tmpdir(), "proxy-logger-"));
    tempDirs.push(debugPath);
    const logger = new DebugMarkdownLogger(
      createOptions(debugPath, LogLevels.Debug),
      true,
    );

    await logger.logSimulatedStreamingDiagnostics({
      completionId: "chatcmpl-test",
      outcome: "completed",
    });

    const files = readdirSync(debugPath).filter((name) =>
      name.endsWith("-simulated-streaming.md"),
    );
    expect(files.length).toBe(0);
  });
});

function createOptions(
  debugPath: string,
  logLevel: (typeof LogLevels)[keyof typeof LogLevels] = LogLevels.Debug,
): WrapperOptions {
  return {
    listenUrl: "http://localhost:4000",
    debugPath,
    logLevel,
    openAiTransformMode: OpenAiTransformModes.Simulated,
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
