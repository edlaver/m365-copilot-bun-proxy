import { beforeAll, describe, expect, test } from "bun:test";
import { readSseEvents, tryParseJsonObject } from "../src/proxy/utils";

const proxyBaseUrl = Bun.env.PROXY_BASE_URL ?? "http://localhost:4000";
const testModel = Bun.env.PROXY_TEST_MODEL ?? "m365-copilot";
const testTransport = Bun.env.PROXY_TEST_TRANSPORT ?? "substrate";

const toolDefinition = [
  {
    type: "function",
    function: {
      name: "get_time",
      description: "Get the current time in a time zone.",
      parameters: {
        type: "object",
        properties: {
          zone: { type: "string" },
        },
        required: ["zone"],
        additionalProperties: false,
      },
    },
  },
];

beforeAll(async () => {
  const response = await fetch(new URL("/healthz", proxyBaseUrl));
  expect(response.ok).toBeTrue();
});

describe("proxy tool-calling integration", () => {
  test(
    "chat/completions non-stream returns tool_calls",
    async () => {
      let lastFailure = "unknown";
      for (let attempt = 1; attempt <= 3; attempt++) {
        const response = await postJson("/v1/chat/completions", {
          model: testModel,
          stream: false,
          temperature: 0,
          messages: [
            {
              role: "system",
              content:
                "You are a tool-calling engine. When possible, call the requested tool and do not output natural language.",
            },
            {
              role: "user",
              content:
                "Call the get_time tool now with zone UTC. Return only the tool call payload.",
            },
          ],
          tools: toolDefinition,
          tool_choice: {
            type: "function",
            function: { name: "get_time" },
          },
        });

        if (!response.ok) {
          lastFailure = `HTTP ${response.status}: ${await response.text()}`;
          continue;
        }

        const body = (await response.json()) as Record<string, unknown>;
        const toolCall = extractFirstChatToolCall(body);
        if (toolCall) {
          const functionNode = toolCall.function as Record<string, unknown>;
          expect(functionNode.name).toBe("get_time");
          expect(typeof functionNode.arguments).toBe("string");
          const choices = body.choices as Array<Record<string, unknown>>;
          expect(choices[0]?.finish_reason).toBe("tool_calls");
          return;
        }

        lastFailure = `attempt ${attempt} produced no tool call: ${JSON.stringify(body)}`;
      }

      throw new Error(
        `Expected a tool call in chat/completions after retries, but failed: ${lastFailure}`,
      );
    },
    180_000,
  );

  test(
    "chat/completions stream emits tool_calls delta and tool_calls finish reason",
    async () => {
      let lastFailure = "unknown";
      for (let attempt = 1; attempt <= 3; attempt++) {
        const response = await postJson("/v1/chat/completions", {
          model: testModel,
          stream: true,
          temperature: 0,
          messages: [
            {
              role: "user",
              content:
                "Call the get_time tool now with zone UTC. Return only the tool call payload.",
            },
          ],
          tools: toolDefinition,
          tool_choice: {
            type: "function",
            function: { name: "get_time" },
          },
        });

        if (!response.ok || !response.body) {
          lastFailure = `HTTP ${response.status}: ${await response.text()}`;
          continue;
        }

        let sawToolDelta = false;
        let finishReason: string | null = null;
        let sawDone = false;
        let sawError = false;

        for await (const event of readSseEvents(response.body)) {
          const data = event.data.trim();
          if (!data) {
            continue;
          }
          if (data.toLowerCase() === "[done]") {
            sawDone = true;
            break;
          }
          if (event.event.toLowerCase() === "error") {
            sawError = true;
            lastFailure = `stream error: ${data}`;
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

          const finish = (first as Record<string, unknown>).finish_reason;
          if (typeof finish === "string") {
            finishReason = finish;
          }

          const delta = (first as Record<string, unknown>).delta;
          if (!delta || typeof delta !== "object" || Array.isArray(delta)) {
            continue;
          }
          const toolCalls = (delta as Record<string, unknown>).tool_calls;
          if (Array.isArray(toolCalls) && toolCalls.length > 0) {
            sawToolDelta = true;
          }
        }

        if (!sawError && sawToolDelta && finishReason === "tool_calls" && sawDone) {
          return;
        }
        lastFailure = `attempt ${attempt} stream assertion failed (toolDelta=${sawToolDelta}, finish=${finishReason}, done=${sawDone}, error=${sawError})`;
      }

      throw new Error(
        `Expected streamed tool_calls delta + finish_reason=tool_calls after retries, but failed: ${lastFailure}`,
      );
    },
    180_000,
  );

  test(
    "responses endpoint returns function_call output items",
    async () => {
      let lastFailure = "unknown";
      for (let attempt = 1; attempt <= 3; attempt++) {
        const response = await postJson("/v1/responses", {
          model: testModel,
          stream: false,
          temperature: 0,
          input: "Call the get_time tool now with zone UTC and do not provide a natural-language answer.",
          tools: toolDefinition,
          tool_choice: {
            type: "function",
            function: { name: "get_time" },
          },
        });

        if (!response.ok) {
          lastFailure = `HTTP ${response.status}: ${await response.text()}`;
          continue;
        }

        const body = (await response.json()) as Record<string, unknown>;
        const output = body.output;
        if (!Array.isArray(output)) {
          lastFailure = `attempt ${attempt} missing output array`;
          continue;
        }
        const functionCall = output.find((item) => {
          if (!item || typeof item !== "object" || Array.isArray(item)) {
            return false;
          }
          const typed = item as Record<string, unknown>;
          return typed.type === "function_call" && typed.name === "get_time";
        }) as Record<string, unknown> | undefined;

        if (functionCall) {
          expect(typeof functionCall.arguments).toBe("string");
          return;
        }

        lastFailure = `attempt ${attempt} did not include function_call output: ${JSON.stringify(body)}`;
      }

      throw new Error(
        `Expected function_call output item from /v1/responses after retries, but failed: ${lastFailure}`,
      );
    },
    180_000,
  );
});

function extractFirstChatToolCall(
  response: Record<string, unknown>,
): Record<string, unknown> | null {
  const choices = response.choices;
  if (!Array.isArray(choices) || choices.length === 0) {
    return null;
  }
  const first = choices[0];
  if (!first || typeof first !== "object" || Array.isArray(first)) {
    return null;
  }
  const message = (first as Record<string, unknown>).message;
  if (!message || typeof message !== "object" || Array.isArray(message)) {
    return null;
  }
  const toolCalls = (message as Record<string, unknown>).tool_calls;
  if (!Array.isArray(toolCalls) || toolCalls.length === 0) {
    return null;
  }
  const toolCall = toolCalls[0];
  if (!toolCall || typeof toolCall !== "object" || Array.isArray(toolCall)) {
    return null;
  }
  return toolCall as Record<string, unknown>;
}

async function postJson(
  path: string,
  body: Record<string, unknown>,
): Promise<Response> {
  const url = new URL(path, proxyBaseUrl);
  return fetch(url, {
    method: "POST",
    headers: {
      "content-type": "application/json",
      "x-m365-transport": testTransport,
    },
    body: JSON.stringify(body),
  });
}
