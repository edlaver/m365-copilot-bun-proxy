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

const codingAgentTools = [
  {
    type: "function",
    function: {
      name: "read_file",
      description: "Read a UTF-8 text file by path.",
      parameters: {
        type: "object",
        properties: {
          path: { type: "string" },
        },
        required: ["path"],
        additionalProperties: false,
      },
    },
  },
  {
    type: "function",
    function: {
      name: "write_file",
      description: "Write UTF-8 file content to a path.",
      parameters: {
        type: "object",
        properties: {
          path: { type: "string" },
          content: { type: "string" },
        },
        required: ["path", "content"],
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

  test(
    "chat/completions multi-step coding-agent simulation edits a mocked file via tool calls",
    async () => {
      const mockFiles = new Map<string, string>([
        ["/workspace/app.ts", "export function sum(a: number, b: number) {\n  return a + b;\n}\n"],
      ]);
      const messages: Array<Record<string, unknown>> = [
        {
          role: "system",
          content:
            "You are a coding agent. Follow the requested tool sequence and use provided tools.",
        },
        {
          role: "user",
          content:
            "We are editing /workspace/app.ts. Use tools only for file operations.",
        },
      ];
      let conversationId: string | null = null;

      messages.push({
        role: "user",
        content:
          "Step 1: Call read_file for /workspace/app.ts.",
      });
      const readStep = await requestSpecificToolCall({
        messages,
        expectedToolName: "read_file",
        tools: codingAgentTools,
        conversationId,
      });
      conversationId = readStep.conversationId;
      const readOutput = runMockCodingTool(readStep.toolCall, mockFiles);
      messages.push(readStep.assistantMessage);
      messages.push(buildToolMessage(readStep.toolCall, readOutput));

      messages.push({
        role: "user",
        content:
          "Step 2: Call write_file to add this first line: // TODO: reviewed by coding agent",
      });
      const writeStep = await requestSpecificToolCall({
        messages,
        expectedToolName: "write_file",
        tools: codingAgentTools,
        conversationId,
      });
      conversationId = writeStep.conversationId;
      const writeOutput = runMockCodingTool(writeStep.toolCall, mockFiles);
      messages.push(writeStep.assistantMessage);
      messages.push(buildToolMessage(writeStep.toolCall, writeOutput));

      messages.push({
        role: "user",
        content:
          "Step 3: Call read_file again for /workspace/app.ts to verify the edit.",
      });
      const verifyStep = await requestSpecificToolCall({
        messages,
        expectedToolName: "read_file",
        tools: codingAgentTools,
        conversationId,
      });
      const verifyOutput = runMockCodingTool(verifyStep.toolCall, mockFiles);
      messages.push(verifyStep.assistantMessage);
      messages.push(buildToolMessage(verifyStep.toolCall, verifyOutput));

      const primaryContent = mockFiles.get("/workspace/app.ts") ?? "";
      expect(primaryContent).toContain("// TODO: reviewed by coding agent");
      expect(primaryContent).toContain("export function sum");

      const verifyParsed = JSON.parse(verifyOutput) as Record<string, unknown>;
      expect(typeof verifyParsed.content).toBe("string");
      expect(String(verifyParsed.content)).toContain(
        "// TODO: reviewed by coding agent",
      );
    },
    300_000,
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
  options?: {
    conversationId?: string | null;
  },
): Promise<Response> {
  const url = new URL(path, proxyBaseUrl);
  const headers = new Headers({
    "content-type": "application/json",
    "x-m365-transport": testTransport,
  });
  if (options?.conversationId?.trim()) {
    headers.set("x-m365-conversation-id", options.conversationId.trim());
  }
  return fetch(url, {
    method: "POST",
    headers,
    body: JSON.stringify(body),
  });
}

async function requestSpecificToolCall(params: {
  messages: Array<Record<string, unknown>>;
  expectedToolName: string;
  tools: Array<Record<string, unknown>>;
  conversationId: string | null;
}): Promise<{
  conversationId: string | null;
  toolCall: Record<string, unknown>;
  assistantMessage: Record<string, unknown>;
}> {
  let lastFailure = "unknown";
  for (let attempt = 1; attempt <= 3; attempt++) {
    const response = await postJson(
      "/v1/chat/completions",
      {
        model: testModel,
        stream: false,
        temperature: 0,
        messages: params.messages,
        tools: params.tools,
        tool_choice: {
          type: "function",
          function: { name: params.expectedToolName },
        },
      },
      { conversationId: params.conversationId },
    );

    const nextConversationId =
      response.headers.get("x-m365-conversation-id") ?? params.conversationId;

    if (!response.ok) {
      lastFailure = `HTTP ${response.status}: ${await response.text()}`;
      continue;
    }

    const body = (await response.json()) as Record<string, unknown>;
    const toolCall = extractFirstChatToolCall(body);
    const assistantMessage = extractAssistantMessage(body);
    if (!toolCall || !assistantMessage) {
      lastFailure = `attempt ${attempt} missing tool call/message: ${JSON.stringify(body)}`;
      continue;
    }

    const toolName = readToolCallName(toolCall);
    if (toolName !== params.expectedToolName) {
      lastFailure = `attempt ${attempt} expected ${params.expectedToolName}, received ${toolName ?? "(none)"}`;
      continue;
    }

    return {
      conversationId: nextConversationId,
      toolCall,
      assistantMessage,
    };
  }

  throw new Error(
    `Expected tool '${params.expectedToolName}' after retries, but failed: ${lastFailure}`,
  );
}

function extractAssistantMessage(
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
  return message as Record<string, unknown>;
}

function buildToolMessage(
  toolCall: Record<string, unknown>,
  content: string,
): Record<string, unknown> {
  const toolCallId = readToolCallId(toolCall);
  const toolName = readToolCallName(toolCall);
  return {
    role: "tool",
    tool_call_id: toolCallId ?? "mock_call",
    name: toolName ?? "unknown_tool",
    content,
  };
}

function readToolCallId(toolCall: Record<string, unknown>): string | null {
  const id = toolCall.id;
  return typeof id === "string" && id.trim() ? id.trim() : null;
}

function readToolCallName(toolCall: Record<string, unknown>): string | null {
  const functionNode = toolCall.function;
  if (!functionNode || typeof functionNode !== "object" || Array.isArray(functionNode)) {
    return null;
  }
  const name = (functionNode as Record<string, unknown>).name;
  return typeof name === "string" && name.trim() ? name.trim() : null;
}

function runMockCodingTool(
  toolCall: Record<string, unknown>,
  files: Map<string, string>,
): string {
  const functionNode = toolCall.function;
  if (!functionNode || typeof functionNode !== "object" || Array.isArray(functionNode)) {
    return JSON.stringify({ ok: false, error: "Invalid tool call format." });
  }
  const toolName = readToolCallName(toolCall) ?? "";
  const argumentsText = (functionNode as Record<string, unknown>).arguments;
  const parsedArgs = tryParseArguments(argumentsText);

  if (toolName === "read_file") {
    const path =
      pickString(parsedArgs.path, parsedArgs.file_path, parsedArgs.filename) ??
      "/workspace/app.ts";
    const content = files.get(path) ?? "";
    return JSON.stringify({ ok: true, path, content });
  }

  if (toolName === "write_file") {
    const path =
      pickString(parsedArgs.path, parsedArgs.file_path, parsedArgs.filename) ??
      "/workspace/app.ts";
    const previous = files.get(path) ?? "";
    const requestedContent = pickString(
      parsedArgs.content,
      parsedArgs.text,
      parsedArgs.new_content,
    );
    const nextContent =
      requestedContent ?? ensureTodoCommentLine(previous);
    files.set(path, nextContent);
    return JSON.stringify({
      ok: true,
      path,
      bytes_written: nextContent.length,
    });
  }

  return JSON.stringify({ ok: false, error: `Unsupported tool '${toolName}'.` });
}

function tryParseArguments(
  rawArguments: unknown,
): Record<string, unknown> {
  if (typeof rawArguments !== "string" || !rawArguments.trim()) {
    return {};
  }
  try {
    const parsed = JSON.parse(rawArguments) as unknown;
    if (parsed && typeof parsed === "object" && !Array.isArray(parsed)) {
      return parsed as Record<string, unknown>;
    }
    return {};
  } catch {
    return {};
  }
}

function ensureTodoCommentLine(content: string): string {
  const todoLine = "// TODO: reviewed by coding agent";
  if (content.includes(todoLine)) {
    return content;
  }
  return `${todoLine}\n${content}`;
}

function pickString(...values: unknown[]): string | null {
  for (const value of values) {
    if (typeof value === "string" && value.trim()) {
      return value.trim();
    }
  }
  return null;
}
