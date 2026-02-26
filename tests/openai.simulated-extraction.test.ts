import { describe, expect, test } from "bun:test";
import {
  tryExtractIncrementalSimulatedChatContent,
  tryExtractSimulatedResponsePayload,
} from "../src/proxy/openai";

describe("tryExtractSimulatedResponsePayload", () => {
  test("prefers chat completion response JSON over echoed request JSON", () => {
    const echoedRequest = {
      model: "gpt-5.1-2025-11-13",
      messages: [{ role: "user", content: "hello" }],
      tools: [
        {
          type: "function",
          function: { name: "write_to_file", parameters: { type: "object" } },
        },
      ],
      tool_choice: "auto",
      parallel_tool_calls: true,
    };

    const responsePayload = {
      id: "chatcmpl-test",
      object: "chat.completion",
      model: "gpt-5.1-2025-11-13",
      choices: [
        {
          index: 0,
          finish_reason: "tool_calls",
          message: {
            role: "assistant",
            content: null,
            tool_calls: [
              {
                id: "call_1",
                type: "function",
                function: {
                  name: "write_to_file",
                  arguments: "{\"path\":\"tests/agent-tests/fizz-buzz.ts\",\"content\":\"x\"}",
                },
              },
            ],
          },
        },
      ],
    };

    const assistantText = [
      "```json",
      JSON.stringify(echoedRequest, null, 2),
      "```",
      "```json",
      JSON.stringify(responsePayload, null, 2),
      "```",
    ].join("\n");

    const extracted = tryExtractSimulatedResponsePayload(
      assistantText,
      "chat.completions",
    );
    expect(extracted).not.toBeNull();
    expect(extracted?.id).toBe("chatcmpl-test");
    expect(Array.isArray(extracted?.choices)).toBeTrue();
  });

  test("returns null when only request-like JSON is present", () => {
    const requestOnly = {
      model: "gpt-5.1-2025-11-13",
      messages: [{ role: "user", content: "hello" }],
      tools: [
        {
          type: "function",
          function: { name: "write_to_file", parameters: { type: "object" } },
        },
      ],
      tool_choice: "required",
      parallel_tool_calls: true,
    };

    const assistantText = `\`\`\`json\n${JSON.stringify(requestOnly, null, 2)}\n\`\`\``;
    const extracted = tryExtractSimulatedResponsePayload(
      assistantText,
      "chat.completions",
    );
    expect(extracted).toBeNull();
  });
});

describe("tryExtractIncrementalSimulatedChatContent", () => {
  test("extracts partial content from incomplete simulated chat completion JSON", () => {
    const partial = [
      "```json",
      "{",
      '  "id": "chatcmpl_test",',
      '  "choices": [',
      "    {",
      '      "message": {',
      '        "role": "assistant",',
      '        "content": "hello\\nwor',
    ].join("\n");

    const extracted = tryExtractIncrementalSimulatedChatContent(partial);
    expect(extracted.hasToolCalls).toBeFalse();
    expect(extracted.content).toBe("hello\nwor");
  });

  test("suppresses incremental content when tool_calls appears before content", () => {
    const partial = [
      "```json",
      "{",
      '  "id": "chatcmpl_test",',
      '  "choices": [',
      "    {",
      '      "message": {',
      '        "role": "assistant",',
      '        "tool_calls": [',
      "          {",
      '            "id": "call_1",',
      '            "type": "function",',
      "            \"function\": {",
      '              "name": "attempt_completion",',
      "              \"arguments\": \"{\\\"result\\\":\\\"x\\\"}\"",
      "            }",
      "          }",
      "        ]",
    ].join("\n");

    const extracted = tryExtractIncrementalSimulatedChatContent(partial);
    expect(extracted.hasToolCalls).toBeTrue();
    expect(extracted.content).toBeNull();
  });
});
