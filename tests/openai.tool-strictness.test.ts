import { describe, expect, test } from "bun:test";
import { buildAssistantResponse } from "../src/proxy/openai";
import { ToolChoiceModes, type ParsedOpenAiRequest } from "../src/proxy/types";

describe("buildAssistantResponse strict tool behavior", () => {
  test("returns strict error when tool_choice is function and no valid tool call is present", () => {
    const request = createRequest(ToolChoiceModes.Function, "get_time");
    const response = buildAssistantResponse(
      request,
      "I cannot call tools for this request.",
    );

    expect(response.toolCalls.length).toBe(0);
    expect(response.content).toBeNull();
    expect(response.strictToolErrorMessage).toBeString();
    expect(response.strictToolErrorMessage).toContain("get_time");
  });

  test("returns strict error when tool_choice is required and no valid tool call is present", () => {
    const request = createRequest(ToolChoiceModes.Required, null);
    const response = buildAssistantResponse(
      request,
      "Still no JSON tool call payload here.",
    );

    expect(response.toolCalls.length).toBe(0);
    expect(response.content).toBeNull();
    expect(response.strictToolErrorMessage).toBeString();
    expect(response.strictToolErrorMessage).toContain("tool_calls");
  });

  test("extracts tool call and clears strict error when valid JSON is present", () => {
    const request = createRequest(ToolChoiceModes.Function, "get_time");
    const response = buildAssistantResponse(
      request,
      '{"tool_calls":[{"name":"get_time","arguments":{"zone":"UTC"}}]}',
    );

    expect(response.strictToolErrorMessage).toBeNull();
    expect(response.finishReason).toBe("tool_calls");
    expect(response.toolCalls.length).toBe(1);
    expect(response.toolCalls[0]?.name).toBe("get_time");
    expect(response.content).toBeNull();
  });

  test("extracts tool call when invalid placeholder JSON appears before valid payload", () => {
    const request = createRequest(ToolChoiceModes.Function, "get_time");
    const response = buildAssistantResponse(
      request,
      [
        "Output shape: {\"tool_calls\":[{\"name\":\"get_time\",\"arguments\":{...}}]}",
        "Here is the correct payload:",
        "{\"tool_calls\":[{\"name\":\"get_time\",\"arguments\":{\"zone\":\"UTC\"}}]}",
      ].join("\n"),
    );

    expect(response.strictToolErrorMessage).toBeNull();
    expect(response.finishReason).toBe("tool_calls");
    expect(response.toolCalls.length).toBe(1);
    expect(response.toolCalls[0]?.name).toBe("get_time");
    expect(response.toolCalls[0]?.argumentsJson).toBe("{\"zone\":\"UTC\"}");
    expect(response.content).toBeNull();
  });
});

function createRequest(
  toolChoiceMode: string,
  toolChoiceFunctionName: string | null,
): ParsedOpenAiRequest {
  return {
    model: "m365-copilot",
    stream: false,
    promptText: "test",
    userKey: null,
    locationHint: { timeZone: "America/New_York" },
    contextualResources: null,
    additionalContext: [],
    tooling: {
      tools: [
        {
          name: "get_time",
          description: "Get time",
          parameters: {
            type: "object",
            properties: { zone: { type: "string" } },
            required: ["zone"],
          },
        },
      ],
      toolChoiceMode,
      toolChoiceFunctionName,
      parallelToolCalls: true,
    },
    responseFormat: null,
    reasoningEffort: null,
    temperature: null,
  };
}
