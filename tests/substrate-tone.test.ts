import { describe, expect, test } from "bun:test";
import { resolveSubstrateTone } from "../src/proxy/clients";

describe("resolveSubstrateTone", () => {
  test("maps quick model to Chat tone", () => {
    expect(resolveSubstrateTone("m365-copilot-quick")).toBe("Chat");
  });

  test("maps reasoning model to Reasoning tone", () => {
    expect(resolveSubstrateTone("m365-copilot-reasoning")).toBe("Reasoning");
  });

  test("maps gpt 5.2 quick model to Gpt_5_2_Chat tone", () => {
    expect(resolveSubstrateTone("m365-copilot-gpt5.2-quick")).toBe(
      "Gpt_5_2_Chat",
    );
  });

  test("maps gpt 5.2 reasoning model to Gpt_5_2_Reasoning tone", () => {
    expect(resolveSubstrateTone("m365-copilot-gpt5.2-reasoning")).toBe(
      "Gpt_5_2_Reasoning",
    );
  });

  test("maps magic-family models to magic tone", () => {
    expect(resolveSubstrateTone("m365-copilot")).toBe("magic");
    expect(resolveSubstrateTone("m365-copilot-auto")).toBe("magic");
    expect(resolveSubstrateTone("m365-copilot-magic")).toBe("magic");
  });

  test("defaults unknown and empty models to magic tone", () => {
    expect(resolveSubstrateTone("some-unknown-model")).toBe("magic");
    expect(resolveSubstrateTone("")).toBe("magic");
    expect(resolveSubstrateTone(null)).toBe("magic");
    expect(resolveSubstrateTone(undefined)).toBe("magic");
  });

  test("is case-insensitive", () => {
    expect(resolveSubstrateTone("M365-COPILOT-GPT5.2-QUICK")).toBe(
      "Gpt_5_2_Chat",
    );
  });
});
