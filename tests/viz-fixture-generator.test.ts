import { describe, expect, test } from "bun:test";
import { promises as fs } from "node:fs";
import path from "node:path";

describe("viz fixture generator output", () => {
  test("writes manifest and generated request json files", async () => {
    const generatedDir = path.join(
      process.cwd(),
      "src",
      "viz-tool",
      "src",
      "fixtures",
      "generated",
    );
    const manifestPath = path.join(generatedDir, "manifest.json");
    const manifestRaw = await fs.readFile(manifestPath, "utf8");
    const manifest = JSON.parse(manifestRaw) as Array<Record<string, unknown>>;

    expect(manifest.length).toBeGreaterThan(0);
    expect(manifest.some((entry) => entry.requestType === "chat/completions")).toBeTrue();
    expect(manifest.some((entry) => entry.requestType === "responses")).toBeTrue();

    const sample = manifest[0];
    expect(typeof sample?.fileName).toBe("string");
    const fixturePath = path.join(
      generatedDir,
      sample?.requestType === "responses" ? "responses" : "chat-completions",
      String(sample?.fileName),
    );
    const fixtureRaw = await fs.readFile(fixturePath, "utf8");
    const fixture = JSON.parse(fixtureRaw) as Record<string, unknown>;

    expect(typeof fixture.model).toBe("string");
  });
});
