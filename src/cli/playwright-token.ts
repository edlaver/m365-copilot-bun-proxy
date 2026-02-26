import { spawn } from "node:child_process";
import { constants as fsConstants, promises as fs } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const NODE_RUNNER_FILENAME = "playwright-token.node.mjs";
const DEFAULT_PLAYWRIGHT_BROWSER = "edge";

export type PlaywrightBrowserName =
  | "edge"
  | "chrome"
  | "chromium"
  | "firefox"
  | "webkit";

export async function fetchTokenWithPlaywright(
  tokenPath: string,
  storageStatePath: string,
  options?: {
    quiet?: boolean;
    browser?: PlaywrightBrowserName;
  },
): Promise<void> {
  const runnerPath = await resolveNodeRunnerPath();
  await runNodePlaywrightFetch(
    runnerPath,
    tokenPath,
    storageStatePath,
    options?.browser ?? DEFAULT_PLAYWRIGHT_BROWSER,
    options?.quiet ?? false,
  );
}

async function resolveNodeRunnerPath(): Promise<string> {
  const moduleDir = path.dirname(fileURLToPath(import.meta.url));
  const candidates = [
    path.join(moduleDir, NODE_RUNNER_FILENAME),
    path.join(process.cwd(), "src", "cli", NODE_RUNNER_FILENAME),
    path.join(process.cwd(), "dist", NODE_RUNNER_FILENAME),
  ];

  for (const candidate of candidates) {
    try {
      await fs.access(candidate, fsConstants.F_OK);
      return candidate;
    } catch {
      // Try the next candidate.
    }
  }

  throw new Error(
    `Unable to locate ${NODE_RUNNER_FILENAME}. Checked:\n${candidates
      .map((entry) => `- ${entry}`)
      .join("\n")}`,
  );
}

function runNodePlaywrightFetch(
  runnerPath: string,
  tokenPath: string,
  storageStatePath: string,
  browser: PlaywrightBrowserName,
  quiet: boolean,
): Promise<void> {
  return new Promise((resolve, reject) => {
    const child = spawn(
      "node",
      [
        runnerPath,
        "--token-path",
        tokenPath,
        "--storage-state-path",
        storageStatePath,
        "--browser",
        browser,
      ],
      {
        stdio: "pipe",
        env: process.env,
        windowsHide: false,
      },
    );
    let output = "";
    const pushOutput = (chunk: string): void => {
      output += chunk;
      if (output.length > 32_768) {
        output = output.slice(output.length - 32_768);
      }
    };

    child.stdout?.on("data", (data) => {
      const chunk = String(data);
      pushOutput(chunk);
      if (!quiet) {
        process.stdout.write(chunk);
      }
    });

    child.stderr?.on("data", (data) => {
      const chunk = String(data);
      pushOutput(chunk);
      if (!quiet) {
        process.stderr.write(chunk);
      }
    });

    child.once("error", (error) => {
      reject(
        new Error(
          `Failed to launch Node.js Playwright runner: ${String(error)}`,
        ),
      );
    });

    child.once("exit", (code, signal) => {
      if (code === 0) {
        resolve();
        return;
      }
      reject(
        new Error(
          `Playwright runner exited with code ${String(code)}${signal ? ` (signal: ${signal})` : ""}.${output.trim() ? `\n${output.trim()}` : ""}`,
        ),
      );
    });
  });
}
