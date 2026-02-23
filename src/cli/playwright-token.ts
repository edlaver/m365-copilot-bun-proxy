import { spawn } from "node:child_process";
import { constants as fsConstants, promises as fs } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const NODE_RUNNER_FILENAME = "playwright-token.node.mjs";

export async function fetchTokenWithPlaywright(
  tokenPath: string,
  storageStatePath: string,
): Promise<void> {
  const runnerPath = await resolveNodeRunnerPath();
  await runNodePlaywrightFetch(runnerPath, tokenPath, storageStatePath);
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
      ],
      {
        stdio: "inherit",
        env: process.env,
        windowsHide: false,
      },
    );

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
          `Playwright runner exited with code ${String(code)}${signal ? ` (signal: ${signal})` : ""}.`,
        ),
      );
    });
  });
}
