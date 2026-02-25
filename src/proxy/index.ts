import { parseArgs } from "util";
import { promises as fs } from "node:fs";
import path from "node:path";
import { CopilotGraphClient, CopilotSubstrateClient } from "./clients";
import { loadWrapperOptions } from "./config";
import { ConversationStore } from "./conversation-store";
import { DebugMarkdownLogger } from "./logger";
import { ResponseStore } from "./response-store";
import { createProxyApp } from "./server";
import { ProxyTokenProvider } from "./token-provider";
import type { WrapperOptions } from "./types";
import { parseListenUrl } from "./utils";

const loadedOptions = await loadWrapperOptions(process.cwd());
const debugEnabled = parseDebugFlag();
const options = await withSessionDebugPath(loadedOptions, debugEnabled);
if (debugEnabled && options.debugPath?.trim()) {
  await fs.mkdir(options.debugPath, { recursive: true });
}
const debugLogger = new DebugMarkdownLogger(options, debugEnabled);
const graphClient = new CopilotGraphClient(options, debugLogger);
const substrateClient = new CopilotSubstrateClient(options, debugLogger);
const conversationStore = new ConversationStore(options);
const responseStore = new ResponseStore(options);
const tokenProvider = new ProxyTokenProvider({
  ignoreIncomingAuthorizationHeader: options.ignoreIncomingAuthorizationHeader,
});

const app = createProxyApp({
  options,
  debugLogger,
  graphClient,
  substrateClient,
  conversationStore,
  responseStore,
  tokenProvider,
});

const listen = parseListenUrl(options.listenUrl);
const debugLogPath =
  debugEnabled && options.debugPath?.trim()
    ? path.resolve(options.debugPath)
    : null;
const server = Bun.serve({
  hostname: listen.hostname,
  port: listen.port,
  fetch: app.fetch,
});

console.log(
  `m365-copilot-bun-proxy listening on http://${server.hostname}:${server.port} (from ${options.listenUrl})`,
);
console.log(
  debugEnabled
    ? `Debugging: enabled (level: ${options.logLevel}, logs at ${debugLogPath})`
    : `Debugging: disabled (configured level: ${options.logLevel})`,
);

function parseDebugFlag(): boolean {
  const { values } = parseArgs({
    args: Bun.argv,
    options: {
      debug: {
        type: "boolean",
        short: "d",
      },
    },
    strict: true,
    allowPositionals: true,
  });

  return values.debug ?? false;
}

async function withSessionDebugPath(
  options: WrapperOptions,
  debugEnabled: boolean,
): Promise<WrapperOptions> {
  if (!debugEnabled || !options.debugPath?.trim()) {
    return options;
  }

  const basePath = options.debugPath.trim();
  const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
  const sessionFolder = await resolveSessionFolderName(basePath, timestamp);

  return {
    ...options,
    debugPath: path.join(basePath, sessionFolder),
  };
}

async function resolveSessionFolderName(
  basePath: string,
  timestamp: string,
): Promise<string> {
  const nextSequence = (await findMaxSessionPrefix(basePath)) + 1;
  for (let sequence = nextSequence; sequence < 1_000_000; sequence++) {
    const prefix = String(sequence).padStart(3, "0");
    const candidate = `${prefix}-${timestamp}`;
    if (!(await pathExists(path.join(basePath, candidate)))) {
      return candidate;
    }
  }

  throw new Error(
    `Unable to allocate a unique debug session folder under ${basePath}`,
  );
}

async function findMaxSessionPrefix(basePath: string): Promise<number> {
  try {
    const entries = await fs.readdir(basePath, { withFileTypes: true });
    let max = 0;
    for (const entry of entries) {
      if (!entry.isDirectory()) {
        continue;
      }
      const match = /^(\d+)-/.exec(entry.name);
      if (!match) {
        continue;
      }
      const value = Number.parseInt(match[1], 10);
      if (Number.isFinite(value) && value > max) {
        max = value;
      }
    }
    return max;
  } catch {
    return 0;
  }
}

async function pathExists(targetPath: string): Promise<boolean> {
  try {
    await fs.access(targetPath);
    return true;
  } catch {
    return false;
  }
}
