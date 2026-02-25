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
const options = withSessionDebugPath(loadedOptions, debugEnabled);
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

function withSessionDebugPath(
  options: WrapperOptions,
  debugEnabled: boolean,
): WrapperOptions {
  if (!debugEnabled || !options.debugPath?.trim()) {
    return options;
  }

  const basePath = options.debugPath.trim();
  const timestamp = new Date().toISOString().replace(/[:.]/g, "-");

  return {
    ...options,
    debugPath: path.join(basePath, timestamp),
  };
}
