import { CopilotGraphClient, CopilotSubstrateClient } from "./clients";
import { loadWrapperOptions } from "./config";
import { ConversationStore } from "./conversation-store";
import { DebugMarkdownLogger } from "./logger";
import { createProxyApp } from "./server";
import { parseListenUrl } from "./utils";

const options = await loadWrapperOptions(process.cwd());
const debugLogger = new DebugMarkdownLogger(options);
const graphClient = new CopilotGraphClient(options, debugLogger);
const substrateClient = new CopilotSubstrateClient(options, debugLogger);
const conversationStore = new ConversationStore(options);

const app = createProxyApp({
  options,
  debugLogger,
  graphClient,
  substrateClient,
  conversationStore,
});

const listen = parseListenUrl(options.listenUrl);
const server = Bun.serve({
  hostname: listen.hostname,
  port: listen.port,
  fetch: app.fetch,
});

console.log(
  `m365-copilot-bun-proxy listening on http://${server.hostname}:${server.port} (from ${options.listenUrl})`,
);
