import { randomUUID } from "node:crypto";
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";
import readline from "node:readline/promises";
import {
  BoxRenderable,
  createCliRenderer,
  InputRenderable,
  InputRenderableEvents,
  TextRenderable,
} from "@opentui/core";
import { readSseEvents, tryParseJsonObject } from "../proxy/utils";

type ParsedArgs = {
  positionals: string[];
  options: Record<string, string | null>;
};

type TokenState = {
  token: string;
  expiresAtUtc: string;
};

type TokenSummary = {
  state: "not set" | "valid" | "expired";
  expiry: string;
};

type ApiMode = "completions" | "responses";

type SessionState = {
  apiMode: ApiMode;
  conversationId: string | null;
  previousResponseId: string | null;
};

const parsed = parseArgs(process.argv.slice(2));
const command = parsed.positionals[0]?.toLowerCase() ?? "chat";

const exitCode =
  command === "chat"
    ? await runChatCommand(parsed.options)
    : command === "status"
      ? await runStatusCommand(parsed.options)
      : command === "token"
        ? await runTokenCommand(parsed)
        : command === "help" || command === "h" || command === "--help"
          ? showUsage()
          : showUnknownCommand(command);

process.exit(exitCode);

function showUsage(): number {
  console.log("YarpPilot CLI (Bun)");
  console.log(
    'Usage: bun src/cli/index.ts chat [--message "..."] [--token "..."] [--proxy "http://localhost:4000"] [--model "m365-copilot"] [--api "completions|responses"] [--responses]',
  );
  console.log(
    '       bun src/cli/index.ts status [--proxy "http://localhost:4000"]',
  );
  console.log('       bun src/cli/index.ts token set [--token "..."]');
  console.log("       bun src/cli/index.ts token clear");
  console.log("       bun src/cli/index.ts token status");
  return 0;
}

function showUnknownCommand(command: string): number {
  console.error(`Unknown command: ${command}`);
  return showUsage() === 0 ? 1 : 1;
}

async function runTokenCommand(parsedArgs: ParsedArgs): Promise<number> {
  const sub = parsedArgs.positionals[1]?.toLowerCase() ?? "status";
  const tokenPath = await getTokenPath();
  if (sub === "set") {
    const provided = firstNonEmpty(
      parsedArgs.options.token,
      process.env.YARPILOT_TOKEN,
    );
    const parsedToken = provided
      ? parseTokenOrThrow(provided)
      : await promptForTokenInteractive();
    await saveToken(tokenPath, parsedToken.token, parsedToken.expiresAtUtc);
    console.log(
      `Saved token. Expires: ${parsedToken.expiresAtUtc.toISOString()}`,
    );
    console.log(`Path: ${tokenPath}`);
    return 0;
  }

  if (sub === "clear") {
    const deleted = await deleteToken(tokenPath);
    console.log(deleted ? "Cleared saved token." : "No saved token to clear.");
    console.log(`Path: ${tokenPath}`);
    return 0;
  }

  if (sub === "status") {
    const tokenState = await loadToken(tokenPath);
    const summary = buildTokenSummary(tokenState);
    console.log(`Path: ${tokenPath}`);
    console.log(`State: ${summary.state}`);
    console.log(`Expiry: ${summary.expiry}`);
    return 0;
  }

  return showUnknownCommand(`token ${sub}`);
}

async function runStatusCommand(
  options: Record<string, string | null>,
): Promise<number> {
  const proxy = options.proxy ?? "http://localhost:4000";
  const status = await getStatusInfo(proxy);

  console.log("YarpPilot Status");
  console.log(`Proxy: ${status.proxy}`);
  console.log(
    `Proxy health: ${status.proxyStatus}${
      status.proxyDetails ? ` (${status.proxyDetails})` : ""
    }`,
  );
  console.log(`Token store: ${status.tokenPath}`);
  console.log(`Token state: ${status.tokenSummary.state}`);
  console.log(`Token expiry: ${status.tokenSummary.expiry}`);
  return 0;
}

async function getStatusInfo(proxy: string): Promise<{
  proxy: string;
  proxyStatus: string;
  proxyDetails: string;
  tokenPath: string;
  tokenSummary: TokenSummary;
}> {
  const tokenPath = await getTokenPath();
  const tokenState = await loadToken(tokenPath);
  const tokenSummary = buildTokenSummary(tokenState);

  let proxyStatus = "unreachable";
  let proxyDetails = "";
  try {
    const healthUrl = new URL(
      "healthz",
      proxy.endsWith("/") ? proxy : `${proxy}/`,
    );
    const response = await fetch(healthUrl, { method: "GET" });
    proxyStatus = response.ok ? "ok" : "error";
    proxyDetails = `HTTP ${response.status}`;
  } catch (error) {
    proxyDetails = String(error);
  }

  return { proxy, proxyStatus, proxyDetails, tokenPath, tokenSummary };
}

async function runChatCommand(
  options: Record<string, string | null>,
): Promise<number> {
  const proxy = options.proxy ?? "http://localhost:4000";
  const model = options.model ?? "m365-copilot";
  let apiMode: ApiMode;
  try {
    apiMode = resolveApiMode(options);
  } catch (error) {
    console.error(String(error));
    return 1;
  }
  let oneShotMessage = options.message;
  const providedToken = firstNonEmpty(
    options.token,
    process.env.YARPILOT_TOKEN,
  );

  const tokenPath = await getTokenPath();
  const cachedToken = await loadToken(tokenPath);
  let token = await ensureValidToken(cachedToken, tokenPath, providedToken);

  if (oneShotMessage?.trim()) {
    const oneShotSession: SessionState = {
      apiMode,
      conversationId: null,
      previousResponseId: null,
    };
    const result = await sendChatTurn(
      proxy,
      token.token,
      model,
      oneShotMessage,
      oneShotSession,
      () => {},
    );
    if (result.errorMessage) {
      console.error(`Error: ${result.errorMessage}`);
      return 1;
    }
    return 0;
  }

  if (process.stdin.isTTY) {
    return runChatTui(proxy, model, tokenPath, token, apiMode);
  }

  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
    terminal: false,
  });

  const session: SessionState = {
    apiMode,
    conversationId: null,
    previousResponseId: null,
  };
  while (true) {
    const line = await rl.question("");
    if (!line || !line.trim()) {
      continue;
    }
    const prompt = line.trim();
    if (prompt.toLowerCase() === "exit" || prompt.toLowerCase() === "quit") {
      break;
    }
    if (prompt.startsWith("/")) {
      const handled = await handleSlashCommand(
        prompt,
        proxy,
        tokenPath,
        token,
        session.apiMode,
        (text) => {
          process.stdout.write(`${text}\n`);
        },
      );
      token = handled.token;
      session.apiMode = handled.apiMode;
      if (handled.didExit) {
        break;
      }
      continue;
    }
    if (
      !token.token.trim() ||
      token.expiresAtUtc.getTime() <= Date.now() + 60_000
    ) {
      token = await ensureValidToken(
        await loadToken(tokenPath),
        tokenPath,
        null,
      );
    }
    const result = await sendChatTurn(
      proxy,
      token.token,
      model,
      prompt,
      session,
      (delta) => {
        process.stdout.write(delta);
      },
    );
    process.stdout.write("\n");
    if (result.errorMessage) {
      console.error(`Error: ${result.errorMessage}`);
      if (result.isAuthError) {
        token = await promptForTokenInteractive();
        await saveToken(tokenPath, token.token, token.expiresAtUtc);
      }
    } else {
      session.conversationId = result.conversationId ?? session.conversationId;
      session.previousResponseId =
        result.responseId ?? session.previousResponseId;
    }
  }
  rl.close();
  return 0;
}

async function runChatTui(
  proxy: string,
  model: string,
  tokenPath: string,
  initialToken: { token: string; expiresAtUtc: Date },
  initialApiMode: ApiMode,
): Promise<number> {
  const renderer = await createCliRenderer({
    exitOnCtrlC: false,
    useMouse: false,
  });
  const root = new BoxRenderable(renderer, {
    width: "100%",
    height: "100%",
    flexDirection: "column",
    padding: 1,
    gap: 1,
  });
  const header = new TextRenderable(renderer, {
    content:
      "YarpPilot CLI (OpenTUI) - /status /api [completions|responses] /token /cleartoken /exit",
  });
  const transcriptPanel = new BoxRenderable(renderer, {
    flexGrow: 1,
    border: true,
    padding: 1,
  });
  const transcript = new TextRenderable(renderer, { content: "" });
  const inputPanel = new BoxRenderable(renderer, {
    border: true,
    paddingLeft: 1,
    paddingRight: 1,
  });
  const input = new InputRenderable(renderer, {
    placeholder: "Ask Copilot...",
    value: "",
  });
  const status = new TextRenderable(renderer, {
    content: `Proxy: ${proxy} | API: ${initialApiMode}`,
  });

  transcriptPanel.add(transcript);
  root.add(header);
  root.add(transcriptPanel);
  root.add(status);
  inputPanel.add(input);
  root.add(inputPanel);
  renderer.root.add(root);
  renderer.start();
  input.focus();

  let token = initialToken;
  const session: SessionState = {
    apiMode: initialApiMode,
    conversationId: null,
    previousResponseId: null,
  };
  let busy = false;
  let closed = false;
  let output = "";

  const appendOutput = (text: string) => {
    output += text;
    transcript.content = output;
    renderer.requestRender();
  };

  const setStatus = (text: string) => {
    status.content = text;
    renderer.requestRender();
  };

  const shutdown = async () => {
    if (closed) {
      return;
    }
    closed = true;
    renderer.destroy();
  };

  renderer.addInputHandler((sequence) => {
    if (sequence === "\u0003") {
      shutdown().catch(() => {});
      return true;
    }
    return false;
  });

  input.on(InputRenderableEvents.ENTER, async () => {
    if (busy || closed) {
      return;
    }
    const prompt = input.value.trim();
    input.value = "";
    if (!prompt) {
      return;
    }
    if (prompt.startsWith("/")) {
      const handled = await handleSlashCommand(
        prompt,
        proxy,
        tokenPath,
        token,
        session.apiMode,
        (text) => appendOutput(`${text}\n`),
        setStatus,
      );
      token = handled.token;
      session.apiMode = handled.apiMode;
      if (handled.didExit) {
        await shutdown();
      }
      return;
    }

    if (
      !token.token.trim() ||
      token.expiresAtUtc.getTime() <= Date.now() + 60_000
    ) {
      token = await ensureValidToken(
        await loadToken(tokenPath),
        tokenPath,
        null,
      );
    }

    busy = true;
    appendOutput(`\nYou: ${prompt}\nCopilot: `);
    setStatus(`Waiting for response... API: ${session.apiMode}`);

    const result = await sendChatTurn(
      proxy,
      token.token,
      model,
      prompt,
      session,
      (delta) => {
        appendOutput(delta);
      },
    );

    appendOutput("\n");
    if (result.errorMessage) {
      appendOutput(`Error: ${result.errorMessage}\n`);
      setStatus("Request failed.");
      if (result.isAuthError) {
        try {
          token = await promptForTokenInteractive();
          await saveToken(tokenPath, token.token, token.expiresAtUtc);
          setStatus("Token refreshed.");
        } catch (error) {
          appendOutput(`Token prompt failed: ${String(error)}\n`);
        }
      }
    } else {
      session.conversationId = result.conversationId ?? session.conversationId;
      session.previousResponseId =
        result.responseId ?? session.previousResponseId;
      setStatus(formatSessionStatus(proxy, session));
    }
    busy = false;
  });

  await renderer.idle();
  while (!closed) {
    await new Promise((resolve) => setTimeout(resolve, 50));
  }
  return 0;
}

async function handleSlashCommand(
  raw: string,
  proxy: string,
  tokenPath: string,
  token: { token: string; expiresAtUtc: Date },
  apiMode: ApiMode,
  writeLine: (text: string) => void,
  setStatus?: (text: string) => void,
): Promise<{
  didExit: boolean;
  token: { token: string; expiresAtUtc: Date };
  apiMode: ApiMode;
}> {
  const command = raw.trim().toLowerCase();
  if (command === "/exit" || command === "/quit") {
    return { didExit: true, token, apiMode };
  }

  if (command === "/status") {
    const status = await getStatusInfo(proxy);
    writeLine("YarpPilot Status");
    writeLine(`Proxy: ${status.proxy}`);
    writeLine(
      `Proxy health: ${status.proxyStatus}${
        status.proxyDetails ? ` (${status.proxyDetails})` : ""
      }`,
    );
    writeLine(`Token store: ${status.tokenPath}`);
    writeLine(`Token state: ${status.tokenSummary.state}`);
    writeLine(`Token expiry: ${status.tokenSummary.expiry}`);
    writeLine(`API mode: ${apiMode}`);
    setStatus?.(
      `Proxy: ${status.proxyStatus}${
        status.proxyDetails ? ` (${status.proxyDetails})` : ""
      } | Token: ${status.tokenSummary.state} | API: ${apiMode}`,
    );
    return { didExit: false, token, apiMode };
  }

  if (command === "/api") {
    writeLine(`API mode: ${apiMode}`);
    setStatus?.(`API mode: ${apiMode}`);
    return { didExit: false, token, apiMode };
  }

  if (command.startsWith("/api ")) {
    const requested = command.slice("/api ".length).trim();
    if (requested === "completions" || requested === "responses") {
      writeLine(`Switched API mode to: ${requested}`);
      setStatus?.(`API mode: ${requested}`);
      return { didExit: false, token, apiMode: requested };
    }
    writeLine("Usage: /api completions | /api responses");
    return { didExit: false, token, apiMode };
  }

  if (command === "/token") {
    try {
      const parsedToken = await promptForTokenInteractive();
      await saveToken(tokenPath, parsedToken.token, parsedToken.expiresAtUtc);
      writeLine(
        `Saved token. Expires: ${parsedToken.expiresAtUtc.toISOString()}`,
      );
      setStatus?.("Token updated.");
      return { didExit: false, token: parsedToken, apiMode };
    } catch (error) {
      writeLine(`Token update failed: ${String(error)}`);
      return { didExit: false, token, apiMode };
    }
  }

  if (command === "/cleartoken") {
    const deleted = await deleteToken(tokenPath);
    writeLine(deleted ? "Cleared saved token." : "No saved token to clear.");
    setStatus?.("Token cleared.");
    return {
      didExit: false,
      token: { token: "", expiresAtUtc: new Date(0) },
      apiMode,
    };
  }

  writeLine(`Unknown command: ${raw}`);
  return { didExit: false, token, apiMode };
}

async function sendChatTurn(
  proxyBaseUrl: string,
  token: string,
  model: string,
  prompt: string,
  session: SessionState,
  onDelta: (text: string) => void,
): Promise<{
  conversationId: string | null;
  responseId: string | null;
  errorMessage: string | null;
  isAuthError: boolean;
}> {
  return session.apiMode === "responses"
    ? sendResponsesTurn(proxyBaseUrl, token, model, prompt, session, onDelta)
    : sendCompletionsTurn(proxyBaseUrl, token, model, prompt, session, onDelta);
}

async function sendCompletionsTurn(
  proxyBaseUrl: string,
  token: string,
  model: string,
  prompt: string,
  session: SessionState,
  onDelta: (text: string) => void,
): Promise<{
  conversationId: string | null;
  responseId: string | null;
  errorMessage: string | null;
  isAuthError: boolean;
}> {
  const requestUrl = new URL(
    "/v1/chat/completions",
    proxyBaseUrl.endsWith("/") ? proxyBaseUrl : `${proxyBaseUrl}/`,
  );
  const headers = new Headers({
    Authorization: `Bearer ${token}`,
    "Content-Type": "application/json",
    "x-m365-transport": "substrate",
  });
  if (session.conversationId) {
    headers.set("x-m365-conversation-id", session.conversationId);
  }

  const body = JSON.stringify({
    model,
    stream: true,
    messages: [{ role: "user", content: prompt }],
  });

  const response = await fetch(requestUrl, {
    method: "POST",
    headers,
    body,
  });

  const returnedConversationId = response.headers.get("x-m365-conversation-id");
  if (!response.ok) {
    const errorBody = await response.text();
    return {
      conversationId: returnedConversationId,
      responseId: null,
      errorMessage: extractErrorMessage(errorBody) ?? `HTTP ${response.status}`,
      isAuthError: response.status === 401 || response.status === 403,
    };
  }

  if (!response.body) {
    return {
      conversationId: returnedConversationId,
      responseId: null,
      errorMessage: null,
      isAuthError: false,
    };
  }

  for await (const event of readSseEvents(response.body)) {
    const data = event.data.trim();
    if (!data) {
      continue;
    }
    if (data.toLowerCase() === "[done]") {
      break;
    }
    if (event.event.toLowerCase() === "error") {
      return {
        conversationId: returnedConversationId,
        responseId: null,
        errorMessage: extractErrorMessage(data) ?? data,
        isAuthError: false,
      };
    }
    const delta = extractCompletionsDeltaContent(data);
    if (delta) {
      onDelta(delta);
    }
  }

  return {
    conversationId: returnedConversationId,
    responseId: null,
    errorMessage: null,
    isAuthError: false,
  };
}

async function sendResponsesTurn(
  proxyBaseUrl: string,
  token: string,
  model: string,
  prompt: string,
  session: SessionState,
  onDelta: (text: string) => void,
): Promise<{
  conversationId: string | null;
  responseId: string | null;
  errorMessage: string | null;
  isAuthError: boolean;
}> {
  const requestUrl = new URL(
    "/v1/responses",
    proxyBaseUrl.endsWith("/") ? proxyBaseUrl : `${proxyBaseUrl}/`,
  );
  const headers = new Headers({
    Authorization: `Bearer ${token}`,
    "Content-Type": "application/json",
    "x-m365-transport": "substrate",
  });
  if (session.conversationId) {
    headers.set("x-m365-conversation-id", session.conversationId);
  }

  const requestBody: Record<string, unknown> = {
    model,
    stream: true,
    input: prompt,
  };
  if (session.previousResponseId) {
    requestBody.previous_response_id = session.previousResponseId;
  }

  const response = await fetch(requestUrl, {
    method: "POST",
    headers,
    body: JSON.stringify(requestBody),
  });

  const returnedConversationId = response.headers.get("x-m365-conversation-id");
  if (!response.ok) {
    const errorBody = await response.text();
    return {
      conversationId: returnedConversationId,
      responseId: null,
      errorMessage: extractErrorMessage(errorBody) ?? `HTTP ${response.status}`,
      isAuthError: response.status === 401 || response.status === 403,
    };
  }

  if (!response.body) {
    return {
      conversationId: returnedConversationId,
      responseId: null,
      errorMessage: null,
      isAuthError: false,
    };
  }

  let responseId: string | null = null;
  let streamConversationId: string | null = returnedConversationId;

  for await (const event of readSseEvents(response.body)) {
    const data = event.data.trim();
    if (!data) {
      continue;
    }
    if (data.toLowerCase() === "[done]") {
      break;
    }
    if (event.event.toLowerCase() === "error") {
      return {
        conversationId: streamConversationId,
        responseId,
        errorMessage: extractErrorMessage(data) ?? data,
        isAuthError: false,
      };
    }

    responseId = extractResponsesResponseId(data) ?? responseId;
    streamConversationId =
      extractResponsesConversationId(data) ?? streamConversationId;
    const delta = extractResponsesDeltaContent(data);
    if (delta) {
      onDelta(delta);
    }
  }

  return {
    conversationId: streamConversationId,
    responseId,
    errorMessage: null,
    isAuthError: false,
  };
}

function extractCompletionsDeltaContent(rawChunk: string): string | null {
  const json = tryParseJsonObject(rawChunk);
  const choices = json?.choices;
  if (!Array.isArray(choices) || choices.length === 0) {
    return null;
  }
  const first = choices[0];
  if (!first || typeof first !== "object" || Array.isArray(first)) {
    return null;
  }
  const delta = (first as Record<string, unknown>).delta;
  if (!delta || typeof delta !== "object" || Array.isArray(delta)) {
    return null;
  }
  const content = (delta as Record<string, unknown>).content;
  return typeof content === "string" && content.length > 0 ? content : null;
}

function extractResponsesDeltaContent(rawChunk: string): string | null {
  const json = tryParseJsonObject(rawChunk);
  if (!json) {
    return null;
  }
  if (json.type !== "response.output_text.delta") {
    return null;
  }
  return typeof json.delta === "string" && json.delta.length > 0
    ? json.delta
    : null;
}

function extractResponsesResponseId(rawChunk: string): string | null {
  const json = tryParseJsonObject(rawChunk);
  if (!json) {
    return null;
  }
  if (typeof json.response_id === "string" && json.response_id.trim()) {
    return json.response_id.trim();
  }
  const response = json.response;
  if (!response || typeof response !== "object" || Array.isArray(response)) {
    return null;
  }
  const id = (response as Record<string, unknown>).id;
  return typeof id === "string" && id.trim() ? id.trim() : null;
}

function extractResponsesConversationId(rawChunk: string): string | null {
  const json = tryParseJsonObject(rawChunk);
  if (!json) {
    return null;
  }
  const response = json.response;
  if (!response || typeof response !== "object" || Array.isArray(response)) {
    return null;
  }
  const conversationId = (response as Record<string, unknown>).conversation_id;
  return typeof conversationId === "string" && conversationId.trim()
    ? conversationId.trim()
    : null;
}

function extractErrorMessage(rawJson: string): string | null {
  const json = tryParseJsonObject(rawJson);
  const error = json?.error;
  if (error && typeof error === "object" && !Array.isArray(error)) {
    const message = (error as Record<string, unknown>).message;
    if (typeof message === "string" && message.trim()) {
      return message.trim();
    }
  }
  const direct = json?.message;
  return typeof direct === "string" && direct.trim() ? direct.trim() : null;
}

function parseArgs(args: string[]): ParsedArgs {
  const options: Record<string, string | null> = {};
  const positionals: string[] = [];

  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    if (!arg.startsWith("--")) {
      positionals.push(arg);
      continue;
    }
    const key = arg.slice(2);
    let value: string | null = null;
    if (i + 1 < args.length && !args[i + 1].startsWith("--")) {
      value = args[++i] ?? null;
    }
    options[key] = value;
  }

  return { positionals, options };
}

function resolveApiMode(options: Record<string, string | null>): ApiMode {
  if ("responses" in options) {
    return "responses";
  }
  if ("completions" in options) {
    return "completions";
  }
  const raw = firstNonEmpty(options.api, options.endpoint, options.mode);
  if (!raw) {
    return "completions";
  }
  const normalized = raw.trim().toLowerCase();
  if (
    normalized === "responses" ||
    normalized === "response" ||
    normalized === "v1/responses"
  ) {
    return "responses";
  }
  if (
    normalized === "completions" ||
    normalized === "completion" ||
    normalized === "chat/completions" ||
    normalized === "v1/chat/completions"
  ) {
    return "completions";
  }
  throw new Error(
    `Invalid API mode '${raw}'. Use --api completions or --api responses.`,
  );
}

function formatSessionStatus(proxy: string, session: SessionState): string {
  const segments = [`Proxy: ${proxy}`, `API: ${session.apiMode}`];
  if (session.conversationId) {
    segments.push(`conversation: ${session.conversationId}`);
  }
  if (session.previousResponseId) {
    segments.push(`response: ${session.previousResponseId}`);
  }
  return segments.join(" | ");
}

function firstNonEmpty(
  ...values: Array<string | null | undefined>
): string | null {
  for (const value of values) {
    if (value && value.trim()) {
      return value;
    }
  }
  return null;
}

async function getTokenPath(): Promise<string> {
  const localAppData =
    process.env.LOCALAPPDATA ?? path.join(os.homedir(), ".local", "share");
  const directory = path.join(localAppData, "YarpPilot", "Cli");
  await fs.mkdir(directory, { recursive: true });
  return path.join(directory, "token.json");
}

async function loadToken(filePath: string): Promise<TokenState | null> {
  try {
    const content = await fs.readFile(filePath, "utf8");
    const parsed = JSON.parse(content) as {
      token?: string;
      expiresAtUtc?: string;
    };
    if (!parsed.token?.trim() || !parsed.expiresAtUtc?.trim()) {
      return null;
    }
    const expires = new Date(parsed.expiresAtUtc);
    if (Number.isNaN(expires.getTime())) {
      return null;
    }
    return { token: parsed.token.trim(), expiresAtUtc: expires.toISOString() };
  } catch {
    return null;
  }
}

async function saveToken(
  filePath: string,
  token: string,
  expiresAtUtc: Date,
): Promise<void> {
  await fs.writeFile(
    filePath,
    JSON.stringify(
      {
        token,
        expiresAtUtc: expiresAtUtc.toISOString(),
      },
      null,
      2,
    ),
    "utf8",
  );
}

async function deleteToken(filePath: string): Promise<boolean> {
  try {
    await fs.unlink(filePath);
    return true;
  } catch {
    return false;
  }
}

function buildTokenSummary(tokenState: TokenState | null): TokenSummary {
  if (!tokenState) {
    return { state: "not set", expiry: "n/a" };
  }
  const expires = new Date(tokenState.expiresAtUtc);
  const remainingMs = expires.getTime() - Date.now();
  if (remainingMs > 60_000) {
    return { state: "valid", expiry: expires.toISOString() };
  }
  return { state: "expired", expiry: expires.toISOString() };
}

async function ensureValidToken(
  tokenState: TokenState | null,
  tokenPath: string,
  providedToken: string | null,
): Promise<{ token: string; expiresAtUtc: Date }> {
  if (providedToken?.trim()) {
    const parsed = parseTokenOrThrow(providedToken);
    await saveToken(tokenPath, parsed.token, parsed.expiresAtUtc);
    return parsed;
  }

  if (tokenState?.token?.trim()) {
    const expires = new Date(tokenState.expiresAtUtc);
    if (expires.getTime() > Date.now() + 60_000) {
      return { token: tokenState.token, expiresAtUtc: expires };
    }
  }

  const prompted = await promptForTokenInteractive();
  await saveToken(tokenPath, prompted.token, prompted.expiresAtUtc);
  return prompted;
}

function parseTokenOrThrow(rawToken: string): {
  token: string;
  expiresAtUtc: Date;
} {
  const normalized = normalizeToken(rawToken);
  if (!normalized) {
    throw new Error("Token cannot be empty.");
  }
  const expiresAtUtc =
    tryGetJwtExpiry(normalized) ?? new Date(Date.now() + 60 * 60 * 1000);
  return { token: normalized, expiresAtUtc };
}

async function promptForTokenInteractive(): Promise<{
  token: string;
  expiresAtUtc: Date;
}> {
  if (!process.stdin.isTTY) {
    throw new Error(
      "No valid cached token and no interactive terminal. Pass --token or set YARPILOT_TOKEN.",
    );
  }
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });
  try {
    const raw = await rl.question(
      "Paste Microsoft Graph/Substrate bearer token: ",
    );
    return parseTokenOrThrow(raw);
  } finally {
    rl.close();
  }
}

function normalizeToken(raw: string): string {
  const trimmed = raw.trim();
  return trimmed.toLowerCase().startsWith("bearer ")
    ? trimmed.slice("Bearer ".length).trim()
    : trimmed;
}

function tryGetJwtExpiry(token: string): Date | null {
  if (!token.trim()) {
    return null;
  }
  const parts = token.split(".");
  if (parts.length < 2) {
    return null;
  }
  try {
    const payload = Buffer.from(
      base64UrlNormalize(parts[1]),
      "base64",
    ).toString("utf8");
    const parsed = JSON.parse(payload) as { exp?: number | string };
    const expRaw = parsed.exp;
    const exp =
      typeof expRaw === "number"
        ? expRaw
        : typeof expRaw === "string"
          ? Number.parseInt(expRaw, 10)
          : Number.NaN;
    if (!Number.isFinite(exp)) {
      return null;
    }
    return new Date(exp * 1000);
  } catch {
    return null;
  }
}

function base64UrlNormalize(encoded: string): string {
  const normalized = encoded.replaceAll("-", "+").replaceAll("_", "/");
  const padding = normalized.length % 4;
  return padding > 0
    ? normalized.padEnd(normalized.length + (4 - padding), "=")
    : normalized;
}
