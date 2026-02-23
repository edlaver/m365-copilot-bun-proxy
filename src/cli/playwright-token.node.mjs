import { promises as fs } from "node:fs";
import path from "node:path";
import { chromium } from "playwright";

const SUBSTRATE_WS_PATTERN = /substrate\.office\.com\/m365Copilot\/Chathub/i;
const CHAT_URL = "https://m365.cloud.microsoft/chat/?auth=2";
const CHAT_URL_GLOB = "**/chat/**";
const LOGIN_HOST_PATTERN = /login\.(microsoftonline|live|microsoft)\.com/i;

const TOKEN_TIMEOUT_MS = 120_000;
const LOGIN_TIMEOUT_MS = 300_000;

const parsed = parseArgs(process.argv.slice(2));
const tokenPath = parsed["token-path"];
const storageStatePath = parsed["storage-state-path"];

if (!tokenPath || !storageStatePath) {
  console.error(
    "Missing required args: --token-path <path> --storage-state-path <path>",
  );
  process.exit(2);
}

await fetchTokenWithPlaywrightNode(tokenPath, storageStatePath);

async function fetchTokenWithPlaywrightNode(tokenPath, storageStatePath) {
  console.log("[playwright] Launching Chromium under Node.js (headed)...");
  const browser = await launchBrowser();
  const storageStateExists = await fileExists(storageStatePath);
  const context = await browser.newContext(
    storageStateExists ? { storageState: storageStatePath } : {},
  );
  console.log(
    `[playwright] Browser launched (${storageStateExists ? "using saved storage state" : "fresh context"}).`,
  );

  try {
    const page = await context.newPage();
    await page.bringToFront().catch(() => {});
    console.log(`[playwright] Page URL: ${page.url()}`);

    const tokenPromise = captureSubstrateToken(page);

    console.log(`[playwright] Navigating to ${CHAT_URL}`);
    await page.goto(CHAT_URL, {
      waitUntil: "domcontentloaded",
      timeout: 60_000,
    });
    console.log(`[playwright] Landed on: ${page.url()}`);

    if (LOGIN_HOST_PATTERN.test(page.url())) {
      console.log("[playwright] Login required - sign in in the browser window.");
      await page.waitForURL(CHAT_URL_GLOB, { timeout: LOGIN_TIMEOUT_MS });
      console.log(`[playwright] Login complete: ${page.url()}`);
    } else {
      console.log("[playwright] Already logged in.");
    }

    try {
      const editor = page.locator("#m365-chat-editor-target-element");
      await editor.waitFor({ state: "visible", timeout: 20_000 });
      console.log("[playwright] Sending message to trigger WebSocket...");
      await editor.fill("Hi");
      await page.keyboard.press("Enter");
    } catch {
      console.log(
        "[playwright] Chat editor not found - waiting passively for WebSocket...",
      );
    }

    console.log(
      `[playwright] Waiting up to ${TOKEN_TIMEOUT_MS / 1000}s for token...`,
    );
    const rawToken = await tokenPromise;
    console.log("[playwright] Token captured!");

    const expiresAtUtc = tryGetJwtExpiry(rawToken) ?? new Date(Date.now() + 3_600_000);
    await saveToken(tokenPath, rawToken, expiresAtUtc);
    await fs.mkdir(path.dirname(storageStatePath), { recursive: true });
    await context.storageState({ path: storageStatePath });
    console.log(`[playwright] Browser state saved: ${storageStatePath}`);
    console.log(`Token saved. Expires: ${expiresAtUtc.toISOString()}`);
  } finally {
    await context?.close().catch(() => {});
    await browser?.close().catch(() => {});
  }
}

async function launchBrowser() {
  try {
    return await chromium.launch({
      headless: false,
      channel: "msedge",
      args: [
        "--no-first-run",
        "--no-default-browser-check",
        "--disable-search-engine-choice-screen",
      ],
    });
  } catch {
    // Fallback for environments without Edge channel support.
    return chromium.launch({
      headless: false,
      args: [
        "--no-first-run",
        "--no-default-browser-check",
        "--disable-search-engine-choice-screen",
      ],
    });
  }
}

function captureSubstrateToken(page) {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      reject(
        new Error(
          `Timed out waiting for Substrate WebSocket after ${TOKEN_TIMEOUT_MS / 1000}s. Try running 'token fetch' again.`,
        ),
      );
    }, TOKEN_TIMEOUT_MS);

    page.on("websocket", (ws) => {
      const url = ws.url();
      if (!SUBSTRATE_WS_PATTERN.test(url)) return;

      console.log(`[playwright] Substrate WebSocket detected: ${url.slice(0, 120)}`);
      try {
        const token = new URL(url).searchParams.get("access_token");
        if (token) {
          clearTimeout(timer);
          console.log("[playwright] access_token extracted.");
          resolve(token);
        }
      } catch {
        // Ignore parse failures from malformed websocket URLs.
      }
    });
  });
}

async function saveToken(filePath, token, expiresAtUtc) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
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

async function fileExists(filePath) {
  try {
    await fs.access(filePath);
    return true;
  } catch {
    return false;
  }
}

function tryGetJwtExpiry(token) {
  if (!token.trim()) {
    return null;
  }
  const parts = token.split(".");
  if (parts.length < 2) {
    return null;
  }

  try {
    const payload = Buffer.from(base64UrlNormalize(parts[1]), "base64").toString(
      "utf8",
    );
    const parsed = JSON.parse(payload);
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

function base64UrlNormalize(encoded) {
  const normalized = encoded.replaceAll("-", "+").replaceAll("_", "/");
  const padding = normalized.length % 4;
  return padding > 0
    ? normalized.padEnd(normalized.length + (4 - padding), "=")
    : normalized;
}

function parseArgs(args) {
  const options = {};
  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    if (!arg.startsWith("--")) continue;
    const key = arg.slice(2);
    if (i + 1 < args.length && !args[i + 1].startsWith("--")) {
      options[key] = args[++i];
    } else {
      options[key] = "";
    }
  }
  return options;
}
