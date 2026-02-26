import { promises as fs } from "node:fs";
import path from "node:path";
import { chromium, firefox, webkit } from "playwright";

const SUBSTRATE_WS_PATTERN = /substrate\.office\.com\/m365Copilot\/Chathub/i;
const CHAT_URL = "https://m365.cloud.microsoft/chat/?auth=2";
const CHAT_URL_GLOB = "**/chat/**";
const LOGIN_HOST_PATTERN = /login\.(microsoftonline|live|microsoft)\.com/i;
const CHROMIUM_LAUNCH_ARGS = [
  "--no-first-run",
  "--no-default-browser-check",
  "--disable-search-engine-choice-screen",
];
const SUPPORTED_BROWSERS = new Set([
  "edge",
  "chrome",
  "chromium",
  "firefox",
  "webkit",
]);

const TOKEN_TIMEOUT_MS = 120_000;
const LOGIN_TIMEOUT_MS = 300_000;

const parsed = parseArgs(process.argv.slice(2));
const tokenPath = parsed["token-path"];
const storageStatePath = parsed["storage-state-path"];
const requestedBrowser = normalizeBrowserName(parsed.browser ?? "edge");

if (!tokenPath || !storageStatePath || !requestedBrowser) {
  const browserHelp = [...SUPPORTED_BROWSERS].join(", ");
  console.error(
    `Missing or invalid args. Required: --token-path <path> --storage-state-path <path> [--browser <${browserHelp}>]`,
  );
  process.exit(2);
}

await fetchTokenWithPlaywrightNode(tokenPath, storageStatePath, requestedBrowser);

async function fetchTokenWithPlaywrightNode(
  tokenPath,
  storageStatePath,
  browserName,
) {
  console.log(
    `[playwright] Launching ${browserName} under Node.js (headed)...`,
  );
  const browser = await launchBrowser(browserName);
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
      await page.waitForTimeout(1_000);

      const sendButton = page
        .locator('button[title="Send"], [role="button"][title="Send"], [title="Send"]')
        .first();
      try {
        await sendButton.waitFor({ state: "visible", timeout: 10_000 });
        await sendButton.click({ timeout: 10_000 });
        console.log("[playwright] Send button clicked.");
      } catch {
        await page.keyboard.press("Enter");
        console.log("[playwright] Send button unavailable, submitted with Enter.");
      }
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

async function launchBrowser(browserName) {
  switch (browserName) {
    case "edge":
      try {
        return await chromium.launch({
          headless: false,
          channel: "msedge",
          args: CHROMIUM_LAUNCH_ARGS,
        });
      } catch {
        console.log(
          "[playwright] Edge channel unavailable, falling back to Chromium.",
        );
        return chromium.launch({
          headless: false,
          args: CHROMIUM_LAUNCH_ARGS,
        });
      }
    case "chrome":
      try {
        return await chromium.launch({
          headless: false,
          channel: "chrome",
          args: CHROMIUM_LAUNCH_ARGS,
        });
      } catch {
        console.log(
          "[playwright] Chrome channel unavailable, falling back to Chromium.",
        );
        return chromium.launch({
          headless: false,
          args: CHROMIUM_LAUNCH_ARGS,
        });
      }
    case "chromium":
      return chromium.launch({
        headless: false,
        args: CHROMIUM_LAUNCH_ARGS,
      });
    case "firefox":
      return firefox.launch({
        headless: false,
      });
    case "webkit":
      return webkit.launch({
        headless: false,
      });
    default:
      throw new Error(`Unsupported browser: ${String(browserName)}`);
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

function normalizeBrowserName(value) {
  if (typeof value !== "string") {
    return null;
  }
  const normalized = value.trim().toLowerCase();
  if (!normalized) {
    return null;
  }
  const canonical = normalized === "msedge" ? "edge" : normalized;
  return SUPPORTED_BROWSERS.has(canonical) ? canonical : null;
}
