import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";
import readline from "node:readline/promises";

export type TokenState = {
  token: string;
  expiresAtUtc: string;
};

export type TokenSummary = {
  state: "not set" | "valid" | "expired";
  expiry: string;
};

export async function getTokenPath(): Promise<string> {
  const localAppData =
    process.env.LOCALAPPDATA ?? path.join(os.homedir(), ".local", "share");
  const directory = path.join(localAppData, "M365 Copilot Bun Proxy", "Cli");
  await fs.mkdir(directory, { recursive: true });
  return path.join(directory, "token.json");
}

export async function getBrowserStatePath(): Promise<string> {
  const localAppData =
    process.env.LOCALAPPDATA ?? path.join(os.homedir(), ".local", "share");
  const directory = path.join(localAppData, "M365 Copilot Bun Proxy", "Cli");
  await fs.mkdir(directory, { recursive: true });
  return path.join(directory, "browser-state.json");
}

export async function loadToken(filePath: string): Promise<TokenState | null> {
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

export async function saveToken(
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

export async function deleteToken(filePath: string): Promise<boolean> {
  try {
    await fs.unlink(filePath);
    return true;
  } catch {
    return false;
  }
}

export function buildTokenSummary(tokenState: TokenState | null): TokenSummary {
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

export async function ensureValidToken(
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

export function parseTokenOrThrow(rawToken: string): {
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

export async function promptForTokenInteractive(): Promise<{
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

export function normalizeToken(raw: string): string {
  const trimmed = raw.trim();
  return trimmed.toLowerCase().startsWith("bearer ")
    ? trimmed.slice("Bearer ".length).trim()
    : trimmed;
}

export function tryGetJwtExpiry(token: string): Date | null {
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
