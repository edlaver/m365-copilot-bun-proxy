import { fetchTokenWithPlaywright } from "../cli/playwright-token";
import {
  getBrowserStatePath,
  getTokenPath,
  loadToken,
  type TokenState,
} from "../cli/token-helpers";
import { normalizeBearerToken } from "./utils";

const TOKEN_EXPIRY_SKEW_MS = 60_000;

export class ProxyTokenProvider {
  private readonly tokenPathPromise: Promise<string>;
  private readonly browserStatePathPromise: Promise<string>;
  private inFlightAcquirePromise: Promise<string | null> | null = null;

  constructor() {
    this.tokenPathPromise = getTokenPath();
    this.browserStatePathPromise = getBrowserStatePath();
  }

  async resolveAuthorizationHeader(
    rawAuthorizationHeader: string | null | undefined,
  ): Promise<string | null> {
    const providedHeader = normalizeBearerToken(rawAuthorizationHeader);
    if (providedHeader) {
      return providedHeader;
    }

    const cachedHeader = await this.tryGetCachedAuthorizationHeader();
    if (cachedHeader) {
      return cachedHeader;
    }

    return this.acquireAuthorizationHeader();
  }

  private async tryGetCachedAuthorizationHeader(): Promise<string | null> {
    const tokenPath = await this.tokenPathPromise;
    const tokenState = await loadToken(tokenPath);
    if (!isTokenStateValid(tokenState)) {
      return null;
    }
    return `Bearer ${tokenState.token}`;
  }

  private async acquireAuthorizationHeader(): Promise<string | null> {
    const pendingAcquire =
      this.inFlightAcquirePromise ?? this.acquireFreshAuthorizationHeader();
    if (!this.inFlightAcquirePromise) {
      this.inFlightAcquirePromise = pendingAcquire;
    }
    try {
      return await pendingAcquire;
    } finally {
      if (this.inFlightAcquirePromise === pendingAcquire) {
        this.inFlightAcquirePromise = null;
      }
    }
  }

  private async acquireFreshAuthorizationHeader(): Promise<string | null> {
    const [tokenPath, browserStatePath] = await Promise.all([
      this.tokenPathPromise,
      this.browserStatePathPromise,
    ]);
    try {
      await fetchTokenWithPlaywright(tokenPath, browserStatePath, {
        quiet: true,
      });
    } catch {
      return null;
    }

    const fetched = await loadToken(tokenPath);
    return isTokenStateValid(fetched) ? `Bearer ${fetched.token}` : null;
  }
}

function isTokenStateValid(tokenState: TokenState | null): tokenState is TokenState {
  if (!tokenState?.token?.trim()) {
    return false;
  }
  const expiresAtUtc = new Date(tokenState.expiresAtUtc);
  return expiresAtUtc.getTime() > Date.now() + TOKEN_EXPIRY_SKEW_MS;
}
