# Plan: `token fetch` — Playwright-based Substrate token harvester

Add a new `token fetch` subcommand to the CLI that silently launches Edge, navigates to M365 Copilot, waits for login if needed, triggers a chat message to force a Substrate WebSocket connection, extracts `access_token` from that connection's URL, then persists both the token and the browser login state to disk.

## Steps

1. Already done.

2. **Create `src/cli/playwright-token.ts`** — new file containing a single exported async function `fetchTokenWithPlaywright(tokenPath: string, storageStatePath: string): Promise<void>`. Internals:
   - **Launch Edge (headed)** via `chromium.launch({ headless: false, channel: 'msedge' })`.
   - **Restore login state** — if `storageStatePath` exists on disk, pass `storageState: storageStatePath` to `browser.newContext(...)` so the user isn't re-prompted to log in.
   - **Navigate** to `https://m365.cloud.microsoft/chat/?auth=2`.
   - **Wait for login** — if the page URL contains an identity/login host (e.g. `login.microsoftonline.com`), `console.log` a hint telling the user to sign in, then `await page.waitForURL('**/chat**', { timeout: 300_000 })` (5 min).
   - **Register WebSocket listener before triggering message** — set up a `Promise<string>` that resolves with the token when `page.on('websocket', ws => ...)` fires for a socket whose `ws.url()` matches `substrate.office.com/m365Copilot/Chathub` and contains an `access_token` query param.
   - **Trigger the WebSocket** — `await page.locator('#m365-chat-editor-target-element').fill('Hi')` then `page.keyboard.press('Enter')`.
   - **Await the token** with a 30-second timeout. If it times out, throw a descriptive error telling the user to retry.
   - **Save browser login state** — `await context.storageState({ path: storageStatePath })`.
   - **Close browser** — `await browser.close()`.
   - **Decode expiry and save token** — call existing `tryGetJwtExpiry(rawToken)` (extracted/imported from `src/cli/index.ts`) and `saveToken(tokenPath, rawToken, expiry)`.

3. **Export shared helpers from `src/cli/index.ts`** — `getTokenPath`, `saveToken`, `tryGetJwtExpiry` are currently plain functions. Extract them (or the logic) so `playwright-token.ts` can import them without importing the whole CLI entry point (which runs `process.exit`). The cleanest approach is to move these helpers into a new **`src/cli/token-helpers.ts`** file and import from both `index.ts` and `playwright-token.ts`.

4. **Add `getBrowserStatePath()` to `token-helpers.ts`** — mirrors `getTokenPath()` but returns `…\YarpPilot\Cli\browser-state.json`.

5. **Wire up `token fetch` in `src/cli/index.ts`** — inside `runTokenCommand`, add:

   ```ts
   if (sub === "fetch") {
     const storageStatePath = await getBrowserStatePath();
     await fetchTokenWithPlaywright(tokenPath, storageStatePath);
     return 0;
   }
   ```

6. **Update `showUsage()`** in `src/cli/index.ts` — add the line `bun src/cli/index.ts token fetch`.

## Verification

- Run `bun src/cli/index.ts token fetch`. A headed Edge window should open M365 Copilot.
- Sign in (first run). After login, the CLI should type "Hi", capture the WebSocket token, save it, and print `Saved token. Expires: <date>`.
- Run `bun src/cli/index.ts token status` — state should be `valid`.
- Run `bun src/cli/index.ts token fetch` a second time — Edge should open already logged in and complete without a login prompt, using the saved `browser-state.json`.
- Run `bun src/cli/index.ts chat --message "Hello"` to confirm the harvested token is accepted by the proxy end-to-end.

## Decisions

- Subcommand: `token fetch` (as confirmed).
- Browser login state persisted to `browser-state.json` alongside `token.json` (as confirmed).
- Helper functions extracted to `src/cli/token-helpers.ts` to avoid circular/entry-point import issues.
- Edge launched via `channel: 'msedge'` (already installed on Windows) — no extra Playwright browser download required for end users.
