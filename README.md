# m365-copilot-bun-proxy

TypeScript/Bun port of the original M365 Copilot Bun Proxy .NET proxy + CLI.

## Stack

- Bun runtime
- Hono for HTTP routing / reverse-proxy behavior
- Zod for configuration validation
- OpenTUI for interactive CLI chat UI

## Install

```bash
bun install
```

## Run proxy

```bash
bun run start:proxy
```

To enable debug markdown logs (requires `debugPath` in config):

```bash
bun run start:proxy -- --debug
```

You can also pass an explicit value:

```bash
bun run start:proxy -- --debug=false
```

Default listen URL is `http://localhost:4000`.

Configuration is loaded from `config.json` (and `config.{env}.json` when `NODE_ENV` is set).

Substrate settings are grouped under the `substrate` object in config (for example `substrate.hubPath`).

You can override config values via env vars with the `CONFIG__` prefix, for example:

```bash
CONFIG__listenUrl=http://localhost:4010 bun run start:proxy
```

To override nested values, use double underscores for each path segment, for example:

```bash
CONFIG__substrate__hubPath=wss://substrate.office.com/m365Copilot/Chathub bun run start:proxy
```

## API endpoints

- `POST /v1/chat/completions`
- `POST /openai/v1/chat/completions`
- `POST /v1/responses`
- `POST /openai/v1/responses`
- `GET /v1/responses`
- `GET /openai/v1/responses`
- `GET /v1/responses/{response_id}`
- `GET /openai/v1/responses/{response_id}`
- `DELETE /v1/responses/{response_id}`
- `DELETE /openai/v1/responses/{response_id}`

## Responses API usage

Create response:

```bash
curl -s http://localhost:4000/v1/responses \
  -H "Content-Type: application/json" \
  -H "x-m365-transport: substrate" \
  -d '{
    "model": "m365-copilot",
    "input": "Write a TypeScript function that validates UUIDs."
  }'
```

Continue a conversation using `previous_response_id`:

```bash
curl -s http://localhost:4000/v1/responses \
  -H "Content-Type: application/json" \
  -H "x-m365-transport: substrate" \
  -d '{
    "model": "m365-copilot",
    "previous_response_id": "resp_abc123",
    "input": "Now add tests."
  }'
```

Streaming (`stream: true`) emits SSE events:

- `response.created`
- `response.in_progress`
- `response.output_item.added`
- `response.output_text.delta`
- `response.output_text.done`
- `response.output_item.done`
- `response.completed`
- `error` (SSE error event on stream failure)

If `Authorization` is missing or empty on chat/responses requests, the proxy now attempts to auto-acquire a token via Playwright and then retries with that token. You can still pass an explicit bearer token when needed.

## Build executable

```bash
bun run build
```

This produces a single-file executable in `dist/` and copies `config.json` alongside it.

## Run CLI

```bash
bun run cli -- help
bun run cli -- status
bun run cli -- chat
bun run cli -- chat --api responses
bun run cli -- token set --token "<jwt>"
```

By default, CLI chat requests do not send an Authorization header. The proxy handles token acquisition when needed. Use `--token` or `YARPILOT_TOKEN` only when you want to force a specific token from the CLI.

In chat mode, the CLI supports these slash commands:

- `/status` (token + connection status)
- `/api` (show current API mode)
- `/api completions` or `/api responses` (toggle endpoint)
- `/token` (paste a new token)
- `/cleartoken` (clear cached token)
- `/exit` (quit)
