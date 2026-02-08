# m365-copilot-bun-proxy

TypeScript/Bun port of the original YarpPilot .NET proxy + CLI.

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
bun run cli -- token set --token "<jwt>"
```

In chat mode, the CLI supports these slash commands:

- `/status` (token + connection status)
- `/token` (paste a new token)
- `/cleartoken` (clear cached token)
- `/exit` (quit)
