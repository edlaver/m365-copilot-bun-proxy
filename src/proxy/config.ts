import { promises as fs } from "node:fs";
import path from "node:path";
import { z } from "zod";
import { LogLevels, type JsonObject, type WrapperOptions } from "./types";
import { deepMerge, isJsonObject, parseEnvValue, setDeepValue } from "./utils";

const LogLevelSchema = z.preprocess(
  (value) => (typeof value === "string" ? value.trim().toLowerCase() : value),
  z.enum([
    LogLevels.Trace,
    LogLevels.Debug,
    LogLevels.Info,
    LogLevels.Warning,
    LogLevels.Error,
  ]),
);

const WrapperOptionsSchema = z.object({
  listenUrl: z.string().default("http://localhost:4000"),
  debugPath: z.string().nullable().default("./Logs"),
  logLevel: LogLevelSchema.default(LogLevels.Info),
  ignoreIncomingAuthorizationHeader: z.boolean().default(true),
  transport: z.string().default("graph"),
  graphBaseUrl: z.string().default("https://graph.microsoft.com"),
  createConversationPath: z.string().default("/beta/copilot/conversations"),
  chatPathTemplate: z
    .string()
    .default("/beta/copilot/conversations/{conversationId}/chat"),
  chatOverStreamPathTemplate: z
    .string()
    .default("/beta/copilot/conversations/{conversationId}/chatOverStream"),
  substrate: z
    .object({
      hubPath: z
        .string()
        .default("wss://substrate.office.com/m365Copilot/Chathub"),
      source: z.string().default("officeweb"),
      quoteSourceInQuery: z.boolean().default(true),
      scenario: z.string().default("OfficeWebIncludedCopilot"),
      origin: z.string().default("https://m365.cloud.microsoft"),
      product: z.string().nullable().default("Office"),
      agentHost: z.string().nullable().default("Bizchat.FullScreen"),
      licenseType: z.string().nullable().default("Starter"),
      agent: z.string().nullable().default("web"),
      variants: z.string().nullable().default(null),
      clientPlatform: z.string().default("web"),
      productThreadType: z.string().default("Office"),
      invocationTimeoutSeconds: z.number().int().default(120),
      keepAliveSeconds: z.number().int().default(15),
      optionsSets: z
        .array(z.string())
        .default([
          "enterprise_flux_web",
          "enterprise_flux_work",
          "enable_request_response_interstitials",
          "enterprise_flux_image_v1",
          "enterprise_toolbox_with_skdsstore",
          "enterprise_toolbox_with_skdsstore_search_message_extensions",
          "enable_ME_auth_interstitial",
          "skdsstorethirdparty",
          "enable_confirmation_interstitial",
          "enable_plugin_auth_interstitial",
          "enable_response_action_processing",
          "enterprise_flux_work_gptv",
          "enterprise_flux_work_code_interpreter",
          "enable_batch_token_processing",
        ]),
      allowedMessageTypes: z
        .array(z.string())
        .default([
          "Chat",
          "Suggestion",
          "InternalSearchQuery",
          "InternalSearchResult",
          "Disengaged",
          "InternalLoaderMessage",
          "RenderCardRequest",
          "AdsQuery",
          "SemanticSerp",
          "GenerateContentQuery",
          "SearchQuery",
          "ConfirmationCard",
          "AuthError",
          "DeveloperLogs",
        ]),
      invocationTarget: z.string().default("chat"),
      invocationType: z.number().int().default(4),
      locale: z.string().default("en-US"),
      experienceType: z.string().default("Default"),
      entityAnnotationTypes: z
        .array(z.string())
        .default(["People", "File", "Event", "Email", "TeamsMessage"]),
    })
    .default({}),
  defaultModel: z.string().default("m365-copilot"),
  defaultTimeZone: z.string().default("America/New_York"),
  conversationTtlMinutes: z.number().int().default(180),
  maxAdditionalContextMessages: z.number().int().default(16),
  includeConversationIdInResponseBody: z.boolean().default(true),
});

export async function loadWrapperOptions(cwd: string): Promise<WrapperOptions> {
  const rootConfig: JsonObject = {};
  const baseConfig = await readJsonFile(path.join(cwd, "config.json"));
  deepMerge(rootConfig, baseConfig ?? {});

  const env = process.env.NODE_ENV;
  if (env?.trim()) {
    const envConfig = await readJsonFile(path.join(cwd, `config.${env}.json`));
    deepMerge(rootConfig, envConfig ?? {});
  }

  applyConfigEnvOverrides(rootConfig, process.env);
  return normalizeWrapperOptions(rootConfig);
}

function normalizeWrapperOptions(wrapper: JsonObject): WrapperOptions {
  const normalized: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(wrapper)) {
    normalized[key] = value;
  }

  for (const [key, value] of Object.entries(normalized)) {
    if (typeof value !== "string") {
      continue;
    }
    if (/^-?\d+(\.\d+)?$/.test(value.trim())) {
      normalized[key] = Number.parseFloat(value);
      continue;
    }
    const lowered = value.trim().toLowerCase();
    if (lowered === "true" || lowered === "false") {
      normalized[key] = lowered === "true";
    }
  }

  return WrapperOptionsSchema.parse(normalized) as WrapperOptions;
}

function applyConfigEnvOverrides(
  wrapper: JsonObject,
  env: NodeJS.ProcessEnv,
): void {
  for (const [key, value] of Object.entries(env)) {
    if (!value) {
      continue;
    }
    if (!key.toUpperCase().startsWith("CONFIG__")) {
      continue;
    }
    const pathParts = key
      .split("__")
      .slice(1)
      .filter((part) => part.trim().length > 0);
    if (pathParts.length === 0) {
      continue;
    }
    setDeepValue(wrapper, pathParts, parseEnvValue(value));
  }
}

async function readJsonFile(filePath: string): Promise<JsonObject | null> {
  try {
    const content = await fs.readFile(filePath, "utf8");
    const parsed = JSON.parse(content) as unknown;
    return isJsonObject(parsed) ? parsed : null;
  } catch {
    return null;
  }
}
