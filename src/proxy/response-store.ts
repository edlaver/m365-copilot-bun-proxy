import type {
  JsonObject,
  StoredOpenAiResponseRecord,
  WrapperOptions,
} from "./types";
import { cloneJsonValue, nowUnix } from "./utils";

type ConversationLinkEntry = {
  conversationId: string;
  expiresAtUtc: number;
};

export class ResponseStore {
  private readonly entries = new Map<string, StoredOpenAiResponseRecord>();
  private readonly conversationLinks = new Map<string, ConversationLinkEntry>();

  constructor(private readonly options: WrapperOptions) {}

  set(
    responseId: string,
    response: JsonObject,
    conversationId: string | null,
  ): void {
    if (!responseId.trim()) {
      return;
    }
    this.purgeExpired();
    const record: StoredOpenAiResponseRecord = {
      responseId,
      createdAtUnix: readCreatedAt(response),
      response: cloneJsonValue(response),
      conversationId: conversationId?.trim() ? conversationId : null,
      expiresAtUtc: this.resolveExpiryMs(),
    };
    this.entries.set(responseId, record);

    if (conversationId?.trim()) {
      this.conversationLinks.set(responseId, {
        conversationId: conversationId.trim(),
        expiresAtUtc: record.expiresAtUtc,
      });
    }
  }

  tryGet(responseId: string): JsonObject | null {
    this.purgeExpired();
    const entry = this.entries.get(responseId);
    if (!entry) {
      return null;
    }
    if (entry.expiresAtUtc <= Date.now()) {
      this.entries.delete(responseId);
      this.conversationLinks.delete(responseId);
      return null;
    }
    return cloneJsonValue(entry.response);
  }

  tryDelete(responseId: string): boolean {
    this.purgeExpired();
    const deletedEntry = this.entries.delete(responseId);
    const deletedLink = this.conversationLinks.delete(responseId);
    return deletedEntry || deletedLink;
  }

  list(limit: number): {
    data: JsonObject[];
    hasMore: boolean;
    firstId: string | null;
    lastId: string | null;
  } {
    this.purgeExpired();
    const normalizedLimit = normalizeLimit(limit);
    const sorted = [...this.entries.values()].sort(
      (a, b) => b.createdAtUnix - a.createdAtUnix,
    );
    const total = sorted.length;
    const selected = sorted.slice(0, normalizedLimit);
    return {
      data: selected.map((entry) => cloneJsonValue(entry.response)),
      hasMore: total > selected.length,
      firstId: selected.length > 0 ? selected[0].responseId : null,
      lastId:
        selected.length > 0 ? selected[selected.length - 1].responseId : null,
    };
  }

  setConversationLink(responseId: string, conversationId: string): void {
    if (!responseId.trim() || !conversationId.trim()) {
      return;
    }
    this.purgeExpired();
    this.conversationLinks.set(responseId, {
      conversationId: conversationId.trim(),
      expiresAtUtc: this.resolveExpiryMs(),
    });
  }

  tryGetConversationLink(responseId: string): string | null {
    this.purgeExpired();
    const entry = this.conversationLinks.get(responseId);
    if (!entry) {
      return null;
    }
    if (entry.expiresAtUtc <= Date.now()) {
      this.conversationLinks.delete(responseId);
      return null;
    }
    return entry.conversationId;
  }

  private resolveExpiryMs(): number {
    const ttlMinutes = this.options.conversationTtlMinutes;
    if (ttlMinutes <= 0) {
      return Number.MAX_SAFE_INTEGER;
    }
    return Date.now() + ttlMinutes * 60_000;
  }

  private purgeExpired(): void {
    if (this.entries.size > 0) {
      const now = Date.now();
      for (const [id, entry] of this.entries.entries()) {
        if (entry.expiresAtUtc <= now) {
          this.entries.delete(id);
          this.conversationLinks.delete(id);
        }
      }
    }

    if (this.conversationLinks.size > 0) {
      const now = Date.now();
      for (const [id, entry] of this.conversationLinks.entries()) {
        if (entry.expiresAtUtc <= now) {
          this.conversationLinks.delete(id);
        }
      }
    }
  }
}

function normalizeLimit(rawLimit: number): number {
  if (!Number.isFinite(rawLimit)) {
    return 20;
  }
  const rounded = Math.trunc(rawLimit);
  if (rounded <= 0) {
    return 20;
  }
  return rounded > 100 ? 100 : rounded;
}

function readCreatedAt(response: JsonObject): number {
  const value = response.created_at;
  if (typeof value === "number" && Number.isFinite(value)) {
    return Math.trunc(value);
  }
  if (typeof value === "string") {
    const parsed = Number.parseInt(value, 10);
    if (Number.isFinite(parsed)) {
      return parsed;
    }
  }
  return nowUnix();
}
