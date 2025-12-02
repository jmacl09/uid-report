import { API_BASE } from "./config";

export type StorageCategory =
  | "Comments"
  | "Notes"
  | "Projects"
  | "Troubleshooting"
  | "Calendar"
  | "Status"
  | "Suggestions";

export interface SaveInput {
  category: StorageCategory;
  uid: string;
  title?: string;
  description: string;
  owner: string;
  timestamp?: string | Date;
  rowKey?: string;
  endpoint?: string;
  extras?: Record<string, unknown>;
}

export interface SaveOptions {
  signal?: AbortSignal;
}

export interface SaveResult {
  ok: boolean;
  status: number;
  text: string;
}

export class SaveError extends Error {
  status?: number;
  body?: string;
  constructor(message: string, status?: number, body?: string) {
    super(message);
    this.name = "SaveError";
    this.status = status;
    this.body = body;
  }
}

export async function saveToStorage(
  input: SaveInput,
  options: SaveOptions = {}
): Promise<string> {
  const { signal } = options;

  // ----------------------------------------------------------
  // ⭐ FIX #1 — Map UI category "Comments" → real table "Troubleshooting"
  // ----------------------------------------------------------
  let resolvedCategory = input.category;
  if (resolvedCategory === "Comments") {
    resolvedCategory = "Troubleshooting";
  }

  // ----------------------------------------------------------
  // ⭐ FIX #2 — Ensure Troubleshooting always has a title
  // ----------------------------------------------------------
  let resolvedTitle = input.title;
  if (
    resolvedCategory === "Troubleshooting" &&
    (!resolvedTitle || resolvedTitle.trim() === "")
  ) {
    resolvedTitle = "Troubleshooting Entry";
  }

  // ----------------------------------------------------------
  // ⭐ FIX #3 — Ensure ANY category has a fallback title
  // ----------------------------------------------------------
  if (!resolvedTitle || resolvedTitle.trim() === "") {
    resolvedTitle = "General Entry";
  }

  // ----------------------------------------------------------
  // Timestamp normalization
  // ----------------------------------------------------------
  const timestamp = input.timestamp
    ? input.timestamp instanceof Date
      ? input.timestamp.toISOString()
      : new Date(input.timestamp).toISOString()
    : new Date().toISOString();

  // ----------------------------------------------------------
  // Select backend endpoint (absolute URL or proxied)
  // ----------------------------------------------------------
  const rawEndpoint = input.endpoint || "HttpTrigger1";
  const isAbsolute = /^https?:\/\//i.test(rawEndpoint);
  const url = isAbsolute
    ? rawEndpoint
    : `${API_BASE}/${rawEndpoint.replace(/^\/+/, "")}`;

  // ----------------------------------------------------------
  // Build payload for Function App
  // Includes camelCase + PascalCase for compatibility
  // ----------------------------------------------------------
  const body = {
    category: resolvedCategory,
    uid: input.uid,
    title: resolvedTitle,
    description: input.description,
    owner: input.owner,
    timestamp,
    ...(input.rowKey ? { rowKey: input.rowKey } : {}),

    // Old function paths expect PascalCase
    Category: resolvedCategory,
    UID: input.uid,
    Title: resolvedTitle,
    Description: input.description,
    Owner: input.owner,
    Timestamp: timestamp,
    ...(input.rowKey ? { RowKey: input.rowKey } : {}),

    ...(input.extras || {}),
  } as Record<string, unknown>;

  let res: Response;

  try {
    console.debug("[saveToStorage] POST", url, body);

    // Avoid credentials for cross-origin requests
    const isCrossOrigin = (() => {
      try {
        if (typeof window === "undefined") return false;
        const target = new URL(url, window.location.href);
        return target.origin !== window.location.origin;
      } catch {
        return false;
      }
    })();

    res = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      credentials: isCrossOrigin ? "omit" : "include",
      body: JSON.stringify(body),
      signal,
    });
  } catch (networkErr: any) {
    console.warn("[saveToStorage] Network/CORS error", networkErr);
    throw new SaveError(
      `Network error while saving data: ${
        networkErr?.message || networkErr
      }`,
      undefined
    );
  }

  const text = await res.text();

  if (!res.ok) {
    const err = new SaveError(`Save failed with status ${res.status}`, res.status, text);
    console.warn("[saveToStorage] Server returned error", res.status, text);
    throw err;
  }

  return text;
}
