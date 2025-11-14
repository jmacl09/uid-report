import { API_BASE } from "./config";

export type StorageCategory = "Comments" | "Notes" | "Projects" | "Troubleshooting" | "Calendar" | "Status" | "Suggestions";

export interface SaveInput {
  category: StorageCategory; // Logical grouping (e.g. Comments, Notes, Projects)
  uid: string;               // The UID being annotated
  title: string;             // Short title for the entry
  description: string;       // Longer free-form text
  owner: string;             // Person saving the record
  // Optional timestamp override; when omitted current time is used.
  timestamp?: string | Date;
  /**
   * Optional row key for updates. When provided the backend will upsert the existing
   * entity instead of creating a new row (used for comment edits).
   */
  rowKey?: string;
  /**
   * (Advanced) Override the target Azure Function route. Defaults to 'HttpTrigger1'
   * which is the new optical360v2-test function that persists to Table Storage.
   * Previous code used 'Projects'; keeping this extensibility lets older calls keep working
   * while new comment saves can explicitly hit the new function.
   */
  endpoint?: string;
  /** Optional extra properties to include when saving (e.g., Status, dcCode). */
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

// Reusable helper that posts to the Azure Function and returns response text.
export async function saveToStorage(input: SaveInput, options: SaveOptions = {}): Promise<string> {
  const { signal } = options;

  const timestamp = input.timestamp
    ? (input.timestamp instanceof Date ? input.timestamp.toISOString() : new Date(input.timestamp).toISOString())
    : new Date().toISOString();

  // Determine which backend function to hit.
  // If caller provides a full URL, use it as-is; otherwise build from API_BASE and endpoint.
  // Default endpoint is the new optical360v2-test function name 'HttpTrigger1'.
  const rawEndpoint = input.endpoint || 'HttpTrigger1';
  const isAbsolute = /^https?:\/\//i.test(rawEndpoint);
  const endpoint = isAbsolute ? rawEndpoint : `${API_BASE}/${rawEndpoint.replace(/^\/+/, '')}`;
  const url = endpoint;

  // Be liberl in what we send: include both lowerCamel and PascalCase keys
  // to maximize compatibility with any deployed Functions code paths.
  const body = {
  // canonical (lower camel case expected by new function)
    category: input.category,
    uid: input.uid,
    title: input.title,
    description: input.description,
    owner: input.owner,
    timestamp,
    ...(input.rowKey ? { rowKey: input.rowKey } : {}),
    // compatibility (some backends expect PascalCase or alternative names)
    Category: input.category,
    UID: input.uid,
    Title: input.title,
    Description: input.description,
    Owner: input.owner,
    Timestamp: timestamp,
    ...(input.rowKey ? { RowKey: input.rowKey } : {}),
    // spread any extra properties provided by the caller
    ...(input.extras || {}),
  } as Record<string, unknown>;

  let res: Response;
  try {
    // Debug: surface the intent in the browser console
    // eslint-disable-next-line no-console
    console.debug('[saveToStorage] POST', url, body);
    // Determine credentials policy: include for same-origin (SWA proxy), omit for cross-origin public Function URL.
    const isCrossOrigin = (() => {
      try {
        if (typeof window === 'undefined') return false;
        const target = new URL(url, window.location.href);
        return target.origin !== window.location.origin;
      } catch { return false; }
    })();

    res = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(body),
      signal,
      credentials: isCrossOrigin ? 'omit' : 'include',
    });
  } catch (networkErr: any) {
    // Network or CORS-level failure
    // eslint-disable-next-line no-console
    console.warn('[saveToStorage] Network/CORS error', networkErr);
    throw new SaveError(`Network error while saving data: ${networkErr?.message || networkErr}`, undefined);
  }

  const text = await res.text();

  if (!res.ok) {
    // Gracefully surface server and client errors to the caller
    const err = new SaveError(`Save failed with status ${res.status}`, res.status, text);
    // eslint-disable-next-line no-console
    console.warn('[saveToStorage] Server returned error', res.status, text);
    throw err;
  }

  return text;
}
