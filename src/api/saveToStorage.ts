import { API_BASE } from "./config";

export type StorageCategory = "Comments" | "Projects" | "Troubleshooting" | "Calendar";

export interface SaveInput {
  category: StorageCategory;
  uid: string;
  title: string;
  description: string;
  owner: string;
  // Optional timestamp override; when omitted current time is used.
  timestamp?: string | Date;
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

  // Post to the deployed Projects function (route: /api/projects)
  const url = `${API_BASE}/projects`;

  const body = {
    category: input.category,
    uid: input.uid,
    title: input.title,
    description: input.description,
    owner: input.owner,
    timestamp,
  };

  let res: Response;
  try {
    // Debug: surface the intent in the browser console
    // eslint-disable-next-line no-console
    console.debug('[saveToStorage] POST', url, body);
    res = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(body),
      signal,
      credentials: 'include' // include auth cookie when SWA protects /api
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
