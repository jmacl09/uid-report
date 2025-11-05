import { API_BASE } from "./config";

export type StorageCategory = "Comments" | "Notes" | "Projects" | "Troubleshooting" | "Calendar";

export interface SaveInput {
  category: StorageCategory;
  uid: string;
  title: string;
  description: string;
  owner: string;
  timestamp?: string | Date; // Optional timestamp override
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

// === MAIN FUNCTION ===
export async function saveToStorage(input: SaveInput, options: SaveOptions = {}): Promise<string> {
  const { signal } = options;

  const timestamp = input.timestamp
    ? (input.timestamp instanceof Date ? input.timestamp.toISOString() : new Date(input.timestamp).toISOString())
    : new Date().toISOString();

  const url = `${API_BASE}/HttpTrigger1`;

  const body = {
    category: input.category,
    uid: input.uid,
    title: input.title,
    description: input.description,
    owner: input.owner,
    timestamp,
  };

  try {
    console.debug("[saveToStorage] POST", url, body);

    const res = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
      signal,
    });

    const text = await res.text();
    if (!res.ok) {
      throw new SaveError(`Save failed with status ${res.status}`, res.status, text);
    }

    return text;
  } catch (networkErr: any) {
    console.warn("[saveToStorage] Network/CORS error", networkErr);
    throw new SaveError(`Network error while saving data: ${networkErr?.message || networkErr}`, undefined);
  }
}
