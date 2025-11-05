import { API_BASE } from "./config";

export type NoteEntity = {
  partitionKey: string;
  rowKey: string;
  category?: string;
  title?: string;
  description?: string;
  owner?: string;
  savedAt?: string;
  [key: string]: any;
};

export type SaveResponse = {
  ok: boolean;
  message?: string;
  entity?: NoteEntity;
};

/**
 * Fetch notes for a given UID.
 * Read endpoint is not implemented yet; return empty array for now.
 */
export async function getNotesForUid(_uid: string): Promise<NoteEntity[]> {
  return [];
}

/**
 * Save a new note for a UID.
 * This uses the Azure Function routed as /api/projects (POST).
 */
export async function saveNote(
  uid: string,
  description: string,
  owner: string = "Unknown"
): Promise<SaveResponse> {
  const url = `${API_BASE}/Projects`;

  const body = {
    category: "Notes",
    uid,
    title: "General comment",
    description,
    owner,
    // compatibility keys for backends expecting PascalCase
    Category: "Notes",
    UID: uid,
    Title: "General comment",
    Description: description,
    Owner: owner,
  } as Record<string, unknown>;

  const res = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
    credentials: "include",
  });

  const text = await res.text();
  if (!res.ok) {
    throw new Error(`saveNote failed ${res.status}: ${text}`);
  }

  try {
    return JSON.parse(text) as SaveResponse;
  } catch {
    return { ok: true, message: text };
  }
}

/**
 * Delete a note (disabled until DeleteItem function is re-added).
 */
export async function deleteNote(
  partitionKey: string,
  rowKey: string
): Promise<void> {
  console.warn(
    `[deleteNote] Called for ${partitionKey}/${rowKey}, but DeleteItem endpoint is not implemented.`
  );
  return;
}
