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
 * Since GetItems API was removed, this now returns an empty list
 * or can be updated later when a new read endpoint exists.
 */
export async function getNotesForUid(uid: string): Promise<NoteEntity[]> {
  console.warn(
    `[getNotesForUid] Called for UID ${uid}, but GetItems endpoint no longer exists. Returning []`
  );
  return [];
}

/**
 * Save a new note for a UID.
 * This uses the current HttpTrigger1 function (POST).
 */
export async function saveNote(
  uid: string,
  description: string,
  owner: string = "Unknown"
): Promise<SaveResponse> {
  const url = `${API_BASE}/HttpTrigger1`;

  const body = {
    category: "Notes",
    uid,
    title: "General comment",
    description,
    owner,
  };

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
