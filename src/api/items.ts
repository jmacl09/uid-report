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
 * Fetch notes for a given UID from the HttpTrigger1 function.
 * Uses category=Notes to align with saveToStorage usage in UIDLookup.
 */
export async function getNotesForUid(uid: string, endpoint?: string): Promise<NoteEntity[]> {
  // Prefer the deployed Function URL by default to avoid proxy/env issues
  const defaultAbsolute = 'https://optical360v2-ffa9ewbfafdvfyd8.westeurope-01.azurewebsites.net/api/HttpTrigger1';
  const rawEndpoint = endpoint || defaultAbsolute;
  const isAbsolute = /^https?:\/\//i.test(rawEndpoint);
  const base = isAbsolute ? rawEndpoint.replace(/\/?$/,'') : `${API_BASE}/${rawEndpoint.replace(/^\/+/, '')}`;
  const url = `${base}?uid=${encodeURIComponent(uid)}&category=${encodeURIComponent('Notes')}`;

  const res = await fetch(url, { method: 'GET' });
  const text = await res.text();
  if (!res.ok) throw new Error(`getNotesForUid failed ${res.status}: ${text}`);
  try {
    const data = JSON.parse(text);
    if (data?.items && Array.isArray(data.items)) return data.items as NoteEntity[];
    return [];
  } catch {
    return [];
  }
}

/**
 * Fetch comments for a given UID from the HttpTrigger1 function.
 * If endpoint is an absolute URL, it's used directly; otherwise we build from API_BASE.
 */
export async function getCommentsForUid(uid: string, endpoint?: string): Promise<NoteEntity[]> {
  const rawEndpoint = endpoint || 'HttpTrigger1';
  const isAbsolute = /^https?:\/\//i.test(rawEndpoint);
  const base = isAbsolute ? rawEndpoint.replace(/\/?$/,'') : `${API_BASE}/${rawEndpoint.replace(/^\/+/, '')}`;
  const url = `${base}?uid=${encodeURIComponent(uid)}&category=${encodeURIComponent('Comments')}`;

  const res = await fetch(url, { method: 'GET' });
  const text = await res.text();
  if (!res.ok) throw new Error(`getCommentsForUid failed ${res.status}: ${text}`);
  try {
    const data = JSON.parse(text);
    if (data?.items && Array.isArray(data.items)) return data.items as NoteEntity[];
    return [];
  } catch {
    return [];
  }
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
