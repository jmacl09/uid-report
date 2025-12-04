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

/* ------------------------------------------------------------------
   Helpers
------------------------------------------------------------------- */
function buildUrl(endpoint: string | undefined, params: Record<string, string>) {
  const ep = endpoint || "HttpTrigger1";
  const base = ep.startsWith("http") ? ep : `${API_BASE}/${ep}`;

  const query = new URLSearchParams(params).toString();
  return `${base}?${query}`;
}

async function safeFetchList(url: string): Promise<NoteEntity[]> {
  try {
    const res = await fetch(url);
    if (!res.ok) return [];

    const json = await res.json();
    if (Array.isArray(json)) return json;
    if (json?.items) return json.items;
    return [];
  } catch {
    return [];
  }
}

/* ------------------------------------------------------------------
   GET FUNCTIONS
------------------------------------------------------------------- */

export async function getNotesForUid(uid: string, endpoint?: string) {
  return safeFetchList(buildUrl(endpoint, { uid, category: "Notes" }));
}

export async function getCommentsForUid(uid: string, endpoint?: string) {
  return safeFetchList(buildUrl(endpoint, { uid, category: "Comments" }));
}

export async function getProjectsForUid(uid: string, endpoint?: string) {
  return safeFetchList(buildUrl(endpoint, { uid, category: "Projects" }));
}

export async function getTroubleshootingForUid(uid: string, endpoint?: string) {
  return safeFetchList(buildUrl(endpoint, { uid, category: "Troubleshooting" }));
}

export async function getStatusForUid(uid: string, endpoint?: string) {
  return safeFetchList(buildUrl(endpoint, { uid, category: "Status" }));
}

export async function getCalendarEntries(uid: string, endpoint?: string) {
  return safeFetchList(buildUrl(endpoint, { uid, category: "Calendar" }));
}

export async function getAllSuggestions(endpoint?: string) {
  return safeFetchList(buildUrl(endpoint, { category: "Suggestions" }));
}

export async function getAllProjects(endpoint?: string) {
  return safeFetchList(buildUrl(endpoint, { category: "Projects" }));
}

/* ------------------------------------------------------------------
   SAVE NOTE
------------------------------------------------------------------- */

export async function saveNote(uid: string, description: string, owner = "Unknown"): Promise<SaveResponse> {
  const body = {
    category: "Notes",
    uid,
    title: "General note",
    description,
    owner,
  };

  const res = await fetch(`${API_BASE}/HttpTrigger1`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });

  const text = await res.text();
  return res.ok ? JSON.parse(text) : { ok: false, message: text };
}

/* ------------------------------------------------------------------
   DELETE NOTE
------------------------------------------------------------------- */

export async function deleteNote(partitionKey: string, rowKey: string, category = "Notes") {
  const body = { partitionKey, rowKey, category };

  const res = await fetch(`${API_BASE}/HttpTrigger1`, {
    method: "DELETE",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });

  if (!res.ok) {
    throw new Error(`Delete failed: ${await res.text()}`);
  }
}
