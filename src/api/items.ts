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

// -------------------------------
// Helpers
// -------------------------------

function buildUrl(rawEndpoint: string | undefined, params: Record<string, string>) {
  const ep = rawEndpoint || "HttpTrigger1";
  const isAbs = /^https?:\/\//i.test(ep);
  const base = isAbs ? ep.replace(/\/?$/, "") : `${API_BASE}/${ep.replace(/^\/+/, "")}`;

  const query = Object.entries(params)
    .map(([k, v]) => `${k}=${encodeURIComponent(v)}`)
    .join("&");

  return `${base}?${query}`;
}

async function safeFetchList(url: string): Promise<NoteEntity[]> {
  try {
    const res = await fetch(url, { method: "GET" });
    const text = await res.text();

    if (!res.ok) return [];

    try {
      const data = JSON.parse(text);
      if (Array.isArray(data)) return data;
      if (data?.items && Array.isArray(data.items)) return data.items;
      if (data?.value && Array.isArray(data.value)) return data.value;
      if (data?.entity && Array.isArray(data.entity)) return data.entity;
      return [];
    } catch {
      return [];
    }
  } catch {
    return [];
  }
}

// -------------------------------
// FETCH FUNCTIONS
// -------------------------------

export async function getNotesForUid(uid: string, endpoint?: string): Promise<NoteEntity[]> {
  const url = buildUrl(endpoint, { uid, category: "Notes" });
  return safeFetchList(url);
}

export async function getCommentsForUid(uid: string, endpoint?: string): Promise<NoteEntity[]> {
  const url = buildUrl(endpoint, { uid, category: "Comments" });
  return safeFetchList(url);
}

export async function getCalendarEntries(uid: string, endpoint?: string): Promise<NoteEntity[]> {
  const url = buildUrl(endpoint, { uid, category: "Calendar" });
  return safeFetchList(url);
}

export async function getTroubleshootingForUid(uid: string, endpoint?: string): Promise<NoteEntity[]> {
  const url = buildUrl(endpoint, {
    uid,
    category: "troubleshooting".toLowerCase(),
  });
  return safeFetchList(url);
}

export async function getProjectsForUid(uid: string, endpoint?: string): Promise<NoteEntity[]> {
  const url = buildUrl(endpoint, { uid, category: "Projects" });
  return safeFetchList(url);
}

// FETCH ALL PROJECTS
export async function getAllProjects(endpoint?: string): Promise<NoteEntity[]> {
  const url = buildUrl(endpoint, { category: "Projects" });
  return safeFetchList(url);
}

// FETCH ALL SUGGESTIONS
export async function getAllSuggestions(endpoint?: string): Promise<NoteEntity[]> {
  const url = buildUrl(endpoint, {
    category: "suggestions".toLowerCase(),
  });
  return safeFetchList(url);
}

export async function getStatusForUid(uid: string, endpoint?: string): Promise<NoteEntity[]> {
  const url = buildUrl(endpoint, { uid, category: "Status" });
  return safeFetchList(url);
}

// -------------------------------
// SAVE NOTE
// -------------------------------

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
  };

  const res = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
    credentials: "include",
  });

  const text = await res.text();
  if (!res.ok) throw new Error(`saveNote failed ${res.status}: ${text}`);

  try {
    return JSON.parse(text) as SaveResponse;
  } catch {
    return { ok: true, message: text };
  }
}

// -------------------------------
// DELETE
// -------------------------------

export async function deleteNote(
  partitionKey: string,
  rowKey: string,
  endpoint?: string,
): Promise<void> {
  const ep = endpoint || "HttpTrigger1";
  const isAbs = /^https?:\/\//i.test(ep);
  const base = isAbs ? ep.replace(/\/?$/, "") : `${API_BASE}/${ep.replace(/^\/+/, "")}`;

  const url = base;

  const body: Record<string, string> = { partitionKey, rowKey };

  // If PK starts with UID_xxx extract UID for troubleshooting
  const uidMatch = /^UID_(.+)$/i.exec(partitionKey || "");
  if (uidMatch && uidMatch[1]) body.uid = uidMatch[1];

  await fetch(url, {
    method: "DELETE",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
    credentials: isAbs ? "omit" : "include",
  }).then(async (res) => {
    if (!res.ok) {
      const text = await res.text();
      throw new Error(`deleteNote failed ${res.status}: ${text}`);
    }
  });
}
