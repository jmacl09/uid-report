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
      if (data?.Entity && Array.isArray(data.Entity)) return data.Entity;
      return [];
    } catch {
      return [];
    }
  } catch {
    return [];
  }
}

// -------------------------------
// Fetch Notes
// -------------------------------

export async function getNotesForUid(uid: string, endpoint?: string): Promise<NoteEntity[]> {
  const url = buildUrl(endpoint, {
    uid,
    category: "Notes",
  });
  return safeFetchList(url);
}

// -------------------------------
// Fetch Comments
// -------------------------------

export async function getCommentsForUid(uid: string, endpoint?: string): Promise<NoteEntity[]> {
  const url = buildUrl(endpoint, {
    uid,
    category: "Comments",
  });
  return safeFetchList(url);
}

// -------------------------------
// Fetch Calendar Entries
// -------------------------------

export async function getCalendarEntries(uid: string, endpoint?: string): Promise<NoteEntity[]> {
  const url = buildUrl(endpoint, {
    uid,
    category: "Calendar",
  });
  return safeFetchList(url);
}

// -------------------------------
// Fetch Troubleshooting
// -------------------------------

export async function getTroubleshootingForUid(uid: string, endpoint?: string): Promise<NoteEntity[]> {
  const url = buildUrl(endpoint, {
    uid,
    category: "Troubleshooting",
    tableName: "Troubleshooting",
    TableName: "Troubleshooting",
    targetTable: "Troubleshooting",
  });
  return safeFetchList(url);
}

// -------------------------------
// Fetch Projects (UID-based)
// -------------------------------

export async function getProjectsForUid(uid: string, endpoint?: string): Promise<NoteEntity[]> {
  const url = buildUrl(endpoint, {
    uid,
    category: "Projects",
  });
  return safeFetchList(url);
}

// -------------------------------
// Fetch ALL Projects
// -------------------------------

export async function getAllProjects(endpoint?: string): Promise<NoteEntity[]> {
  const url = buildUrl(endpoint, {
    category: "Projects",
    tableName: "Projects",
    TableName: "Projects",
    targetTable: "Projects",
  });
  return safeFetchList(url);
}

// -------------------------------
// Fetch ALL Suggestions
// -------------------------------

export async function getAllSuggestions(endpoint?: string): Promise<NoteEntity[]> {
  const url = buildUrl(endpoint, {
    category: "Suggestions",
    tableName: "Suggestions",
    TableName: "Suggestions",
    targetTable: "Suggestions",
  });
  return safeFetchList(url);
}

// -------------------------------
// Fetch Status (UID-based)
// -------------------------------

export async function getStatusForUid(uid: string, endpoint?: string): Promise<NoteEntity[]> {
  const url = buildUrl(endpoint, {
    uid,
    category: "Status",
  });
  return safeFetchList(url);
}

// -------------------------------
// Save a Note
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
  if (!res.ok) throw new Error(`saveNote failed ${res.status}: ${text}`);

  try {
    return JSON.parse(text) as SaveResponse;
  } catch {
    return { ok: true, message: text };
  }
}

// -------------------------------
// DELETE Note
// -------------------------------

export async function deleteNote(
  partitionKey: string,
  rowKey: string,
  endpoint?: string,
  tableName?: string
): Promise<void> {
  const ep = endpoint || "HttpTrigger1";
  const isAbs = /^https?:\/\//i.test(ep);
  const base = isAbs ? ep.replace(/\/?$/, "") : `${API_BASE}/${ep.replace(/^\/+/, "")}`;
  const url = base;

  const body: Record<string, string> = {
    partitionKey,
    rowKey,
  };

  if (tableName) {
    body.tableName = tableName;
    body.TableName = tableName;
    body.targetTable = tableName;
  }

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
