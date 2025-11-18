import { API_BASE } from "./config";

// Absolute Function URL fallback (used only when the proxied/API_BASE endpoint fails).
const DEFAULT_FUNCTION_URL = 'https://optical360v2-ffa9ewbfafdvfyd8.westeurope-01.azurewebsites.net/api/HttpTrigger1';

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

  try {
    const res = await fetch(url, { method: 'GET' });
    const text = await res.text();
    if (!res.ok) {
      // don't throw to avoid uncaught runtime errors in the UI; surface via empty list
      // eslint-disable-next-line no-console
      console.warn(`getNotesForUid returned ${res.status}: ${text}`);
      return [];
    }
    try {
      const data = JSON.parse(text);
      if (data?.items && Array.isArray(data.items)) return data.items as NoteEntity[];
      return [];
    } catch {
      return [];
    }
  } catch (networkErr) {
    // Network/CORS or other fetch-level failure — don't throw to keep UI stable
    // eslint-disable-next-line no-console
    console.warn('[getNotesForUid] Network error', networkErr);
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

  try {
    const res = await fetch(url, { method: 'GET' });
    const text = await res.text();
    if (!res.ok) {
      // eslint-disable-next-line no-console
      console.warn(`getCommentsForUid returned ${res.status}: ${text}`);
      return [];
    }
    try {
      const data = JSON.parse(text);
      if (data?.items && Array.isArray(data.items)) return data.items as NoteEntity[];
      return [];
    } catch {
      return [];
    }
  } catch (networkErr) {
    // eslint-disable-next-line no-console
    console.warn('[getCommentsForUid] Network error', networkErr);
    return [];
  }
}

/**
 * Fetch calendar entries saved under a given UID. If no UID provided, callers may pass
 * the logical UID used by calendar saves (e.g. 'VsoCalendar'). Returns raw NoteEntity list.
 */
export async function getCalendarEntries(uid: string, endpoint?: string): Promise<NoteEntity[]> {
  const rawEndpoint = endpoint || 'HttpTrigger1';
  const isAbsolute = /^https?:\/\//i.test(rawEndpoint);
  const base = isAbsolute ? rawEndpoint.replace(/\/?$/,'') : `${API_BASE}/${rawEndpoint.replace(/^\/+/, '')}`;
  const url = `${base}?uid=${encodeURIComponent(uid)}&category=${encodeURIComponent('Calendar')}`;

  const parseList = (text: string) => {
    try {
      const data = JSON.parse(text);
      if (Array.isArray(data)) return data as NoteEntity[];
      if (data?.items && Array.isArray(data.items)) return data.items as NoteEntity[];
      if (data?.value && Array.isArray(data.value)) return data.value as NoteEntity[];
      if (data?.entity && Array.isArray(data.entity)) return data.entity as NoteEntity[];
      if (data?.Entity && Array.isArray(data.Entity)) return data.Entity as NoteEntity[];
      return [];
    } catch {
      return [];
    }
  };

  try {
    const res = await fetch(url, { method: 'GET' });
    const text = await res.text();
    if (!res.ok) {
      // eslint-disable-next-line no-console
      console.warn(`getCalendarEntries returned ${res.status}: ${text}`);
    }
    let parsed = parseList(text);

    // If proxied API returned nothing and we used a non-absolute endpoint, try the absolute Function URL as a fallback
    if ((!parsed || !parsed.length) && !isAbsolute && DEFAULT_FUNCTION_URL) {
      try {
        const fbRes = await fetch(`${DEFAULT_FUNCTION_URL}?uid=${encodeURIComponent(uid)}&category=${encodeURIComponent('Calendar')}`, { method: 'GET' });
        const fbText = await fbRes.text();
        parsed = parseList(fbText);
        // eslint-disable-next-line no-console
        console.debug('[getCalendarEntries] fallback parsedCount', parsed?.length || 0);
      } catch (fbErr) {
        // eslint-disable-next-line no-console
        console.warn('[getCalendarEntries] fallback network error', fbErr);
      }
    }

    return parsed || [];
  } catch (networkErr) {
    // eslint-disable-next-line no-console
    console.warn('[getCalendarEntries] Network error', networkErr);
    return [];
  }
}

/**
 * Fetch troubleshooting entries for a given UID.
 * Mirrors the same HttpTrigger1 contract as notes/calendar and requests category=Troubleshooting.
 */
export async function getTroubleshootingForUid(uid: string, endpoint?: string): Promise<NoteEntity[]> {
  const rawEndpoint = endpoint || 'HttpTrigger1';
  const isAbsolute = /^https?:\/\//i.test(rawEndpoint);
  const base = isAbsolute ? rawEndpoint.replace(/\/?$/,'') : `${API_BASE}/${rawEndpoint.replace(/^\/+/, '')}`;
  const url = `${base}?uid=${encodeURIComponent(uid)}&category=${encodeURIComponent('Troubleshooting')}&tableName=${encodeURIComponent('Troubleshooting')}&TableName=${encodeURIComponent('Troubleshooting')}&targetTable=${encodeURIComponent('Troubleshooting')}`;

  try {
    const res = await fetch(url, { method: 'GET' });
    const text = await res.text();
    if (!res.ok) {
      // eslint-disable-next-line no-console
      console.warn(`getTroubleshootingForUid returned ${res.status}: ${text}`);
      return [];
    }
    try {
      const data = JSON.parse(text);
      if (data?.items && Array.isArray(data.items)) return data.items as NoteEntity[];
      return [];
    } catch {
      return [];
    }
  } catch (networkErr) {
    // eslint-disable-next-line no-console
    console.warn('[getTroubleshootingForUid] Network error', networkErr);
    return [];
  }
}

/**
 * Fetch project entries saved for a given UID. Uses category=Projects so the Function
 * routes to the Projects table (or explicit table override via payload/tableName).
 */
export async function getProjectsForUid(uid: string, endpoint?: string): Promise<NoteEntity[]> {
  const rawEndpoint = endpoint || 'HttpTrigger1';
  const isAbsolute = /^https?:\/\//i.test(rawEndpoint);
  const base = isAbsolute ? rawEndpoint.replace(/\/?$/,'') : `${API_BASE}/${rawEndpoint.replace(/^\/+/, '')}`;
  const url = `${base}?uid=${encodeURIComponent(uid)}&category=${encodeURIComponent('Projects')}`;
  // Helper to parse a response text into an array of NoteEntity if possible
  const parseList = (text: string) => {
    try {
      const data = JSON.parse(text);
      if (Array.isArray(data)) return data as NoteEntity[];
      if (data?.items && Array.isArray(data.items)) return data.items as NoteEntity[];
      if (data?.value && Array.isArray(data.value)) return data.value as NoteEntity[];
      // Some function responses wrap the entity under 'entity' or 'Entity'
      if (data?.entity && Array.isArray(data.entity)) return data.entity as NoteEntity[];
      if (data?.Entity && Array.isArray(data.Entity)) return data.Entity as NoteEntity[];
      return [];
    } catch {
      return [];
    }
  };

  try {
    const res = await fetch(url, { method: 'GET' });
    const text = await res.text();
    if (!res.ok) {
      // eslint-disable-next-line no-console
      console.warn(`getProjectsForUid returned ${res.status}: ${text}`);
    }
  let parsed = parseList(text);
  // Debug: surface what we received from the proxied endpoint
  // eslint-disable-next-line no-console
  console.debug('[getProjectsForUid] proxied url', url, 'parsedCount', parsed?.length || 0);
    // If proxied API returned nothing and we used a non-absolute endpoint, try the absolute Function URL as a fallback
    if ((!parsed || !parsed.length) && !isAbsolute && DEFAULT_FUNCTION_URL) {
      try {
        // eslint-disable-next-line no-console
        console.debug('[getProjectsForUid] trying fallback url', `${DEFAULT_FUNCTION_URL}?uid=${encodeURIComponent(uid)}&category=${encodeURIComponent('Projects')}`);
        const fbRes = await fetch(`${DEFAULT_FUNCTION_URL}?uid=${encodeURIComponent(uid)}&category=${encodeURIComponent('Projects')}`, { method: 'GET' });
        const fbText = await fbRes.text();
        parsed = parseList(fbText);
        // eslint-disable-next-line no-console
        console.debug('[getProjectsForUid] fallback parsedCount', parsed?.length || 0);
      } catch (fbErr) {
        // eslint-disable-next-line no-console
        console.warn('[getProjectsForUid] fallback network error', fbErr);
      }
    }
    return parsed || [];
  } catch (networkErr) {
    // eslint-disable-next-line no-console
    console.warn('[getProjectsForUid] Network error', networkErr);
    return [];
  }
}

/**
 * Fetch all Projects rows (no uid) so clients can pre-load project snapshots.
 */
export async function getAllProjects(endpoint?: string): Promise<NoteEntity[]> {
  const rawEndpoint = endpoint || 'HttpTrigger1';
  const isAbsolute = /^https?:\/\//i.test(rawEndpoint);
  const base = isAbsolute ? rawEndpoint.replace(/\/?$/,'') : `${API_BASE}/${rawEndpoint.replace(/^\/+/, '')}`;
  const url = `${base}?category=${encodeURIComponent('Projects')}&tableName=${encodeURIComponent('Projects')}&TableName=${encodeURIComponent('Projects')}&targetTable=${encodeURIComponent('Projects')}`;
  const parseList = (text: string) => {
    try {
      const data = JSON.parse(text);
      if (Array.isArray(data)) return data as NoteEntity[];
      if (data?.items && Array.isArray(data.items)) return data.items as NoteEntity[];
      if (data?.value && Array.isArray(data.value)) return data.value as NoteEntity[];
      if (data?.entity && Array.isArray(data.entity)) return data.entity as NoteEntity[];
      if (data?.Entity && Array.isArray(data.Entity)) return data.Entity as NoteEntity[];
      return [];
    } catch {
      return [];
    }
  };

  try {
    const res = await fetch(url, { method: 'GET' });
    const text = await res.text();
    if (!res.ok) {
      // eslint-disable-next-line no-console
      console.warn(`getAllProjects returned ${res.status}: ${text}`);
    }
    let parsed = parseList(text);
    // Debug: show proxied fetch result
    // eslint-disable-next-line no-console
    console.debug('[getAllProjects] proxied url', url, 'parsedCount', parsed?.length || 0);
    if ((!parsed || !parsed.length) && !isAbsolute && DEFAULT_FUNCTION_URL) {
      try {
        // eslint-disable-next-line no-console
        console.debug('[getAllProjects] trying fallback url', `${DEFAULT_FUNCTION_URL}?category=${encodeURIComponent('Projects')}`);
        const fbRes = await fetch(`${DEFAULT_FUNCTION_URL}?category=${encodeURIComponent('Projects')}&tableName=${encodeURIComponent('Projects')}&TableName=${encodeURIComponent('Projects')}&targetTable=${encodeURIComponent('Projects')}`, { method: 'GET' });
        const fbText = await fbRes.text();
        parsed = parseList(fbText);
        // eslint-disable-next-line no-console
        console.debug('[getAllProjects] fallback parsedCount', parsed?.length || 0);
      } catch (fbErr) {
        // eslint-disable-next-line no-console
        console.warn('[getAllProjects] fallback network error', fbErr);
      }
    }
    return parsed || [];
  } catch (networkErr) {
    // eslint-disable-next-line no-console
    console.warn('[getAllProjects] Network error', networkErr);
    return [];
  }
}

/**
 * Fetch all Suggestions rows (no uid) so the Suggestions page can show community submissions.
 */
export async function getAllSuggestions(endpoint?: string): Promise<NoteEntity[]> {
  const rawEndpoint = endpoint || 'HttpTrigger1';
  const isAbsolute = /^https?:\/\//i.test(rawEndpoint);
  const base = isAbsolute ? rawEndpoint.replace(/\/?$/,'') : `${API_BASE}/${rawEndpoint.replace(/^\/+/, '')}`;
  const url = `${base}?category=${encodeURIComponent('Suggestions')}&tableName=${encodeURIComponent('Suggestions')}&TableName=${encodeURIComponent('Suggestions')}&targetTable=${encodeURIComponent('Suggestions')}`;
  const parseList = (text: string) => {
    try {
      const data = JSON.parse(text);
      if (Array.isArray(data)) return data as NoteEntity[];
      if (data?.items && Array.isArray(data.items)) return data.items as NoteEntity[];
      if (data?.value && Array.isArray(data.value)) return data.value as NoteEntity[];
      if (data?.entity && Array.isArray(data.entity)) return data.entity as NoteEntity[];
      if (data?.Entity && Array.isArray(data.Entity)) return data.Entity as NoteEntity[];
      return [];
    } catch {
      return [];
    }
  };

  try {
    const res = await fetch(url, { method: 'GET' });
    const text = await res.text();
    if (!res.ok) {
      // eslint-disable-next-line no-console
      console.warn(`getAllSuggestions returned ${res.status}: ${text}`);
    }
    let parsed = parseList(text);
    // If proxied API returned nothing and we used a non-absolute endpoint, try the absolute Function URL as a fallback
    if ((!parsed || !parsed.length) && !isAbsolute) {
      try {
        const DEFAULT_FUNCTION_URL = 'https://optical360v2-ffa9ewbfafdvfyd8.westeurope-01.azurewebsites.net/api/HttpTrigger1';
        const fbRes = await fetch(`${DEFAULT_FUNCTION_URL}?category=${encodeURIComponent('Suggestions')}&tableName=${encodeURIComponent('Suggestions')}&TableName=${encodeURIComponent('Suggestions')}&targetTable=${encodeURIComponent('Suggestions')}`, { method: 'GET' });
        const fbText = await fbRes.text();
        parsed = parseList(fbText);
      } catch (fbErr) {
        // eslint-disable-next-line no-console
        console.warn('[getAllSuggestions] fallback network error', fbErr);
      }
    }
    return parsed || [];
  } catch (networkErr) {
    // eslint-disable-next-line no-console
    console.warn('[getAllSuggestions] Network error', networkErr);
    return [];
  }
}

/**
 * Fetch status entries for a given UID (category=Status).
 * This will be used by the UI to load persisted status fields such as
 * expectedDeliveryDate for the UIDStatusPanel.
 */
export async function getStatusForUid(uid: string, endpoint?: string): Promise<NoteEntity[]> {
  const rawEndpoint = endpoint || 'HttpTrigger1';
  const isAbsolute = /^https?:\/\//i.test(rawEndpoint);
  const base = isAbsolute ? rawEndpoint.replace(/\/?$/,'') : `${API_BASE}/${rawEndpoint.replace(/^\/+/, '')}`;
  const url = `${base}?uid=${encodeURIComponent(uid)}&category=${encodeURIComponent('Status')}`;

  try {
    const res = await fetch(url, { method: 'GET' });
    const text = await res.text();
    if (!res.ok) {
      // eslint-disable-next-line no-console
      console.warn(`getStatusForUid returned ${res.status}: ${text}`);
      return [];
    }
    try {
      const data = JSON.parse(text);
      if (data?.items && Array.isArray(data.items)) return data.items as NoteEntity[];
      return [];
    } catch {
      return [];
    }
  } catch (networkErr) {
    // eslint-disable-next-line no-console
    console.warn('[getStatusForUid] Network error', networkErr);
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
 * Delete a note by partition/row key via the HttpTrigger1 function.
 */
export async function deleteNote(
  partitionKey: string,
  rowKey: string,
  endpoint?: string,
  tableName?: string
): Promise<void> {
  const rawEndpoint = endpoint || 'HttpTrigger1';
  const isAbsolute = /^https?:\/\//i.test(rawEndpoint);
  const base = isAbsolute ? rawEndpoint.replace(/\/?$/,'') : `${API_BASE}/${rawEndpoint.replace(/^\/+/, '')}`;
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
  const uidMatch = /^UID_(.+)$/i.exec(partitionKey || '');
  if (uidMatch && uidMatch[1]) body.uid = uidMatch[1];

  const isCrossOrigin = (() => {
    if (isAbsolute) return true;
    try {
      if (typeof window === 'undefined') return false;
      const target = new URL(url, window.location.href);
      return target.origin !== window.location.origin;
    } catch {
      return false;
    }
  })();

  const res = await fetch(url, {
    method: 'DELETE',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
    credentials: isCrossOrigin ? 'omit' : 'include',
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`deleteNote failed ${res.status}: ${text}`);
  }
}
