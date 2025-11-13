import React, { useEffect, useMemo, useState } from "react";
import { Stack, Text, TextField, PrimaryButton, Dropdown, IDropdownOption, Checkbox } from "@fluentui/react";

// Suggestion model
type Suggestion = {
  id: string;
  ts: number;
  type: string;
  summary: string;
  description: string;
  anonymous?: boolean;
  authorEmail?: string;
  authorAlias?: string;
};

const SUGGESTIONS_KEY = "uidSuggestions";

// Table Storage configuration (client-side). Prefer REACT_APP_* env vars when built.
const TABLES_ACCOUNT_URL = (process.env.REACT_APP_TABLES_ACCOUNT_URL as string) || (window as any).REACT_APP_TABLES_ACCOUNT_URL || "https://optical360.table.core.windows.net";
// Variable name requested: TABLES_TABLE_NAME_SUGGESTIONS. Try common env forms; default to 'Suggestions'.
const TABLES_TABLE_NAME_SUGGESTIONS = (process.env.REACT_APP_TABLES_TABLE_NAME_SUGGESTIONS as string) || (process.env as any).TABLES_TABLE_NAME_SUGGESTIONS || (window as any).TABLES_TABLE_NAME_SUGGESTIONS || "Suggestions";

// Helper to build a table endpoint URL. If account URL includes a query (SAS token), additional
// query params will be appended using '&', otherwise start with '?'.
function buildTableUrlForTable(tableName: string, suffix = "") {
  // Ensure no trailing slash on account URL
  const base = TABLES_ACCOUNT_URL.replace(/\/$/, "");
  const path = `${base}/${tableName}${suffix}`;
  return path;
}

const typeOptions: IDropdownOption[] = [
  { key: "Feature", text: "Feature" },
  { key: "Improvement", text: "Improvement" },
  { key: "Bug", text: "Bug" },
  { key: "UI/UX", text: "UI/UX" },
  { key: "Data", text: "Data" },
  { key: "Other", text: "Other" },
];

function getEmail(): string {
  try { return localStorage.getItem("loggedInEmail") || ""; } catch { return ""; }
}

function getAlias(email?: string | null) {
  const e = (email || "").trim();
  if (!e) return "";
  const at = e.indexOf("@");
  return at > 0 ? e.slice(0, at) : e;
}

const SuggestionsPage: React.FC = () => {
  const [items, setItems] = useState<Suggestion[]>([]);
  const [loading, setLoading] = useState<boolean>(false);
  const [error, setError] = useState<string | null>(null);

  const [type, setType] = useState<string>("Improvement");
  const [summary, setSummary] = useState<string>("");
  const [description, setDescription] = useState<string>("");
  const [anonymous, setAnonymous] = useState<boolean>(false);

  // Persist locally as a lightweight fallback, but primary source is Table Storage.
  useEffect(() => {
    try { localStorage.setItem(SUGGESTIONS_KEY, JSON.stringify(items)); } catch {}
  }, [items]);

  // Load suggestions from Azure Table Storage on mount. Falls back to localStorage when
  // Table configuration is not available or request fails.
  useEffect(() => {
    let cancelled = false;
    async function load() {
      setLoading(true);
      setError(null);
      try {
        // Build URL to query table entities. We will request JSON minimal metadata.
        const tableName = TABLES_TABLE_NAME_SUGGESTIONS || 'Suggestions';
        // Query: list entities. Use OData '()' form to query the table itself.
        // We'll request up to 500 items; ordering will be handled client-side.
        const suffix = `()?$top=500`;
        const url = buildTableUrlForTable(tableName, suffix);

        const headers: Record<string,string> = {
          'Accept': 'application/json;odata=nometadata',
          'Content-Type': 'application/json',
          // x-ms-version is often accepted by Table endpoints
          'x-ms-version': '2019-02-02'
        };

        // Try to fetch from Table Storage directly. If the account URL requires a SAS
        // token it should be included in TABLES_ACCOUNT_URL.
        let fetchedItems: Suggestion[] = [];
        try {
          // eslint-disable-next-line no-console
          console.debug('[Suggestions] GET table', url);
          const res = await fetch(url, { method: 'GET', headers, credentials: 'omit' });
          if (!res.ok) {
            throw new Error(`Table GET failed ${res.status}`);
          }
          const body = await res.json();
          // Azure Table REST returns entities under 'value'.
          const entities = Array.isArray(body?.value) ? body.value : (Array.isArray(body) ? body : []);
          fetchedItems = entities.map((e: any) => {
            const id = e.RowKey || e.rowKey || `${e.savedAt || Date.now()}-${Math.random().toString(36).slice(2,8)}`;
            const ts = (() => {
              const d = e.savedAt || e.SavedAt || e.Timestamp || e.timestamp || e.RowKey;
              const parsed = Date.parse(d || '');
              return Number.isNaN(parsed) ? (e.ts || Date.now()) : parsed;
            })();
            return {
              id: String(id),
              ts: ts,
              type: e.category || e.Category || e.type || 'Other',
              summary: e.title || e.Title || e.description?.slice?.(0,80) || e.description || '',
              description: e.description || e.Description || '',
              anonymous: !!e.anonymous,
              authorEmail: e.owner || e.Owner || undefined,
              authorAlias: e.authorAlias || e.author || undefined,
            } as Suggestion;
          });
        } catch (tableErr) {
          // If table fetch fails, fall back to localStorage silently but surface an error.
          // eslint-disable-next-line no-console
          console.warn('[Suggestions] Table fetch failed, falling back to localStorage', tableErr);
          const raw = localStorage.getItem(SUGGESTIONS_KEY);
          const arr = raw ? JSON.parse(raw) : [];
          fetchedItems = Array.isArray(arr) ? arr : [];
          if (!cancelled) setError(String((tableErr as any)?.message || tableErr));
        }

        if (!cancelled) {
          // Sort newest first by ts
          fetchedItems.sort((a,b) => (b.ts||0) - (a.ts||0));
          setItems(fetchedItems);
        }
      } catch (e: any) {
        // Top-level failure: fall back to localStorage
  // eslint-disable-next-line no-console
  console.error('[Suggestions] Load failed', e);
        try {
          const raw = localStorage.getItem(SUGGESTIONS_KEY);
          const arr = raw ? JSON.parse(raw) : [];
          if (!cancelled) setItems(Array.isArray(arr) ? arr : []);
        } catch { /* ignore */ }
  if (!cancelled) setError(String((e as any)?.message || e));
      } finally {
        if (!cancelled) setLoading(false);
      }
    }
    void load();
    return () => { cancelled = true; };
  }, []);

  const email = getEmail();
  const alias = getAlias(email);

  const submit = () => {
    const s = summary.trim();
    const d = description.trim();
    if (!s || !d) return;
    // Build entity for Azure Table Storage
    const nowIso = new Date().toISOString();
    const rowKey = nowIso;
    const tableName = TABLES_TABLE_NAME_SUGGESTIONS || 'Suggestions';
    const entity: Record<string, unknown> = {
      PartitionKey: 'Suggestions',
      RowKey: rowKey,
      category: type,
      title: s,
      description: d,
      anonymous: !!anonymous,
      owner: anonymous ? undefined : (email || 'Unknown'),
      authorAlias: anonymous ? undefined : (alias || undefined),
      savedAt: nowIso,
    };

    // Optimistically update UI while we POST. If POST fails, we keep local copy
    const optimistic: Suggestion = {
      id: rowKey,
      ts: Date.parse(nowIso),
      type,
      summary: s,
      description: d,
      anonymous,
      authorEmail: anonymous ? undefined : (email || undefined),
      authorAlias: anonymous ? undefined : (alias || undefined),
    };
    setItems([optimistic, ...items]);
    setSummary("");
    setDescription("");
    setAnonymous(false);

    // Send to Table Storage using REST API. The TABLES_ACCOUNT_URL should include any SAS token
    // required for browser-side writes. If not available, we skip network save and only keep
    // the optimistic local entry.
    (async () => {
      try {
        const url = buildTableUrlForTable(tableName, '');
        const headers: Record<string,string> = {
          'Accept': 'application/json;odata=nometadata',
          'Content-Type': 'application/json',
          'x-ms-version': '2019-02-02'
        };
        // Azure Table REST expects entity body without section wrapper in JSON
        // eslint-disable-next-line no-console
        console.debug('[Suggestions] POST table', url, entity);
        const res = await fetch(url, { method: 'POST', headers, body: JSON.stringify(entity), credentials: 'omit' });
        if (!res.ok) {
          // eslint-disable-next-line no-console
          console.warn('[Suggestions] Table POST failed', res.status, await res.text());
          // leave optimistic state; user will still see entry locally
        } else {
          // Optionally refresh list from table to get canonical data
          // (we'll do a lightweight refresh by re-running the load effect via setting items)
          // Attempt to read the response; Azure Table returns 201 with entity in body for some SDKs
          try {
            const body = await res.json().catch(() => null);
            // If body contains a saved entity, we could map and replace optimistic entry.
            // For now we keep optimistic entry as-is; the background load on next mount or manual refresh
            // will sync canonical data.
            // eslint-disable-next-line no-console
            console.debug('[Suggestions] Table POST succeeded', body);
          } catch {}
        }
      } catch (err) {
        // eslint-disable-next-line no-console
        console.warn('[Suggestions] Table POST error', err);
      }
    })();
  };

  const [expanded, setExpanded] = useState<string | null>(null);
  const sorted = useMemo(() => [...items].sort((a, b) => b.ts - a.ts), [items]);

  return (
    <div style={{ maxWidth: 900, margin: "0 auto" }}>
      <div className="vso-form-container glow" style={{ width: "100%" }}>
        <div className="banner-title">
          <span className="title-text">Suggestions</span>
          <span className="title-sub">Share ideas, fixes, and improvements</span>
        </div>

        <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
          <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
            <div style={{ width: 220 }}>
              <Dropdown
                label="Type"
                options={typeOptions}
                selectedKey={type}
                onChange={(_, opt) => setType(String(opt?.key || "Improvement"))}
                styles={{ dropdown: { width: 220 } }}
              />
            </div>
            <div style={{ flex: 1 }}>
              <TextField
                label="Name / short summary"
                placeholder="e.g., Align export columns with CIS order"
                value={summary}
                onChange={(_, v) => setSummary(v || "")}
              />
            </div>
          </div>
          <TextField
            label="Description"
            placeholder="Describe the idea, why it helps, and any details"
            multiline
            rows={4}
            value={description}
            onChange={(_, v) => setDescription(v || "")}
          />
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <Checkbox label="Post anonymously" checked={anonymous} onChange={(_, c) => setAnonymous(!!c)} />
            <PrimaryButton text="Submit suggestion" onClick={submit} disabled={!summary.trim() || !description.trim()} className="search-btn" />
          </div>
        </div>
      </div>

      <div className="notes-card" style={{ marginTop: 16 }}>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Text className="section-title">Community suggestions</Text>
          <span style={{ color: "#a6b7c6", fontSize: 12 }}>{sorted.length} total</span>
        </Stack>

        {sorted.length === 0 ? (
          <div className="note-empty">No suggestions yet. Be the first to post one.</div>
        ) : (
          <div className="notes-list">
            {sorted.map((s) => {
              const open = expanded === s.id;
              return (
                <div key={s.id} className="note-item">
                  <div className="note-header" style={{ alignItems: "center" }}>
                    <div className="note-meta" style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <span className="wf-inprogress-badge" style={{ color: "#50b3ff", border: "1px solid rgba(80,179,255,0.28)", borderRadius: 8, padding: "2px 8px", fontWeight: 700, fontSize: 12 }}>{s.type}</span>
                      <span className="note-alias" style={{ color: "#e6f1ff" }}>{s.summary}</span>
                      <span className="note-dot">·</span>
                      <span className="note-time">{new Date(s.ts).toLocaleString()}</span>
                      {!s.anonymous && (s.authorAlias || s.authorEmail) && (
                        <>
                          <span className="note-dot">·</span>
                          <span className="note-email">{s.authorAlias || s.authorEmail}</span>
                        </>
                      )}
                    </div>
                    <div className="note-controls">
                      <button className="note-btn" onClick={() => setExpanded(open ? null : s.id)} title={open ? "Collapse" : "Expand"}>
                        {open ? "Hide" : "Show"}
                      </button>
                    </div>
                  </div>
                  {open && (
                    <div className="note-body">
                      <div className="note-text" style={{ whiteSpace: "pre-wrap" }}>{s.description}</div>
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        )}
      </div>
    </div>
  );
};

export default SuggestionsPage;
