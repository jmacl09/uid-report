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

// API endpoint for server-side function (HttpTrigger1). Prefer REACT_APP_API_TRIGGER_URL when built.
const API_TRIGGER_URL = (process.env.REACT_APP_API_TRIGGER_URL as string) || (window as any).REACT_APP_API_TRIGGER_URL || '/api/HttpTrigger1';
// Variable name requested: TABLES_TABLE_NAME_SUGGESTIONS. Use env var when present; default to the provided name 'Sugestions'.
const TABLES_TABLE_NAME_SUGGESTIONS = (process.env.REACT_APP_TABLES_TABLE_NAME_SUGGESTIONS as string) || (process.env as any).TABLES_TABLE_NAME_SUGGESTIONS || (window as any).TABLES_TABLE_NAME_SUGGESTIONS || "Sugestions";

// Helper to build a table endpoint URL. If account URL includes a query (SAS token), additional
// query params will be appended using '&', otherwise start with '?'.
// We use the server-side function `HttpTrigger1` for table operations,
// so no direct table URL builder is necessary in the client.

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
  // loading/error are only used to set state; we don't read them in the UI so avoid unused-var by omitting the read-side
  const [, setLoading] = useState<boolean>(false);
  const [, setError] = useState<string | null>(null);

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
        // Use the server-side function to load suggestions. Provide uid to namespace suggestions
        // and tableName to target the 'Sugestions' table in the function's chooseTable logic.
        const tableName = TABLES_TABLE_NAME_SUGGESTIONS || 'Sugestions';
        const url = `${API_TRIGGER_URL}?uid=suggestions&tableName=${encodeURIComponent(tableName)}`;
        const headers: Record<string,string> = { 'Accept': 'application/json' };

        let fetchedItems: Suggestion[] = [];
        try {
          // eslint-disable-next-line no-console
          console.debug('[Suggestions] GET via function', url);
          const res = await fetch(url, { method: 'GET', headers, credentials: 'same-origin' });
          if (!res.ok) throw new Error(`Function GET failed ${res.status}`);
          const body = await res.json();
          const entities = Array.isArray(body?.items) ? body.items : [];
          fetchedItems = entities.map((e: any) => {
            const id = e.rowKey || e.RowKey || `${e.savedAt || Date.now()}-${Math.random().toString(36).slice(2,8)}`;
            const ts = (() => {
              const d = e.savedAt || e.SavedAt || e.Timestamp || e.timestamp || e.rowKey || e.RowKey;
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
        } catch (fnErr) {
          // Fallback to localStorage if function call fails
          // eslint-disable-next-line no-console
          console.warn('[Suggestions] Function GET failed, falling back to localStorage', fnErr);
          const raw = localStorage.getItem(SUGGESTIONS_KEY);
          const arr = raw ? JSON.parse(raw) : [];
          fetchedItems = Array.isArray(arr) ? arr : [];
          if (!cancelled) setError(String((fnErr as any)?.message || fnErr));
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
  const tableName = TABLES_TABLE_NAME_SUGGESTIONS || 'Sugestions';
    // `entity` structure prepared for Table Storage (not used directly client-side); kept here as documentation of server shape

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

    // Send suggestion to the server-side function (HttpTrigger1). The function will persist
    // to the configured table server-side. We provide uid='suggestions' so entries are namespaced.
    (async () => {
      try {
        const payload = {
          uid: 'suggestions',
          category: type,
          title: s,
          description: d,
          owner: anonymous ? undefined : (email || 'Unknown'),
          timestamp: nowIso,
          rowKey: rowKey,
          tableName: tableName,
        } as any;

        const url = API_TRIGGER_URL;
        const headers: Record<string,string> = {
          'Accept': 'application/json',
          'Content-Type': 'application/json'
        };
        // eslint-disable-next-line no-console
        console.debug('[Suggestions] POST via function', url, payload);
        const res = await fetch(url, { method: 'POST', headers, body: JSON.stringify(payload), credentials: 'same-origin' });
        if (!res.ok) {
          // eslint-disable-next-line no-console
          console.warn('[Suggestions] Function POST failed', res.status, await res.text());
          if (!anonymous) setError(`Save failed: ${res.status}`);
        } else {
          try {
            const body = await res.json().catch(() => null);
            // eslint-disable-next-line no-console
            console.debug('[Suggestions] Function POST succeeded', body);
          } catch {}
        }
      } catch (err) {
        // eslint-disable-next-line no-console
        console.warn('[Suggestions] Function POST error', err);
        if (!anonymous) setError(String((err as any)?.message || err));
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

        <div className="suggestions-form" style={{ display: "flex", flexDirection: "column", gap: 12 }}>
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
          <div style={{ width: '100%' }}>
            <label style={{ color: 'var(--vso-label-color)', fontWeight: 600, fontSize: 14, marginBottom: 6, display: 'block' }}>Description</label>
            <TextField
              placeholder="Describe the idea, why it helps, and any details"
              multiline
              rows={4}
              value={description}
              onChange={(_, v) => setDescription(v || "")}
            />
          </div>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <Checkbox
              label="Post anonymously"
              checked={anonymous}
              onChange={(_, c) => setAnonymous(!!c)}
              styles={(props) => ({
                root: { display: 'flex', alignItems: 'center' },
                label: { color: 'var(--vso-label-color)', fontWeight: 600, fontSize: 13 },
                checkbox: {
                  borderColor: 'var(--vso-label-color)',
                  background: '#ffffff',
                  selectors: {
                    '&.is-checked': {
                      background: 'var(--vso-label-color)',
                      borderColor: 'var(--vso-label-color)'
                    }
                  }
                },
                checkmark: { color: '#ffffff', fontWeight: 700 },
                text: { color: 'var(--vso-label-color)', fontWeight: 600 }
              })}
            />
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
