import React, { useEffect, useMemo, useState } from "react";
import { Stack, Text, TextField, PrimaryButton, IconButton, Dropdown, IDropdownOption, Checkbox } from "@fluentui/react";
import { saveToStorage } from "../api/saveToStorage";
import { getAllSuggestions } from "../api/items";

// Simple suggestion model
type Suggestion = {
  id: string;
  ts: number;
  type: string; // Feature, Improvement, Bug, UI/UX, Data, Other
  summary: string; // short name/title
  description: string; // full text
  anonymous?: boolean;
  authorEmail?: string;
  authorAlias?: string;
};

const SUGGESTIONS_KEY = "uidSuggestions";

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
  const [items, setItems] = useState<Suggestion[]>(() => {
    try {
      const raw = localStorage.getItem(SUGGESTIONS_KEY);
      const arr = raw ? JSON.parse(raw) : [];
      return Array.isArray(arr) ? arr : [];
    } catch { return []; }
  });

  // form state
  const [type, setType] = useState<string>("Improvement");
  const [summary, setSummary] = useState<string>("");
  const [description, setDescription] = useState<string>("");
  const [anonymous, setAnonymous] = useState<boolean>(false);

  // persist on change
  useEffect(() => {
    try { localStorage.setItem(SUGGESTIONS_KEY, JSON.stringify(items)); } catch {}
  }, [items]);

  // Load server-side suggestions on mount
  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const rows = await getAllSuggestions();
        if (cancelled) return;
        const mapped: Suggestion[] = (rows || []).map((e: any) => {
          const rk = (e.rowKey || e.RowKey || e.rowkey || '').toString();
          const saved = (e.savedAt || e.savedAt || e.Timestamp || e.timestamp || e.savedAt || rk || new Date().toISOString()).toString();
          const ts = Number.isFinite(Date.parse(saved)) ? Date.parse(saved) : Date.now();
          return {
            id: rk || `${Date.now()}-${Math.random().toString(36).slice(2,8)}`,
            ts,
            type: (e.category || e.Category || 'Other').toString(),
            summary: (e.title || e.Title || '').toString(),
            description: (e.description || e.Description || '').toString(),
            anonymous: String((e.owner || e.Owner || '')).toLowerCase() === 'anonymous',
            authorAlias: (e.owner || e.Owner || '').toString() || undefined,
          } as Suggestion;
        });
        if (mapped && mapped.length) setItems(prev => {
          // merge server items with local optimistic ones (by id)
          const existingIds = new Set(mapped.map(m => m.id));
          const leftovers = prev.filter(p => !existingIds.has(p.id));
          return [...mapped, ...leftovers];
        });
      } catch (e) {
        // ignore failures; keep local-only suggestions
      }
    })();
    return () => { cancelled = true; };
  }, []);

  const email = getEmail();
  const alias = getAlias(email);

  const submit = () => {
    const s = summary.trim();
    const d = description.trim();
    if (!s || !d) return;
    const id = `${Date.now()}-${Math.random().toString(36).slice(2,8)}`;
    const next: Suggestion = {
      id,
      ts: Date.now(),
      type,
      summary: s,
      description: d,
      anonymous,
      authorEmail: anonymous ? undefined : (email || undefined),
      authorAlias: anonymous ? undefined : (alias || undefined),
    };
    // optimistic update
    setItems(prev => [next, ...prev]);
    // reset
    setSummary("");
    setDescription("");
    setAnonymous(false);

    // Fire-and-forget save to server. Use a shared 'public' uid so rows land in the Suggestions table
    try {
      void saveToStorage({
        category: 'Suggestions',
        uid: 'public',
        title: s,
        description: d,
        owner: anonymous ? 'Anonymous' : (alias || email || 'Unknown'),
      }).then(async (resultText) => {
        try {
          const parsed = JSON.parse(resultText);
          const entity = (parsed?.entity || parsed?.Entity) as any | undefined;
          if (entity) {
            const rk = (entity.rowKey || entity.RowKey || '').toString();
            const saved = (entity.savedAt || entity.timestamp || entity.Timestamp || rk || new Date().toISOString()).toString();
            const ts = Number.isFinite(Date.parse(saved)) ? Date.parse(saved) : Date.now();
            const official: Suggestion = {
              id: rk || `${Date.now()}-${Math.random().toString(36).slice(2,8)}`,
              ts,
              type: (entity.category || entity.Category || 'Other').toString(),
              summary: (entity.title || entity.Title || '').toString(),
              description: (entity.description || entity.Description || '').toString(),
              anonymous: String((entity.owner || entity.Owner || '')).toLowerCase() === 'anonymous',
              authorAlias: (entity.owner || entity.Owner || '').toString() || undefined,
            };
            // replace optimistic entry (match by summary+description and optimistic id) conservatively
            setItems(prev => {
              const idx = prev.findIndex(x => x.id === id);
              if (idx === -1) return [official, ...prev];
              const copy = [...prev]; copy[idx] = official; return copy;
            });
          }
        } catch {
          // ignore parse errors
        }

        // Refresh list after short delay to pick up other people's suggestions too
        await new Promise(resolve => setTimeout(resolve, 900));
        try {
          const rows = await getAllSuggestions();
          if (rows && rows.length) {
            const mapped = rows.map((e: any) => {
              const rk = (e.rowKey || e.RowKey || e.rowkey || '').toString();
              const saved = (e.savedAt || e.Timestamp || e.timestamp || rk || new Date().toISOString()).toString();
              const ts = Number.isFinite(Date.parse(saved)) ? Date.parse(saved) : Date.now();
              return {
                id: rk || `${Date.now()}-${Math.random().toString(36).slice(2,8)}`,
                ts,
                type: (e.category || e.Category || 'Other').toString(),
                summary: (e.title || e.Title || '').toString(),
                description: (e.description || e.Description || '').toString(),
                anonymous: String((e.owner || e.Owner || '')).toLowerCase() === 'anonymous',
                authorAlias: (e.owner || e.Owner || '').toString() || undefined,
              } as Suggestion;
            });
            setItems(prev => {
              const existingIds = new Set(mapped.map(m => m.id));
              const leftovers = prev.filter(p => !existingIds.has(p.id));
              return [...mapped, ...leftovers];
            });
          }
        } catch {}
      }).catch((e) => {
        // eslint-disable-next-line no-console
        console.warn('Failed to save suggestion to server (kept locally):', e?.body || e?.message || e);
      });
    } catch {}
  };

  const [expanded, setExpanded] = useState<string | null>(null);

  const sorted = useMemo(() => {
    return [...items].sort((a, b) => b.ts - a.ts);
  }, [items]);

  return (
    <div style={{ maxWidth: 900, margin: "0 auto" }}>
      {/* Header */}
      <div className="vso-form-container glow" style={{ width: "100%" }}>
        <div className="banner-title">
          <span className="title-text">Suggestions</span>
          <span className="title-sub">Share ideas, fixes, and improvements</span>
        </div>

        {/* Form */}
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
            <Checkbox
              label="Post anonymously"
              checked={anonymous}
              onChange={(_, c) => setAnonymous(!!c)}
            />
            <PrimaryButton
              text="Submit suggestion"
              onClick={submit}
              disabled={!summary.trim() || !description.trim()}
              className="search-btn"
            />
          </div>
        </div>
      </div>

      {/* List */}
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
