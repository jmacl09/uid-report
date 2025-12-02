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

// API endpoint for server-side function (HttpTrigger1)
const API_TRIGGER_URL =
  (process.env.REACT_APP_API_TRIGGER_URL as string) ||
  (window as any).REACT_APP_API_TRIGGER_URL ||
  "/api/HttpTrigger1";

// Table name used by HttpTrigger1 via chooseTable logic
const TABLES_TABLE_NAME_SUGGESTIONS =
  (process.env.REACT_APP_TABLES_TABLE_NAME_SUGGESTIONS as string) ||
  (window as any).TABLES_TABLE_NAME_SUGGESTIONS ||
  "Suggestions";

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
  const at = e.indexOf("@");
  return at > 0 ? e.slice(0, at) : e;
}

const SuggestionsPage: React.FC = () => {
  const [items, setItems] = useState<Suggestion[]>([]);
  const [, setLoading] = useState<boolean>(false);
  const [, setError] = useState<string | null>(null);

  const [type, setType] = useState<string>("Improvement");
  const [summary, setSummary] = useState<string>("");
  const [description, setDescription] = useState<string>("");
  const [anonymous, setAnonymous] = useState<boolean>(false);

  // Persist locally as fallback
  useEffect(() => {
    try { localStorage.setItem(SUGGESTIONS_KEY, JSON.stringify(items)); } catch {}
  }, [items]);

  // Load suggestions from Table Storage via HttpTrigger1
  useEffect(() => {
    let cancelled = false;
    async function load() {
      setLoading(true);
      setError(null);

      const tableName = TABLES_TABLE_NAME_SUGGESTIONS;
      const url = `${API_TRIGGER_URL}?uid=suggestions&tableName=${encodeURIComponent(tableName)}`;

      try {
        const res = await fetch(url, {
          method: "GET",
          headers: { Accept: "application/json" },
          credentials: "same-origin",
        });

        if (!res.ok) throw new Error(`Function GET failed ${res.status}`);

        const body = await res.json();
        const entities = Array.isArray(body?.items) ? body.items : [];

        const mapped = entities.map((e: any) => {
          const rowKey = e.rowKey || e.RowKey || e.savedAt || Date.now();
          const tsParsed = Date.parse(
            e.savedAt || e.Timestamp || e.timestamp || e.rowKey || ""
          );
          const ts = Number.isNaN(tsParsed) ? Date.now() : tsParsed;

          return {
            id: String(rowKey),
            ts,
            type: e.category || e.Category || e.type || "Other",
            summary: e.title || e.Title || e.description || "",
            description: e.description || e.Description || "",
            anonymous: !!e.anonymous,
            authorEmail: e.owner || e.Owner || undefined,
            authorAlias: e.authorAlias || e.author || undefined,
          } as Suggestion;
        });

        if (!cancelled) {
          mapped.sort((a: Suggestion, b: Suggestion) => b.ts - a.ts);
          setItems(mapped);
        }
      } catch (err) {
        // fallback to localStorage
        const raw = localStorage.getItem(SUGGESTIONS_KEY);
        const arr = raw ? JSON.parse(raw) : [];
        if (!cancelled) {
          setItems(Array.isArray(arr) ? arr : []);
          setError(String((err as any)?.message || err));
        }
      } finally {
        if (!cancelled) setLoading(false);
      }
    }

    load();
    return () => { cancelled = true; };
  }, []);

  const email = getEmail();
  const alias = getAlias(email);

  // Submit a new suggestion
  const submit = () => {
    const s = summary.trim();
    const d = description.trim();
    if (!s || !d) return;

    const nowIso = new Date().toISOString();
    const tableName = TABLES_TABLE_NAME_SUGGESTIONS;

    // Optimistic update
    const optimistic: Suggestion = {
      id: nowIso,
      ts: Date.parse(nowIso),
      type,
      summary: s,
      description: d,
      anonymous,
      authorEmail: anonymous ? undefined : email,
      authorAlias: anonymous ? undefined : alias,
    };

    setItems([optimistic, ...items]);
    setSummary("");
    setDescription("");
    setAnonymous(false);

    // Save to table storage via HttpTrigger1
    (async () => {
      try {
        const payload = {
          uid: "suggestions",
          category: type,
          title: s,
          description: d,
          owner: anonymous ? undefined : email,
          timestamp: nowIso,
          rowKey: nowIso,
          tableName,
        };

        const res = await fetch(API_TRIGGER_URL, {
          method: "POST",
          headers: {
            Accept: "application/json",
            "Content-Type": "application/json",
          },
          credentials: "same-origin",
          body: JSON.stringify(payload),
        });

        if (!res.ok) {
          setError(`Save failed: ${res.status}`);
        }
      } catch (err) {
        setError(String((err as any)?.message || err));
      }
    })();
  };

  const [expanded, setExpanded] = useState<string | null>(null);
  const sorted = useMemo(
    () => [...items].sort((a: Suggestion, b: Suggestion) => b.ts - a.ts),
    [items]
  );

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

          <div style={{ width: "100%" }}>
            <label
              style={{
                color: "var(--vso-label-color)",
                fontWeight: 600,
                fontSize: 14,
                marginBottom: 6,
                display: "block",
              }}
            >
              Description
            </label>

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
            />

            <PrimaryButton
              text="Submit suggestion"
              disabled={!summary.trim() || !description.trim()}
              onClick={submit}
            />
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
                    <div
                      className="note-meta"
                      style={{ display: "flex", alignItems: "center", gap: 8 }}
                    >
                      <span
                        className="wf-inprogress-badge"
                        style={{
                          color: "#50b3ff",
                          border: "1px solid rgba(80,179,255,0.28)",
                          borderRadius: 8,
                          padding: "2px 8px",
                          fontWeight: 700,
                          fontSize: 12,
                        }}
                      >
                        {s.type}
                      </span>

                      <span className="note-alias" style={{ color: "#e6f1ff" }}>
                        {s.summary}
                      </span>

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
                      <button
                        className="note-btn"
                        onClick={() => setExpanded(open ? null : s.id)}
                        title={open ? "Collapse" : "Expand"}
                      >
                        {open ? "Hide" : "Show"}
                      </button>
                    </div>
                  </div>

                  {open && (
                    <div className="note-body">
                      <div className="note-text" style={{ whiteSpace: "pre-wrap" }}>
                        {s.description}
                      </div>
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
