import React, { useEffect, useMemo, useState } from "react";
import {
  Stack,
  Text,
  TextField,
  PrimaryButton,
  Dropdown,
  IDropdownOption,
  Checkbox
} from "@fluentui/react";

import { API_BASE } from "../api/config";

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

// Dropdown options
const typeOptions: IDropdownOption[] = [
  { key: "Feature", text: "Feature" },
  { key: "Improvement", text: "Improvement" },
  { key: "Bug", text: "Bug" },
  { key: "UI/UX", text: "UI/UX" },
  { key: "Data", text: "Data" },
  { key: "Other", text: "Other" }
];

function getEmail(): string {
  try {
    return localStorage.getItem("loggedInEmail") || "";
  } catch {
    return "";
  }
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
    } catch {
      return [];
    }
  });

  // Form state
  const [type, setType] = useState<string>("Improvement");
  const [summary, setSummary] = useState<string>("");
  const [description, setDescription] = useState<string>("");
  const [anonymous, setAnonymous] = useState<boolean>(false);

  // Persist items locally (fallback)
  useEffect(() => {
    try {
      localStorage.setItem(SUGGESTIONS_KEY, JSON.stringify(items));
    } catch {}
  }, [items]);

  // Load suggestions from backend
  useEffect(() => {
    let cancelled = false;

    async function load() {
      try {
        const res = await fetch(`${API_BASE}/suggestions`, {
          method: "GET",
          headers: { Accept: "application/json" }
        });

        if (!res.ok) throw new Error(`Failed to load suggestions: ${res.status}`);

        const rows = await res.json();
        if (cancelled || !Array.isArray(rows)) return;

        const mapped: Suggestion[] = rows.map((e: any) => {
          const owner = (e.owner || e.Owner || "").toString();
          const anonymous = owner.toLowerCase() === "anonymous";

          return {
            id: (e.rowKey || e.RowKey || "").toString(),
            ts: Number.isFinite(Date.parse(e.timestamp || e.Timestamp || "")) ?
              Date.parse(e.timestamp || e.Timestamp) :
              Date.now(),
            type: (e.category || e.Category || "Other").toString(),
            summary: (e.title || e.Title || "").toString(),
            description: (e.description || e.Description || "").toString(),
            anonymous,
            authorAlias: anonymous ? undefined : owner,
            authorEmail: anonymous ? undefined : owner
          };
        });

        // newest first
        mapped.sort((a, b) => b.ts - a.ts);

        setItems(mapped);
      } catch {
        // keep local fallback only
      }
    }

    load();
    return () => {
      cancelled = true;
    };
  }, []);

  const email = getEmail();
  const alias = getAlias(email);

  // Submit a suggestion
  const submit = async () => {
    const s = summary.trim();
    const d = description.trim();
    if (!s || !d) return;

    // Optimistic entry
    const optimistic: Suggestion = {
      id: `opt-${Date.now()}`,
      ts: Date.now(),
      type,
      summary: s,
      description: d,
      anonymous,
      authorEmail: anonymous ? undefined : email,
      authorAlias: anonymous ? undefined : alias
    };

    setItems(prev => [optimistic, ...prev]);
    setSummary("");
    setDescription("");
    setAnonymous(false);

    try {
      await fetch(`${API_BASE}/suggestions`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          category: type,
          title: s,
          description: d,
          owner: anonymous ? "Anonymous" : (alias || email || "Unknown")
        })
      });

      // Refresh after submit
      const res = await fetch(`${API_BASE}/suggestions`);
      if (res.ok) {
        const rows = await res.json();
        const mapped: Suggestion[] = rows.map((e: any) => {
          const owner = (e.owner || "").toString();
          const anon = owner.toLowerCase() === "anonymous";

          return {
            id: (e.rowKey || "").toString(),
            ts: Number.isFinite(Date.parse(e.timestamp)) ? Date.parse(e.timestamp) : Date.now(),
            type: (e.category || "Other").toString(),
            summary: (e.title || "").toString(),
            description: (e.description || "").toString(),
            anonymous: anon,
            authorAlias: anon ? undefined : owner,
            authorEmail: anon ? undefined : owner
          };
        });

        mapped.sort((a, b) => b.ts - a.ts);
        setItems(mapped);
      }
    } catch {
      // keep optimistic only
    }
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

        {/* Form */}
        <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
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
              disabled={!summary.trim() || !description.trim()}
              onClick={submit}
              className="search-btn"
            />
          </div>
        </div>
      </div>

      {/* Suggestions list */}
      <div className="notes-card" style={{ marginTop: 16 }}>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Text className="section-title">Community suggestions</Text>
          <span style={{ color: "#a6b7c6", fontSize: 12 }}>{sorted.length} total</span>
        </Stack>

        {sorted.length === 0 ? (
          <div className="note-empty">No suggestions yet. Be the first to post one.</div>
        ) : (
          <div className="notes-list">
            {sorted.map(s => {
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
                          fontSize: 12
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
