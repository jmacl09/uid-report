import React, { useEffect, useMemo, useState } from "react";
import {
  Stack,
  Text,
  TextField,
  PrimaryButton,
  Dropdown,
  IDropdownOption,
  Checkbox,
} from "@fluentui/react";

import { API_BASE } from "../api/config";

/* ---------------------------------------------------------
   Suggestion Model
--------------------------------------------------------- */
type Suggestion = {
  id: string;
  ts: number;
  type: string;
  summary: string;
  description: string;
  anonymous: boolean;
  authorEmail?: string;
  authorAlias?: string;
};

/* ---------------------------------------------------------
   Dropdown Options
--------------------------------------------------------- */
const typeOptions: IDropdownOption[] = [
  { key: "Feature", text: "Feature" },
  { key: "Improvement", text: "Improvement" },
  { key: "Bug", text: "Bug" },
  { key: "UI/UX", text: "UI/UX" },
  { key: "Data", text: "Data" },
  { key: "Other", text: "Other" },
];

function getEmail(): string {
  try {
    return localStorage.getItem("loggedInEmail") || "";
  } catch {
    return "";
  }
}

function getAlias(email?: string) {
  const e = (email || "").trim();
  const at = e.indexOf("@");
  return at > 0 ? e.slice(0, at) : e;
}

/* ---------------------------------------------------------
   MAIN COMPONENT
--------------------------------------------------------- */
const SuggestionsPage: React.FC = () => {
  const [items, setItems] = useState<Suggestion[]>([]);
  const [loading, setLoading] = useState<boolean>(false);
  const [type, setType] = useState<string>("Improvement");
  const [summary, setSummary] = useState<string>("");
  const [description, setDescription] = useState<string>("");
  const [anonymous, setAnonymous] = useState<boolean>(false);

  const email = getEmail();
  const alias = getAlias(email);

  /* ---------------------------------------------------------
     Load suggestions from API
  --------------------------------------------------------- */
  useEffect(() => {
    async function load() {
      setLoading(true);
      try {
        const res = await fetch(`${API_BASE}/HttpTrigger1?category=suggestions`);
        if (!res.ok) throw new Error("Failed to load suggestions");

        const json = await res.json();
        const rows = Array.isArray(json) ? json : json.items || [];

        const mapped: Suggestion[] = rows.map((e: any) => {
          const owner = (e.owner || "").toString();
          const anon = owner.toLowerCase() === "anonymous";

          return {
            id: e.rowKey,
            ts: Date.parse(e.savedAt || new Date().toISOString()),
            type: e.type || "Other",
            summary: e.title || "",
            description: e.description || "",
            anonymous: anon,
            authorEmail: anon ? undefined : owner,
            authorAlias: anon ? undefined : owner,
          };
        });

        mapped.sort((a: Suggestion, b: Suggestion) => b.ts - a.ts);
        setItems(mapped);
      } catch (err) {
        console.warn("Suggestions load error:", err);
      } finally {
        setLoading(false);
      }
    }

    load();
  }, []);

  /* ---------------------------------------------------------
     Submit suggestion
  --------------------------------------------------------- */
  const submit = async () => {
    const s = summary.trim();
    const d = description.trim();
    if (!s || !d) return;

    const now = Date.now();

    // Optimistic UI entry
    const optimistic: Suggestion = {
      id: `temp-${now}`,
      ts: now,
      type,
      summary: s,
      description: d,
      anonymous,
      authorEmail: anonymous ? undefined : email,
      authorAlias: anonymous ? undefined : alias,
    };

    setItems((prev) => [optimistic, ...prev]);

    // Reset fields
    setSummary("");
    setDescription("");
    setAnonymous(false);

    try {
      await fetch(`${API_BASE}/HttpTrigger1?category=suggestions`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          category: "Suggestions",
          title: s,
          description: d,
          owner: anonymous ? "Anonymous" : alias || email || "Unknown",
        }),
      });

      // Reload fresh from server
      const res = await fetch(`${API_BASE}/HttpTrigger1?category=suggestions`);
      if (!res.ok) return;

      const json = await res.json();
      const rows = Array.isArray(json) ? json : json.items || [];

      const mapped = rows.map((e: any) => {
        const owner = (e.owner || "").toString();
        const anon = owner.toLowerCase() === "anonymous";

        return {
          id: e.rowKey,
          ts: Date.parse(e.savedAt || new Date().toISOString()),
          type: e.type || "Other",
          summary: e.title || "",
          description: e.description || "",
          anonymous: anon,
          authorEmail: anon ? undefined : owner,
          authorAlias: anon ? undefined : owner,
        };
      });

      mapped.sort((a: Suggestion, b: Suggestion) => b.ts - a.ts);
      setItems(mapped);
    } catch (err) {
      console.warn("Suggestion submit failed:", err);
    }
  };

  const [expanded, setExpanded] = useState<string | null>(null);

  const sorted = useMemo(
    () => [...items].sort((a: Suggestion, b: Suggestion) => b.ts - a.ts),
    [items]
  );

  /* ---------------------------------------------------------
     RENDER
--------------------------------------------------------- */
  return (
    <div style={{ maxWidth: 900, margin: "0 auto" }}>
      <div className="vso-form-container glow" style={{ width: "100%" }}>
        <div className="banner-title">
          <span className="title-text">Suggestions</span>
          <span className="title-sub">Share ideas, fixes, and improvements</span>
        </div>

        {/* Form */}
        <div style={{ display: "flex", flexDirection: "column", gap: 12, marginTop: 16 }}>
          <div style={{ display: "flex", gap: 10 }}>
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
                label="Summary"
                placeholder="e.g., Improve search performance"
                value={summary}
                onChange={(_, v) => setSummary(v || "")}
              />
            </div>
          </div>

          <TextField
            label="Description"
            multiline
            rows={4}
            placeholder="Describe your idea…"
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
            />
          </div>
        </div>
      </div>

      {/* Suggestions List */}
      <div className="notes-card" style={{ marginTop: 20 }}>
        <Stack horizontal horizontalAlign="space-between">
          <Text className="section-title">Community suggestions</Text>
          <span style={{ color: "#a6b7c6", fontSize: 12 }}>{sorted.length} total</span>
        </Stack>

        {loading && <div className="note-empty">Loading…</div>}

        {!loading && sorted.length === 0 && (
          <div className="note-empty">No suggestions yet.</div>
        )}

        {sorted.length > 0 && (
          <div className="notes-list">
            {sorted.map((s) => {
              const open = expanded === s.id;

              return (
                <div key={s.id} className="note-item">
                  <div className="note-header">
                    <div className="note-meta" style={{ display: "flex", gap: 8 }}>
                      <span
                        className="wf-inprogress-badge"
                        style={{
                          borderRadius: 8,
                          padding: "2px 8px",
                          fontWeight: 700,
                          fontSize: 12,
                          color: "#50b3ff",
                          border: "1px solid rgba(80,179,255,0.28)",
                        }}
                      >
                        {s.type}
                      </span>

                      <Text className="note-alias">{s.summary}</Text>

                      <span className="note-dot">·</span>
                      <span className="note-time">{new Date(s.ts).toLocaleString()}</span>

                      {!s.anonymous && (s.authorAlias || s.authorEmail) && (
                        <>
                          <span className="note-dot">·</span>
                          <span className="note-email">
                            {s.authorAlias || s.authorEmail}
                          </span>
                        </>
                      )}
                    </div>

                    <button
                      className="note-btn"
                      onClick={() => setExpanded(open ? null : s.id)}
                    >
                      {open ? "Hide" : "Show"}
                    </button>
                  </div>

                  {open && (
                    <div className="note-body">
                      <div style={{ whiteSpace: "pre-wrap" }}>{s.description}</div>
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
