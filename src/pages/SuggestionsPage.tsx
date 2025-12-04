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

// Theme-friendly input/button styles for dark mode
const INPUT_BG = "#0f1112";
const INPUT_BORDER = "1px solid rgba(255,255,255,0.04)";
const INPUT_TEXT = "#e6eef6";

const textFieldStyles = {
  fieldGroup: {
    background: INPUT_BG,
    border: INPUT_BORDER,
    selectors: {
      ':hover': { border: '1px solid rgba(80,179,255,0.12)' },
    },
  },
  field: { color: INPUT_TEXT },
  label: { color: '#9fb2c9' },
};

const dropdownStyles = {
  dropdown: { background: INPUT_BG, color: INPUT_TEXT, border: INPUT_BORDER },
  title: { background: 'transparent', color: INPUT_TEXT },
  caretDown: { color: INPUT_TEXT },
  label: { color: '#9fb2c9' },
};

const buttonStyles = {
  root: {
    background: 'linear-gradient(180deg, rgba(80,179,255,0.06), rgba(80,179,255,0.04))',
    color: '#d6f5ff',
    border: '1px solid rgba(80,179,255,0.14)',
  },
  rootDisabled: {
    background: 'rgba(255,255,255,0.02)',
    color: 'rgba(255,255,255,0.5)',
    border: '1px solid rgba(255,255,255,0.02)',
  },
};

const checkboxStyles = { label: { color: '#b5c7d8' } };

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
  const [loading, setLoading] = useState(false);
  const [type, setType] = useState("Improvement");
  const [summary, setSummary] = useState("");
  const [description, setDescription] = useState("");
  const [anonymous, setAnonymous] = useState(false);

  const email = getEmail();
  const alias = getAlias(email);

  /* ---------------------------------------------------------
     Load suggestions from DB
  --------------------------------------------------------- */
  useEffect(() => {
    async function load() {
      setLoading(true);
      try {
        const res = await fetch(`${API_BASE}/HttpTrigger1?category=suggestions`);
        const json = await res.json();
        const rows = Array.isArray(json) ? json : json.items || [];

        const mapped: Suggestion[] = rows.map((e: any) => {
          const owner = (e.owner || "").toString();
          const anon = owner.toLowerCase() === "anonymous";

          return {
            id: e.rowKey,
            ts: Date.parse(e.savedAt || new Date().toISOString()),
            type: e.type || "Other",
            summary: e.summary || e.title || "",
            description: e.description || "",
            anonymous: anon,
            authorEmail: anon ? undefined : owner,
            authorAlias: anon ? undefined : owner,
          };
        });

        mapped.sort((a, b) => b.ts - a.ts);
        setItems(mapped);
      } catch (err) {
        console.error("❌ Suggestions load error:", err);
      } finally {
        setLoading(false);
      }
    }

    load();
  }, []);

  /* ---------------------------------------------------------
     Submit suggestion (DB save)
  --------------------------------------------------------- */
  const submit = async () => {
    const s = summary.trim();
    const d = description.trim();
    if (!s || !d) return;

    const owner = anonymous ? "Anonymous" : alias || email || "Unknown";
    const now = Date.now();

    // optimistic UI (instant add)
    setItems(prev => [
      {
        id: `temp-${now}`,
        ts: now,
        type,
        summary: s,
        description: d,
        anonymous,
        authorEmail: anonymous ? undefined : email,
        authorAlias: anonymous ? undefined : alias,
      },
      ...prev,
    ]);

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
          summary: s,
          description: d,
          type,
          owner,
        }),
      });

      // reload from DB to replace temp items with real saved ones
      const res = await fetch(`${API_BASE}/HttpTrigger1?category=suggestions`);
      const json = await res.json();
      const rows = Array.isArray(json) ? json : json.items || [];

      const mapped: Suggestion[] = rows.map((e: any) => {
        const owner = (e.owner || "").toString();
        const anon = owner.toLowerCase() === "anonymous";

        return {
          id: e.rowKey,
          ts: Date.parse(e.savedAt || new Date().toISOString()),
          type: e.type || "Other",
          summary: e.summary || e.title || "",
          description: e.description || "",
          anonymous: anon,
          authorEmail: anon ? undefined : owner,
          authorAlias: anon ? undefined : owner,
        };
      });

      mapped.sort((a, b) => b.ts - a.ts);
      setItems(mapped);
    } catch (err) {
      console.error("❌ submit() error:", err);
    }
  };

  const [expanded, setExpanded] = useState<string | null>(null);

  const sorted = useMemo(() => [...items].sort((a, b) => b.ts - a.ts), [items]);

  /* ---------------------------------------------------------
     RENDER UI
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
                styles={dropdownStyles}
              />
            </div>

            <div style={{ flex: 1 }}>
              <TextField
                label="Summary"
                placeholder="Short title"
                value={summary}
                onChange={(_, v) => setSummary(v || "")}
                styles={textFieldStyles}
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
            styles={textFieldStyles}
          />

          {/* Single real submit button */}
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <Checkbox
              label="Post anonymously"
              checked={anonymous}
              onChange={(_, c) => setAnonymous(!!c)}
              styles={checkboxStyles}
            />

            <PrimaryButton
              text="Submit suggestion"
              disabled={!summary.trim() || !description.trim()}
              onClick={submit}
              styles={buttonStyles}
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
