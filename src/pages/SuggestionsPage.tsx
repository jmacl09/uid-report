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

const SuggestionsPage: React.FC = () => {
  const [items, setItems] = useState<Suggestion[]>([]);
  const [loading, setLoading] = useState(false);
  const [type, setType] = useState("Improvement");
  const [summary, setSummary] = useState("");
  const [description, setDescription] = useState("");
  const [anonymous, setAnonymous] = useState(false);

  const email = getEmail();
  const alias = getAlias(email);

  /* LOAD */
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
      } finally {
        setLoading(false);
      }
    }

    load();
  }, []);

  /* AUTO POST */
  useEffect(() => {
    async function post() {
      await fetch(`${API_BASE}/HttpTrigger1?category=suggestions`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          category: "Suggestions",
          title: "AutoTest",
          summary: "AutoTest summary",
          description: "Posted on mount",
          type: "Auto",
          owner: "AutoTester",
        }),
      });
    }
    post();
  }, []);

  /* SUBMIT */
  const submit = async () => {
    const s = summary.trim();
    const d = description.trim();
    if (!s || !d) return;

    const owner = anonymous ? "Anonymous" : alias || email || "Unknown";
    const now = Date.now();

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

        {/* TEST MARKER */}
        <div style={{ fontSize: 72, color: "red", textAlign: "center" }}>+</div>

        {/* FORM */}
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
                placeholder="Short title"
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

          <Checkbox
            label="Post anonymously"
            checked={anonymous}
            onChange={(_, c) => setAnonymous(!!c)}
          />

          {/* ⭐⭐⭐ FIX — BUTTONS IN THEIR OWN BLOCK ⭐⭐⭐ */}
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginTop: 12 }}>
            <PrimaryButton text="Test Console Log" onClick={() => console.log("🔥 BUTTON 1 CLICKED")} />

            <PrimaryButton
              text="Test Alert"
              onClick={() => {
                console.log("🔥 BUTTON 2 CLICKED");
                alert("🔥 Alert test successful");
              }}
            />

            <PrimaryButton
              text="Test submit()"
              disabled={!summary.trim() || !description.trim()}
              onClick={() => {
                console.log("🔥 BUTTON 3 CLICKED");
                submit();
              }}
            />

            <PrimaryButton
              text="Force POST"
              onClick={async () => {
                console.log("🔥 BUTTON 4 CLICKED");
                await fetch(`${API_BASE}/HttpTrigger1?category=suggestions`, {
                  method: "POST",
                  headers: { "Content-Type": "application/json" },
                  body: JSON.stringify({
                    category: "Suggestions",
                    title: "ForcePost",
                    summary: "Force post summary",
                    description: "Manual POST test",
                    type: "Force",
                    owner: "ForceTester",
                  }),
                });
              }}
            />
          </div>
        </div>
      </div>

      {/* LIST */}
      <div className="notes-card" style={{ marginTop: 20 }}>
        <Stack horizontal horizontalAlign="space-between">
          <Text className="section-title">Community suggestions</Text>
          <span style={{ color: "#a6b7c6", fontSize: 12 }}>{sorted.length} total</span>
        </Stack>

        {loading && <div className="note-empty">Loading…</div>}
        {!loading && sorted.length === 0 && <div className="note-empty">No suggestions yet.</div>}

        {sorted.length > 0 && (
          <div className="notes-list">
            {sorted.map(s => {
              const open = expanded === s.id;

              return (
                <div key={s.id} className="note-item">
                  <div className="note-header">
                    <div className="note-meta" style={{ display: "flex", gap: 8 }}>
                      <span className="wf-inprogress-badge" style={{ padding: "2px 8px" }}>{s.type}</span>
                      <Text className="note-alias">{s.summary}</Text>
                      <span className="note-time">{new Date(s.ts).toLocaleString()}</span>
                    </div>

                    <button onClick={() => setExpanded(open ? null : s.id)}>
                      {open ? "Hide" : "Show"}
                    </button>
                  </div>

                  {open && <div className="note-body">{s.description}</div>}
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
