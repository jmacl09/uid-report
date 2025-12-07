import React, { useEffect, useMemo, useState } from "react";
import {
  Stack,
  Text,
  TextField,
  PrimaryButton,
  DefaultButton,
  Dropdown,
  IDropdownOption,
  Checkbox,
  Dialog,
  DialogType,
  DialogFooter,
} from "@fluentui/react";
import { API_BASE } from "../api/config";
import { logAction } from "../api/log";

// Dark theme input/button styles
const INPUT_BG = "#0f1112";
const INPUT_BORDER = "1px solid rgba(255,255,255,0.05)";
const INPUT_TEXT = "#e6eef6";
// -------------------- TEXT FIELD STYLES --------------------
const textFieldStyles = {
  fieldGroup: {
    background: INPUT_BG,
    border: INPUT_BORDER,
    selectors: {
      ":hover": { border: "1px solid rgba(80,179,255,0.18)" },
    },
  },
  field: { 
    color: INPUT_TEXT,
    paddingTop: "10px",
    paddingBottom: "8px",
  },
  label: { color: "#9fb2c9" },
};
// -------------------- DROPDOWN STYLES --------------------
const dropdownStyles = {
  root: { width: "100%" },
  dropdown: {
    background: "transparent",
    border: "none",
  },
  title: {
    background: "linear-gradient(180deg, #101417, #0b0d0f)",
    color: INPUT_TEXT,
    borderRadius: 12,
    border: "1px solid rgba(80,179,255,0.28)",
    minHeight: 48,
    padding: "0 16px",
    display: "flex",
    alignItems: "center",
    fontWeight: "600",
    boxShadow: "0 6px 18px rgba(0,0,0,0.45)",
    selectors: {
      ":hover": { borderColor: "rgba(80,179,255,0.45)" },
      ":after": { border: "none" },
    },
  },
  caretDown: { color: "#7ac6ff" },
  label: { color: "#9fb2c9" },
  callout: {
    background: INPUT_BG,
    border: INPUT_BORDER,
  },
  dropdownItem: {
    background: INPUT_BG,
    color: INPUT_TEXT,
    selectors: {
      ":hover": { background: "#1a1d1f", color: "#d4e6ff" },
    },
  },
  dropdownItemSelected: {
    background: "#1e2427",
    color: "#50b3ff",
  },
};

// -------------------- BUTTON STYLES --------------------
const buttonStyles = {
  root: {
    background: "linear-gradient(90deg,#00b4ff,#0078d4)",
    color: "#ffffff",
    border: "1px solid rgba(80,179,255,0.35)",
    boxShadow: "0 6px 22px rgba(0,140,255,0.28), inset 0 -2px 6px rgba(255,255,255,0.04)",
    transition: "transform 120ms ease, box-shadow 120ms ease, opacity 120ms ease",
    selectors: {
      ':hover': {
        transform: 'translateY(-2px)',
        boxShadow: '0 10px 30px rgba(0,152,255,0.36), inset 0 -2px 6px rgba(255,255,255,0.06)'
      },
      ':active': {
        transform: 'translateY(0)',
        boxShadow: '0 6px 18px rgba(0,120,210,0.28)'
      },
      ':focus': {
        outline: 'none',
        boxShadow: '0 10px 30px rgba(0,152,255,0.36), 0 0 0 4px rgba(80,179,255,0.08)'
      }
    }
  },
  rootDisabled: {
    background: 'linear-gradient(90deg, rgba(255,255,255,0.02), rgba(255,255,255,0.01))',
    color: 'rgba(255,255,255,0.45)',
    border: '1px solid rgba(255,255,255,0.02)',
    opacity: 0.8
  },
};

const checkboxStyles = {
  label: { color: "#ffffff" },
};

// ---------------------------------------------------------
// Suggestion Model
// ---------------------------------------------------------
type Suggestion = {
  id: string;
  ts: number;
  type: string;
  summary: string;
  description: string;
  anonymous: boolean;
  authorEmail?: string;
  authorAlias?: string;
  status?: string;
};

// ---------------------------------------------------------
// Dropdown Options
// ---------------------------------------------------------
const typeOptions: IDropdownOption[] = [
  { key: "Feature", text: "Feature" },
  { key: "Improvement", text: "Improvement" },
  { key: "Bug", text: "Bug" },
  { key: "UI/UX", text: "UI/UX" },
  { key: "Data", text: "Data" },
  { key: "Other", text: "Other" },
];

const statusOptions: IDropdownOption[] = [
  { key: "New", text: "New" },
  { key: "Under Review", text: "Under Review" },
  { key: "Inprogress", text: "Inprogress" },
  { key: "Completed", text: "Completed" },
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

// ---------------------------------------------------------
// MAIN COMPONENT
// ---------------------------------------------------------
const SuggestionsPage: React.FC = () => {
  const [items, setItems] = useState<Suggestion[]>([]);
  const [loading, setLoading] = useState(false);

  const [type, setType] = useState("Improvement");
  const [summary, setSummary] = useState("");
  const [description, setDescription] = useState("");
  const [anonymous, setAnonymous] = useState(false);

  const email = getEmail();
  const alias = getAlias(email);

  useEffect(() => {
    const resolvedEmail = getEmail();
    logAction(resolvedEmail || "", "View Suggestions");
  }, []);

  // ---------------------------------------------------------
  // Load suggestions
  // ---------------------------------------------------------
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
          const status = (e.status || e.state || "New").toString();

          return {
            id: e.rowKey,
            ts: Date.parse(e.savedAt || new Date().toISOString()),
            type: e.type || "Other",
            summary: e.summary || e.title || "",
            description: e.description || "",
            anonymous: anon,
            authorEmail: anon ? undefined : owner,
            authorAlias: anon ? undefined : owner,
            status,
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

  // ---------------------------------------------------------
  // Submit suggestion
  // ---------------------------------------------------------
  const submit = async () => {
    const s = summary.trim();
    const d = description.trim();
    if (!s || !d) return;

    const owner = anonymous ? "Anonymous" : alias || email || "Unknown";
    const now = Date.now();

    // Optimistic UI
    setItems((prev) => [
      {
        id: `temp-${now}`,
        ts: now,
        type,
        summary: s,
        description: d,
        anonymous,
        authorEmail: anonymous ? undefined : email,
        authorAlias: anonymous ? undefined : alias,
        status: "New",
      },
      ...prev,
    ]);

    setSummary("");
    setDescription("");
    setAnonymous(false);

    logAction(email || "", "Submit Suggestion", {
      type,
      anonymous,
      owner,
    });

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

      // Reload final DB data
      const res = await fetch(`${API_BASE}/HttpTrigger1?category=suggestions`);
      const json = await res.json();
      const rows = Array.isArray(json) ? json : json.items || [];

      const mapped: Suggestion[] = rows.map((e: any) => {
        const owner = (e.owner || "").toString();
        const anon = owner.toLowerCase() === "anonymous";
        const status = (e.status || e.state || "New").toString();

        return {
          id: e.rowKey,
          ts: Date.parse(e.savedAt || new Date().toISOString()),
          type: e.type || "Other",
          summary: e.summary || e.title || "",
          description: e.description || "",
          anonymous: anon,
          authorEmail: anon ? undefined : owner,
          authorAlias: anon ? undefined : owner,
          status,
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

  // Delete dialog state
  const [deleteDialogOpen, setDeleteDialogOpen] = useState(false);
  const [deleteTarget, setDeleteTarget] = useState<{ rowKey: string; authorEmail?: string } | null>(null);

  const openDeleteDialog = (rowKey: string, authorEmail?: string) => {
    setDeleteTarget({ rowKey, authorEmail });
    setDeleteDialogOpen(true);
  };

  const closeDeleteDialog = () => {
    setDeleteDialogOpen(false);
    setDeleteTarget(null);
  };

  const confirmDelete = async () => {
    if (!deleteTarget) return;
    const { rowKey, authorEmail } = deleteTarget;
    await deleteSuggestion(rowKey, authorEmail);
    closeDeleteDialog();
  };

  // ---------------------------------------------------------
  // Delete suggestion (author or admin only)
  // ---------------------------------------------------------
  const deleteSuggestion = async (rowKey: string, authorEmail?: string) => {
    const me = email || "";
    const isAdmin = me.toLowerCase() === "joshmaclean@microsoft.com";
    const isOwner = !!(authorEmail && authorEmail.toLowerCase() === me.toLowerCase());
    if (!isAdmin && !isOwner) return;

    // Optimistic remove
    setItems((prev) => prev.filter((i) => i.id !== rowKey));

    try {
      await fetch(`${API_BASE}/HttpTrigger1?category=suggestions`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ category: "Suggestions", operation: "delete", rowKey }),
      });
      logAction(me, "Delete Suggestion", { rowKey });
    } catch (err) {
      console.error("❌ deleteSuggestion error:", err);
      // on error, reload list to restore state
      const res = await fetch(`${API_BASE}/HttpTrigger1?category=suggestions`);
      const json = await res.json();
      const rows = Array.isArray(json) ? json : json.items || [];
      const mapped: Suggestion[] = rows.map((e: any) => {
        const owner = (e.owner || "").toString();
        const anon = owner.toLowerCase() === "anonymous";
        const status = (e.status || e.state || "New").toString();

        return {
          id: e.rowKey,
          ts: Date.parse(e.savedAt || new Date().toISOString()),
          type: e.type || "Other",
          summary: e.summary || e.title || "",
          description: e.description || "",
          anonymous: anon,
          authorEmail: anon ? undefined : owner,
          authorAlias: anon ? undefined : owner,
          status,
        };
      });
      mapped.sort((a, b) => b.ts - a.ts);
      setItems(mapped);
    }
  };

  // ---------------------------------------------------------
  // Update suggestion status
  // ---------------------------------------------------------
  const updateSuggestionStatus = async (rowKey: string, newStatus: string) => {
    const me = email || "";
    const isAdmin = me.toLowerCase() === "joshmaclean@microsoft.com";
    if (!isAdmin) {
      // only admin may update status
      return;
    }

    // Optimistic UI
    setItems((prev) => prev.map((it) => (it.id === rowKey ? { ...it, status: newStatus } : it)));

    try {
      await fetch(`${API_BASE}/HttpTrigger1?category=suggestions`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ category: "Suggestions", operation: "update", rowKey, status: newStatus }),
      });
      logAction(email || "", "Update Suggestion Status", { rowKey, status: newStatus });
    } catch (err) {
      console.error("❌ updateSuggestionStatus error:", err);
    }
  };

  // ---------------------------------------------------------
  // RENDER
  // ---------------------------------------------------------
  return (
    <div style={{ maxWidth: 900, margin: "0 auto" }}>
      <div className="vso-form-container glow suggestions-form" style={{ width: "100%" }}>
        <div className="banner-title">
          <span className="title-text">Suggestions</span>
          <span className="title-sub">Share ideas, fixes, and improvements</span>
        </div>

        {/* FORM */}
        <div style={{ display: "flex", flexDirection: "column", gap: 16, marginTop: 20 }}>

          {/* TYPE + SUMMARY */}
          <div style={{ display: "flex", gap: 10, alignItems: "flex-start" }}>
            <div style={{ width: 220, display: "flex", flexDirection: "column", justifyContent: "flex-start" }}>
              <Dropdown
                label="Type"
                options={typeOptions}
                selectedKey={type}
                onChange={(_, opt) => setType(String(opt?.key || "Improvement"))}
                styles={dropdownStyles}
                className="suggestion-type-dropdown"
              />
            </div>

            <div
              style={{
                flex: 1,
                display: "flex",
                flexDirection: "column",
                justifyContent: "flex-start",
              }}
            >
              <TextField
                label="Summary"
                placeholder="Enter a brief summary of your suggestion…"
                value={summary}
                onChange={(_, v) => setSummary(v || "")}
                styles={textFieldStyles}
                className="suggestion-summary-input"
              />
            </div>
          </div>

          {/* DESCRIPTION */}
          <TextField
            label="Description"
            multiline
            rows={4}
            placeholder=""
            value={description}
            onChange={(_, v) => setDescription(v || "")}
            styles={textFieldStyles}
            className="suggestion-description-input"
          />

          {/* CHECKBOX + BUTTON */}
          <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
              marginTop: 4,
            }}
          >
            <Checkbox
              label="Post anonymously"
              checked={anonymous}
              onChange={(_, c) => setAnonymous(!!c)}
              styles={checkboxStyles}
              className="suggestion-checkbox"
              boxSide="start"
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

      {/* LIST */}
      <div className="notes-card" style={{ marginTop: 24 }}>
        <Stack horizontal horizontalAlign="space-between">
          <Text className="section-title">Community Suggestions</Text>
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
                    <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                      <span
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

                      {/* Status badge (visual for all users) */}
                      <span
                        className={`note-status-badge note-status-${(s.status || "New").toLowerCase().replace(/\s+/g, "-")}`}
                        title={`Status: ${s.status || "New"}`}
                      >
                        {s.status || "New"}
                      </span>

                      {/* Status dropdown (admin only) */}
                      {((email || "").toLowerCase() === "joshmaclean@microsoft.com") && (
                        <div style={{ width: 140, marginLeft: 8 }}>
                          <Dropdown
                            className="note-status-dropdown"
                            options={statusOptions}
                            selectedKey={s.status || "New"}
                            onChange={(_, opt) => updateSuggestionStatus(s.id, String(opt?.key || "New"))}
                            styles={{ root: { width: "100%" } } as any}
                          />
                        </div>
                      )}

                      <Text className="note-alias">{s.summary}</Text>

                      <span className="note-dot">·</span>
                      <span className="note-time">
                        {new Date(s.ts).toLocaleString()}
                      </span>

                      {!s.anonymous && (s.authorAlias || s.authorEmail) && (
                        <>
                          <span className="note-dot">·</span>
                          <span className="note-email">
                            {s.authorAlias || s.authorEmail}
                          </span>
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
                      {/* Delete button shown only to owner or admin */}
                      {(s.authorEmail?.toLowerCase() === (email || "").toLowerCase() || (email || "").toLowerCase() === "joshmaclean@microsoft.com") && (
                          <button
                            className="note-delete-btn"
                            onClick={() => openDeleteDialog(s.id, s.authorEmail)}
                          >
                            Delete
                          </button>
                        )}
                    </div>
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

      {/* Delete confirmation dialog */}
      <Dialog
        hidden={!deleteDialogOpen}
        onDismiss={closeDeleteDialog}
        dialogContentProps={{
          type: DialogType.normal,
          title: "Delete suggestion",
          subText: "Are you sure you want to delete this suggestion? This action cannot be undone.",
        }}
      >
        <DialogFooter>
          <DefaultButton onClick={closeDeleteDialog} text="Cancel" />
          <PrimaryButton onClick={confirmDelete} text="Delete" />
        </DialogFooter>
      </Dialog>
    </div>
  );
};

export default SuggestionsPage;
