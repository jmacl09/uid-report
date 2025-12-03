import React, { useEffect, useMemo, useState } from "react";
import { API_BASE } from "../api/config";
import { getStatusForUid } from "../api/items";

type Status = "Not Started" | "In progress" | "Completed";

interface Props {
  uid: string | null;
  data: any | null;
  style?: React.CSSProperties;
  bare?: boolean;
}

type PersistShape = {
  configPush: Status;
  circuitsQc: Status;
  expectedDeliveryDate: string | null;
};

const defaultState: PersistShape = {
  configPush: "Not Started",
  circuitsQc: "Not Started",
  expectedDeliveryDate: null,
};

const STATUS_OPTS: Status[] = ["Not Started", "In progress", "Completed"];

const UIDStatusPanel: React.FC<Props> = ({ uid, data, style, bare }) => {

  const [state, setState] = useState<PersistShape>(defaultState);
  const storageKey = uid ? `uidStatus:${uid}` : null;

  // -----------------------------
  // Auto-detected rules (unchanged)
  // -----------------------------
  const autos = useMemo(() => {
    const wfRaw = String(data?.KQLData?.WorkflowStatus ?? "").trim();
    const wfFinished = /wffinished|wf finished|finished/i.test(wfRaw);

    const tickets: any[] = Array.isArray(data?.ReleatedTickets)
      ? data.ReleatedTickets
      : Array.isArray(data?.GDCOTickets)
      ? data.GDCOTickets
      : Array.isArray(data?.AssociatedTickets)
      ? data.AssociatedTickets
      : Array.isArray(data?.AssociatedUIDs)
      ? data.AssociatedUIDs.filter((r: any) => r && (r.TicketId || r.CleanTitle || r.TicketLink || r.TicketID))
      : [];

    const normalize = (s: string) => String(s ?? "").replace(/\s+/g, " ").trim();
    const normKey = (k: string) => String(k ?? "").toLowerCase().replace(/[\s:]+/g, "");

    const circuitsResolved = tickets.some((t) => {
      if (!t || typeof t !== "object") return false;

      let titleRaw = "";
      let stateRaw = "";

      for (const k of Object.keys(t)) {
        const nk = normKey(k);
        if (!titleRaw && (nk.includes("title") || nk.includes("task") || nk.includes("name") || nk.includes("description"))) {
          titleRaw = (t as any)[k];
        }
        if (!stateRaw && (nk.includes("state") || nk.includes("status"))) {
          stateRaw = (t as any)[k];
        }
      }

      const title = normalize(titleRaw);
      const state = normalize(stateRaw);

      return /^resolved$/i.test(state) && /\bcircuit\s*[^\w\s]*\s*qc\b/i.test(title);
    });

    return { wfFinished, circuitsResolved };
  }, [data]);

  // -----------------------------
  // Load local + server status
  // -----------------------------
  useEffect(() => {
    const loadStatus = async () => {
      if (!uid || !storageKey) return;

      let local = defaultState;

      try {
        const raw = localStorage.getItem(storageKey);
        if (raw) local = { ...local, ...JSON.parse(raw) };
      } catch {}

      try {
        const items = await getStatusForUid(uid, "HttpTrigger1");
        if (items && items.length) {
          // Prefer most recent
          const sorted = items.slice().sort((a: any, b: any) => {
            const ta = Date.parse(a?.savedAt || a?.timestamp || a?.Timestamp || "") || 0;
            const tb = Date.parse(b?.savedAt || b?.timestamp || b?.Timestamp || "") || 0;
            return tb - ta;
          });
          const s = sorted[0];

          // s may have direct fields or a JSON description
          let parsedDesc: any = null;
          try { parsedDesc = s.description ? JSON.parse(s.description) : null; } catch {}

          local = {
            configPush: s.configPush ?? parsedDesc?.configPush ?? local.configPush,
            circuitsQc: s.circuitsQc ?? parsedDesc?.circuitsQc ?? local.circuitsQc,
            expectedDeliveryDate: s.expectedDeliveryDate ?? parsedDesc?.expectedDeliveryDate ?? local.expectedDeliveryDate,
          };
        }
      } catch {}

      // Apply automatic rules
      if (autos.wfFinished) local.configPush = "Completed";
      if (autos.circuitsResolved) {
        local.circuitsQc = "Completed";
        local.configPush = "Completed";
      }

      setState(local);
    };

    void loadStatus();
  }, [uid, storageKey, autos.wfFinished, autos.circuitsResolved]);

  // -----------------------------
  // Persist to localStorage
  // -----------------------------
  useEffect(() => {
    if (!storageKey) return;
    try { localStorage.setItem(storageKey, JSON.stringify(state)); } catch {}
  }, [state, storageKey]);

  // -----------------------------
  // Persist to API (debounced)
  // -----------------------------
  useEffect(() => {
    if (!uid) return;

    const timer = setTimeout(async () => {
      try {
        // POST to the function endpoint `HttpTrigger1` with category 'Status'.
        // The function expects `title` for generic category saves, so include a title and
        // store the status payload in `description` as JSON for later parsing.
        const payload = {
          category: "Status",
          uid,
          title: "Status",
          description: JSON.stringify({ configPush: state.configPush, circuitsQc: state.circuitsQc, expectedDeliveryDate: state.expectedDeliveryDate }),
        };

        await fetch(`${API_BASE}/HttpTrigger1`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload),
          credentials: "include",
        });
      } catch (e) {
        console.warn("[UIDStatusPanel] Failed to save status:", e);
      }
    }, 900);

    return () => clearTimeout(timer);
  }, [state, uid]);

  // -----------------------------
  // Helpers
  // -----------------------------
  const setField = (k: keyof PersistShape, v: any) =>
    setState((prev) => ({ ...prev, [k]: v }));

  const colorFor = (s: Status) => {
    const isLight = document.documentElement.classList.contains("light-theme") || document.body.classList.contains("light-theme");
    switch (s) {
      case "Completed": return { accent: "#00cc55" };
      case "In progress": return isLight ? { accent: "#f59e0b" } : { accent: "#50b3ff" };
      default: return { accent: "#9aa0a6" };
    }
  };

  const row = (label: string, value: Status, onChange: (s: Status) => void) => (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 6 }}>
      <div className="ai-summary-label">{label}</div>
      <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
        <span
          style={{
            width: 7,
            height: 7,
            borderRadius: 999,
            background: colorFor(value).accent,
            boxShadow: `0 0 6px ${colorFor(value).accent}`,
          }}
        />
        <select
          className="sleek-select"
          value={value}
          onChange={(e) => onChange(e.target.value as Status)}
          style={{
            height: 22, padding: "2px 6px", fontSize: 11, borderRadius: 6,
            borderColor: colorFor(value).accent, borderWidth: 1, borderStyle: "solid"
          }}
        >
          {STATUS_OPTS.map((o) => <option key={o}>{o}</option>)}
        </select>
      </div>
    </div>
  );

  // -----------------------------
  // Render
  // -----------------------------
  const content = (
    <>
      <div className="ai-summary-header">Status Tracker</div>
      <div className="ai-summary-body" style={{ gap: 6, fontSize: 11, padding: "8px 10px" }}>

        {row("Config Push", state.configPush, (s) => setField("configPush", s))}
        {row("Circuits QC", state.circuitsQc, (s) => setField("circuitsQc", s))}

        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 }}>
          <div className="ai-summary-label">Expected Delivery</div>
          <input
            type="date"
            className="ai-summary-date sleek-select"
            value={state.expectedDeliveryDate || ""}
            onChange={(e) => setField("expectedDeliveryDate", e.target.value)}
            style={{
              height: 22, padding: "0 6px", fontSize: 11, borderRadius: 6, minWidth: 100, maxWidth: 140
            }}
          />
        </div>

      </div>
    </>
  );

  if (bare) {
    return <div className="ai-summary-subpanel" style={style}>{content}</div>;
  }

  return <div className="ai-summary-panel" style={style}>{content}</div>;
};

export default UIDStatusPanel;
