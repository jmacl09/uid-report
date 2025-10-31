import React, { useEffect, useMemo, useState } from "react";

type Status = "Unknown" | "In progress" | "Completed";

interface Props {
  uid: string | null;
  data: any | null;
  style?: React.CSSProperties;
  bare?: boolean; // render without outer panel wrapper for composition
}

type PersistShape = {
  configPush: Status;
  circuitsQc: Status;
  expectedDeliveryDate: string | null;
};

const defaultState: PersistShape = {
  configPush: "Unknown",
  circuitsQc: "Unknown",
  expectedDeliveryDate: null,
};

const STATUS_OPTS: Status[] = ["Unknown", "In progress", "Completed"];

const UIDStatusPanel: React.FC<Props> = ({ uid, data, style, bare }) => {
  const storageKey = uid ? `uidStatus:${uid}` : null;
  const [state, setState] = useState<PersistShape>(defaultState);

  // Derive autos from incoming data
  const autos = useMemo(() => {
    const wfRaw = String(data?.KQLData?.WorkflowStatus ?? "").trim();
    const wfFinished = /wffinished|wf finished|finished/i.test(wfRaw);
    const tickets: any[] = Array.isArray(data?.GDCOTickets) ? data.GDCOTickets : [];
    // Consider Circuits QC complete when we see a task "Circuit QC" (case-insensitive) in a resolved state
    // under the GDCO Tickets table. Only act on an explicit "Circuit QC" title (no broad fallbacks).
    const normalize = (s: string) => String(s ?? "").replace(/\s+/g, " ").trim();
    const normKey = (k: string) => String(k ?? "").toLowerCase().replace(/[\s:]+/g, "");

    const circuitsResolved = tickets.some((t) => {
      if (!t || typeof t !== 'object') return false;
      let titleRaw: any = "";
      let stateRaw: any = "";
      // Scan keys once and pick best candidates by normalized key
      for (const k of Object.keys(t)) {
        const nk = normKey(k);
        if (!titleRaw && (nk === 'title' || nk === 'workitemtitle' || nk === 'tickettitle' || nk === 'task' || nk === 'name' || nk === 'taskname' || nk === 'description')) {
          titleRaw = (t as any)[k];
        }
        if (!stateRaw && (nk === 'state' || nk === 'status')) {
          stateRaw = (t as any)[k];
        }
      }
      const title = normalize(String(titleRaw ?? ""));
      const state = normalize(String(stateRaw ?? ""));

      const isResolvedState = /^resolved$/i.test(state);
      const isCircuitQcTask = /\bcircuit\s*[^\w\s]*\s*qc\b/i.test(title);
      return isResolvedState && isCircuitQcTask;
    });
    return { wfFinished, circuitsResolved };
  }, [data]);

  // Temporary debug to verify field shapes at runtime (safe no-op in production)
  useEffect(() => {
    try {
      const tickets: any[] = Array.isArray(data?.GDCOTickets) ? data.GDCOTickets : [];
      if (tickets.length) {
        const sample = tickets.slice(0, 5).map((t) => {
          const kv: any = {};
          try {
            Object.keys(t || {}).forEach((k) => {
              if (/(title|work item title|ticket title|state|status)/i.test(k)) kv[k] = (t as any)[k];
            });
          } catch {}
          return kv;
        });
        // eslint-disable-next-line no-console
        console.debug("[UIDStatusPanel] GDCO sample:", sample);
      }
    } catch {}
  }, [data]);

  // Load from storage on uid change, and apply auto rules immediately to avoid race conditions
  useEffect(() => {
    if (!storageKey) return;
    try {
      const raw = localStorage.getItem(storageKey);
      const parsed = raw ? (JSON.parse(raw) as any) : defaultState;
      // Back-compat: migrate older shape { etaForDelivery: { status, date } }
      const migrated: PersistShape = {
        configPush: parsed?.configPush ?? defaultState.configPush,
        circuitsQc: parsed?.circuitsQc ?? defaultState.circuitsQc,
        expectedDeliveryDate: (parsed?.expectedDeliveryDate != null)
          ? parsed.expectedDeliveryDate
          : (parsed?.etaForDelivery?.date ?? null),
      };
      const next: PersistShape = { ...defaultState, ...migrated };
      if (autos.wfFinished) next.configPush = "Completed";
      if (autos.circuitsResolved) next.circuitsQc = "Completed";
      setState(next);
    } catch {
      const next = { ...defaultState };
      if (autos.wfFinished) next.configPush = "Completed";
      if (autos.circuitsResolved) next.circuitsQc = "Completed";
      setState(next);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [storageKey, autos.wfFinished, autos.circuitsResolved]);

  // Apply auto rules from data without overwriting explicit non-completed choices
  useEffect(() => {
    setState((prev) => {
      let next = { ...prev };
      if (autos.wfFinished) {
        next.configPush = "Completed";
      }
      if (autos.circuitsResolved) {
        next.circuitsQc = "Completed";
      }
      return next;
    });
  }, [autos.wfFinished, autos.circuitsResolved, storageKey]);

  // Persist whenever state changes for this uid
  useEffect(() => {
    if (!storageKey) return;
    try { localStorage.setItem(storageKey, JSON.stringify(state)); } catch {}
  }, [state, storageKey]);

  const setField = (k: keyof PersistShape, v: any) => setState((s) => ({ ...s, [k]: v }));

  const colorFor = (s: Status) => {
    switch (s) {
      case "Completed":
        return { accent: "#00cc55" };
      case "In progress":
        return { accent: "#50b3ff" };
      default:
        return { accent: "#9aa0a6" };
    }
  };

  const row = (label: string, value: Status, onChange: (s: Status) => void, extra?: React.ReactNode) => (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 6 }}>
      <div style={{ fontSize: 11, color: "#d2f2ff", whiteSpace: "nowrap", marginRight: 6 }}>{label}</div>
      <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
        <span
          title={value}
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
            height: 22,
            padding: "2px 6px",
            fontSize: 11,
            minWidth: 100,
            maxWidth: 140,
            background: "#222222",
            borderColor: colorFor(value).accent,
            borderWidth: 1,
            borderStyle: 'solid',
            color: "#ffffff",
            fontWeight: 600,
            borderRadius: 6,
            WebkitAppearance: 'none',
            MozAppearance: 'none',
            appearance: 'none',
          }}
        >
          {STATUS_OPTS.map((o) => (
            <option key={o} value={o}>{o}</option>
          ))}
        </select>
        {extra}
      </div>
    </div>
  );

  const content = (
    <>
      <div className="ai-summary-header">Status Tracker</div>
      <div className="ai-summary-body" style={{ gap: 6, fontSize: 11, padding: '8px 10px' }}>
        {row("Config Push", state.configPush, (s) => setField("configPush", s))}
        {row(
          "Circuits QC",
          state.circuitsQc,
          (s) => setField("circuitsQc", s),
          autos.circuitsResolved ? (
            <span title="Detected resolved 'Circuit QC' ticket" style={{ color: '#8ef3b7', fontSize: 10, fontWeight: 600 }}>auto</span>
          ) : null
        )}
        {/* Expected Delivery: date only (no status) */}
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 }}>
          <div style={{ fontSize: 11, color: "#d2f2ff", whiteSpace: "nowrap" }}>Expected Delivery</div>
          <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <input
              type="date"
              value={(state.expectedDeliveryDate || "") as string}
              onChange={(e) => setField("expectedDeliveryDate", e.target.value)}
              title="Target delivery date"
              style={{
                background: "#222222",
                color: "#ffffff",
                border: "1px solid rgba(255,255,255,0.06)",
                borderRadius: 6,
                height: 22,
                padding: "0 6px",
                fontSize: 11,
                minWidth: 100,
                maxWidth: 140,
              }}
            />
          </div>
        </div>
      </div>
    </>
  );

  if (bare) {
    return (
      <div className="ai-summary-subpanel" style={style}>
        {content}
      </div>
    );
  }

  return (
    <div className="ai-summary-panel" style={style}>
      {content}
    </div>
  );
};

export default UIDStatusPanel;
