import React, { useState, useEffect } from "react";
import { useSearchParams } from 'react-router-dom';
import { TextField, PrimaryButton, MessageBar, MessageBarType, Text } from "@fluentui/react";
import TrafficChart from "../components/TrafficChart";
import { getSpanUtilization } from "../api/fetchLogicApp";
import ThemedProgressBar from "../components/ThemedProgressBar";
import "../Theme.css";

const FiberSpanUtilization: React.FC = () => {
  const [span, setSpan] = useState<string>("");
  const [loading, setLoading] = useState<boolean>(false);
  const [spansData, setSpansData] = useState<any[] | null>(null);
  const [dataCenter, setDataCenter] = useState<string | null>(null);
  const [maxVal, setMaxVal] = useState<number | null>(null);
  const [minVal, setMinVal] = useState<number | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [submitted, setSubmitted] = useState<boolean>(false);
  const [progressVisible, setProgressVisible] = useState<boolean>(false);
  const [progressComplete, setProgressComplete] = useState<boolean>(false);

  const handleSubmit = async (spanValue?: string) => {
    setSubmitted(true);
    setError(null);
    setSpansData(null);
    setDataCenter(null);
    const target = (spanValue ?? span) || "";
    if (!target || !target.trim()) {
      setError("Span is required.");
      return;
    }
    // show themed progress bar similar to UIDLookup
    setProgressVisible(true);
    setProgressComplete(false);
    setLoading(true);
    try {
  const resp = await getSpanUtilization(target.trim());
      const raw = resp?.Spans;
      const arr = Array.isArray(raw) ? raw : (raw ? [raw] : []);
      setSpansData(arr as any[]);
      // Compute max/min across returned points (in_gbps)
      try {
        let maxN: number | null = null;
        let minN: number | null = null;
        for (const p of arr || []) {
          const v = p?.in_gbps ?? p?.inGbps ?? p?.value ?? null;
          const num = v !== null && v !== undefined ? Number(v) : NaN;
          if (!isNaN(num)) {
            if (maxN === null || num > maxN) maxN = num;
            if (minN === null || num < minN) minN = num;
          }
        }
        setMaxVal(maxN);
        setMinVal(minN);
      } catch {
        setMaxVal(null);
        setMinVal(null);
      }

      // Helper to safely parse nested body if it's a JSON string
      const tryParse = (v: any) => {
        if (!v) return null;
        if (typeof v === "string") {
          try {
            return JSON.parse(v);
          } catch {
            return null;
          }
        }
        return v;
      };

      // Collect candidate locations where SpanDC/DataCenter might appear
      const candidates: Array<string | null> = [];
      const pushIf = (v: any) => { if (v !== undefined && v !== null) candidates.push(String(v)); };

      pushIf(resp?.SpanDC);
      pushIf(resp?.DataCenter);
      // Some Logic App responses wrap payload under `body` or `Body` or `value` or return an array
      const body = tryParse((resp as any)?.body || (resp as any)?.Body);
      if (body) {
        pushIf(body?.SpanDC);
        pushIf(body?.DataCenter);
      }
      // Check first span object if present
      if (arr && arr.length > 0) {
        pushIf(arr[0]?.SpanDC);
        pushIf(arr[0]?.DataCenter);
      }
      // If resp itself is an array or has a value array
      const val = Array.isArray(resp) ? resp : tryParse((resp as any)?.value) || null;
      if (val && Array.isArray(val) && val.length > 0) {
        pushIf(val[0]?.SpanDC);
        pushIf(val[0]?.DataCenter);
      }

      // Prefer the first candidate that is not empty and not the literal 'UNKNOWN'
      const topDc = candidates.find((c) => c && c.trim() && c.trim().toUpperCase() !== "UNKNOWN") || candidates.find((c) => c && c.trim()) || null;
      if (topDc) setDataCenter(String(topDc));
    } catch (e: any) {
      setError(String(e?.message || e));
    } finally {
      // signal progress completion; ThemedProgressBar will call onDone which hides it
      setProgressComplete(true);
      setLoading(false);
    }
  };

  // Read query param `spans` and auto-run lookup when present
  const [searchParams] = useSearchParams();
  useEffect(() => {
    try {
      const s = searchParams.get('spans') || searchParams.get('span');
      if (s && s.trim()) {
        const val = s.toString();
        // set input and trigger search
        setSpan(val.toUpperCase());
        // small timeout to ensure state updates (safe against sync issues)
        setTimeout(() => void handleSubmit(val), 50);
      }
    } catch (e) {
      // ignore
    }
    // only run when search params change
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [searchParams]);

  // Nicely format route strings from the Logic App (e.g. replace "<->" with a unicode arrow)
  const formatRoute = (r?: string | null) => {
    if (!r) return r;
    try {
      return String(r).replace(/\s*<\s*-\s*>\s*/g, " ↔ ");
    } catch {
      return String(r);
    }
  };

  return (
    <div className="main-content fade-in">
      <div className="vso-form-container glow" style={{ width: "92%", maxWidth: 1100 }}>
        <div className="banner-title">
          <span className="title-text">Fiber Span Utilization</span>
          <span className="title-sub">Span Utilization – Traffic Analysis</span>
          <div style={{ marginTop: 6 }}>
            <div style={{ display: 'inline-flex', alignItems: 'center', gap: 8, background: '#071a21', border: '1px solid #16383f', padding: '6px 10px', borderRadius: 999, boxShadow: 'inset 0 -1px 0 rgba(255,255,255,0.02)', verticalAlign: 'middle' }}>
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden>
                <rect x="3" y="6" width="18" height="14" rx="2" stroke="#2aa6d6" strokeWidth="1.2" fill="rgba(42,166,214,0.04)" />
                <path d="M8 3v4M16 3v4" stroke="#2aa6d6" strokeWidth="1.4" strokeLinecap="round" />
              </svg>
              <Text styles={{ root: { color: '#9fe0ff', fontSize: 12, fontWeight: 600 } }}>Showing last 7 days</Text>
            </div>
          </div>
        </div>

        <div style={{ display: "flex", gap: 12, alignItems: "center", marginTop: 8 }}>
          <div style={{ flex: 1 }}>
            <div style={{ marginBottom: 6, display: 'flex', alignItems: 'center', gap: 6 }}>
              <Text styles={{ root: { color: '#cfe7ff', fontWeight: 600 } }}>Spans</Text>
              <span aria-hidden style={{ color: '#ff6b6b', fontWeight: 700 }}>*</span>
            </div>
              <TextField
                placeholder="Enter Span ID or comma-separated list (e.g. ZLP94,ZLP95)"
                value={span}
                onChange={(_, v) => setSpan((v || "").toUpperCase())}
                aria-label="Span"
                onKeyDown={(e) => {
                  if (e.key === 'Enter') {
                    e.preventDefault();
                    if (!loading) void handleSubmit();
                  }
                }}
                styles={{
                  fieldGroup: { backgroundColor: "#141414", border: "1px solid #333", borderRadius: 8, height: 42 },
                  field: { color: "#fff" },
                }}
              />
            <div style={{ marginTop: 8, color: '#9fb3c6', fontSize: 12 }}>
              Tip: You can search multiple spans by entering comma-separated Span IDs.
            </div>
          </div>
          <div>
            <PrimaryButton text="Get Utilization" onClick={() => void handleSubmit()} className="search-btn" />
          </div>
        </div>

        {progressVisible && (
          <ThemedProgressBar
            active={progressVisible}
            complete={progressComplete}
            expectedMs={20000}
            label="Fetching data..."
            onDone={() => { setProgressVisible(false); setProgressComplete(false); }}
            style={{ marginTop: 12 }}
          />
        )}

        {error && (
          <MessageBar messageBarType={MessageBarType.error} isMultiline={false} styles={{ root: { marginTop: 12 } }}>
            {error}
          </MessageBar>
        )}

        {/* Default: show nothing until user submits */}
        {!loading && submitted && spansData && spansData.length === 0 && (
          <MessageBar messageBarType={MessageBarType.info} isMultiline={false} styles={{ root: { marginTop: 12 } }}>
            No utilization data found for this span.
          </MessageBar>
        )}

        {!loading && spansData && spansData.length > 0 && (
          <div style={{ marginTop: 18 }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
              <div>
                {dataCenter && (
                  <div>
                    <Text styles={{ root: { color: "#dfefff", fontWeight: 700, fontSize: 15 } }}>Route: {formatRoute(dataCenter)}</Text>
                  </div>
                )}
              </div>
              <div style={{ textAlign: 'right', color: '#cfe3ff' }}>
                <div style={{ display: 'flex', gap: 12, alignItems: 'baseline', justifyContent: 'flex-end' }}>
                  <div style={{ background: '#071821', border: '1px solid #28474f', padding: '6px 10px', borderRadius: 8 }}>
                    <div style={{ fontSize: 11, color: '#9fb3c6' }}>MAX</div>
                    <div style={{ fontSize: 14, fontWeight: 700 }}>{maxVal !== null && !isNaN(maxVal) ? `${Number(maxVal).toFixed(2)} Gbps` : '—'}</div>
                  </div>
                  <div style={{ background: '#071821', border: '1px solid #28474f', padding: '6px 10px', borderRadius: 8 }}>
                    <div style={{ fontSize: 11, color: '#9fb3c6' }}>MIN</div>
                    <div style={{ fontSize: 14, fontWeight: 700 }}>{minVal !== null && !isNaN(minVal) ? `${Number(minVal).toFixed(2)} Gbps` : '—'}</div>
                  </div>
                </div>
              </div>
            </div>

            <div style={{ background: '#071821', border: '1px solid #20343f', padding: 12, borderRadius: 8 }}>
              <TrafficChart data={spansData} />
            </div>

            <div style={{ marginTop: 12 }}>
              <MessageBar messageBarType={MessageBarType.warning} styles={{ root: { background: '#1a1a1a', border: '1px solid #333' } }}>
                <div style={{ color: '#dfefff', fontWeight: 600, marginBottom: 6 }}>Disclaimer</div>
                <div style={{ color: '#cfe3ff' }}>
                  This tool provides analytical insight into historical traffic levels for a specified fiber span. The information displayed is sourced from Microsoft Kusto datasets and represents a processed snapshot of past activity, not real-time operational data. While this tool can assist with visibility and trend analysis, it should not be relied upon as the sole source of truth. Always validate findings against official Microsoft systems, monitoring platforms, and operational records before making any engineering, planning, or operational decisions.
                </div>
              </MessageBar>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

export default FiberSpanUtilization;
