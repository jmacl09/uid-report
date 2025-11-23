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
  // Removed unused spansData state (was only ever set, never read)
  const [fullSpansData, setFullSpansData] = useState<any[] | null>(null); // raw full window from server
  const [displayedSpansData, setDisplayedSpansData] = useState<any[] | null>(null); // filtered by timeframe
  const [routes, setRoutes] = useState<Array<{route: string; solutionId?: string}> | null>(null);
  const [routesExpanded, setRoutesExpanded] = useState<boolean>(false);
  const [selectedSolutionId, setSelectedSolutionId] = useState<string | null>(null);
  const [selectedRoute, setSelectedRoute] = useState<string | null>(null);
  const COLORS = ["#60A5FA", "#34D399", "#F59E0B", "#F472B6", "#A78BFA", "#FB7185", "#60C8E8", "#FCD34D"];
  const [maxVal, setMaxVal] = useState<number | null>(null);
  const [minVal, setMinVal] = useState<number | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [submitted, setSubmitted] = useState<boolean>(false);
  const [progressVisible, setProgressVisible] = useState<boolean>(false);
  const [progressComplete, setProgressComplete] = useState<boolean>(false);
  const [timeframeDays, setTimeframeDays] = useState<number>(7); // 7D default

  const handleSubmit = async (spanValue?: string, days?: number) => {
    setSubmitted(true);
    setError(null);
    // spansData removed
    setRoutes(null);
    setRoutesExpanded(false);
    setSelectedSolutionId(null);
    setSelectedRoute(null);
    const target = (spanValue ?? span) || "";
    // Validate number of spans (comma-separated)
    const parts = target.split(',').map((s) => (s || '').trim()).filter(Boolean);
    if (parts.length > 20) {
      setError('Maximum 20 spans allowed per request. Please reduce your selection.');
      setProgressComplete(true);
      setLoading(false);
      return;
    }
    if (!target || !target.trim()) {
      setError("Span is required.");
      return;
    }
    // show themed progress bar similar to UIDLookup
    setProgressVisible(true);
    setProgressComplete(false);
    setLoading(true);
    try {
      // Always request the maximum window (7 days) from server so we can client-side filter.
      // If server supports returning exactly requested days, we still ask for 7 by default
      // and rely on server when days arg is provided. For now request 7 days to populate full data.
      const resp = await getSpanUtilization(target.trim(), 7);
      const raw = resp?.Spans;
      const arr = Array.isArray(raw) ? raw : (raw ? [raw] : []);
      setFullSpansData(arr as any[]);
      // compute displayed subset according to current timeframeDays
      setDisplayedSpansData(filterByDays(arr as any[], timeframeDays));
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

      // Collect route values but only one SPAN DC per SolutionID (use latest timestamp per SolutionID)
      const solutionMap = new Map<string, { route: string; ts: number }>();
      const routeMap = new Map<string, number>(); // fallback when SolutionID not present
      const pushRouteFallback = (route: any, tsMs: number | null) => {
        if (route === undefined || route === null) return;
        try {
          const s = String(route).trim();
          if (!s) return;
          if (s.toUpperCase() === 'UNKNOWN') return;
          const existing = routeMap.get(s);
          const t = tsMs ?? 0;
          if (existing === undefined || t > existing) routeMap.set(s, t);
        } catch { }
      };

      // helper to extract timestamp (ms) from an object that may contain TIMESTAMP/timestamp/Time
      const extractTsMs = (o: any): number | null => {
        if (!o) return null;
        const t = o?.TIMESTAMP ?? o?.timestamp ?? o?.time ?? o?.Time ?? null;
        if (!t) return null;
        const dt = new Date(t as any);
        const ms = isNaN(dt.getTime()) ? null : dt.getTime();
        return ms;
      };

      // Process span datapoints: prefer grouping by SolutionID
      if (arr && arr.length > 0) {
        for (const sObj of arr) {
          const ts = extractTsMs(sObj) ?? 0;
          const sol = sObj?.SolutionID ?? sObj?.solutionId ?? sObj?.SolutionId ?? null;
          const route = sObj?.SpanDC ?? sObj?.DataCenter ?? null;
          if (sol) {
            const key = String(sol).trim();
            if (key) {
              const existing = solutionMap.get(key);
              if (!existing || ts > existing.ts) {
                const r = route && String(route).trim() && String(route).trim().toUpperCase() !== 'UNKNOWN' ? String(route).trim() : '';
                solutionMap.set(key, { route: r, ts });
              }
            }
          } else {
            // no SolutionID: fall back to route-level dedupe
            pushRouteFallback(route, ts);
          }
        }
      }

      // Include top-level/body/value entries: if they have SolutionID use it, else fallback
      const addEntry = (entry: any) => {
        if (!entry) return;
        const ts = extractTsMs(entry) ?? 0;
        const sol = entry?.SolutionID ?? entry?.solutionId ?? entry?.SolutionId ?? null;
        const route = entry?.SpanDC ?? entry?.DataCenter ?? null;
        if (sol) {
          const key = String(sol).trim();
          if (key) {
            const existing = solutionMap.get(key);
            if (!existing || ts > existing.ts) {
              const r = route && String(route).trim() && String(route).trim().toUpperCase() !== 'UNKNOWN' ? String(route).trim() : '';
              solutionMap.set(key, { route: r, ts });
            }
          }
        } else {
          pushRouteFallback(route, ts);
        }
      };

      addEntry(resp);
      const body = tryParse((resp as any)?.body || (resp as any)?.Body);
      if (body) addEntry(body);
      const val = Array.isArray(resp) ? resp : tryParse((resp as any)?.value) || null;
      if (val && Array.isArray(val) && val.length > 0) {
        for (const it of val) addEntry(it);
      }

      // Build final array: prefer per-SolutionID entries (one per SolutionID), sorted by latest ts desc
      let finalRoutes: Array<{ route: string; solutionId?: string }> = [];
      if (solutionMap.size > 0) {
        finalRoutes = Array.from(solutionMap.entries())
          .sort((a, b) => (b[1].ts - a[1].ts))
          .map(([sid, v]) => ({ route: v.route, solutionId: sid }))
          .filter((r) => r.route && r.route.length > 0);
      }

      // If no SolutionID-based routes found, fall back to route-level dedupe
      if (finalRoutes.length === 0 && routeMap.size > 0) {
        finalRoutes = Array.from(routeMap.entries()).sort((a, b) => (b[1] - a[1])).map(([route]) => ({ route }));
      }

      if (finalRoutes.length > 0) setRoutes(finalRoutes);
    } catch (e: any) {
      setError(String(e?.message || e));
    } finally {
      // signal progress completion; ThemedProgressBar will call onDone which hides it
      setProgressComplete(true);
      setLoading(false);
    }
  };

  // Filter helper: return points whose TIMESTAMP is within last `days` days
  const filterByDays = (data: any[] | null, days: number) => {
    try {
      if (!data || !data.length) return [];
      const now = Date.now();
      const cutoff = now - days * 24 * 60 * 60 * 1000;
      return data.filter((p) => {
        try {
          const t = p?.TIMESTAMP ?? p?.timestamp ?? p?.time ?? p?.Time ?? null;
          const dt = new Date(t as any);
          const ms = isNaN(dt.getTime()) ? null : dt.getTime();
          if (ms === null) return false;
          return ms >= cutoff;
        } catch { return false; }
      });
    } catch { return data || []; }
  };

  // Recompute max/min whenever the displayed dataset changes so chart stats update
  useEffect(() => {
    try {
      const arr = displayedSpansData || [];
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
  }, [displayedSpansData]);

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

  // Derive solution order and color map based on displayed data so colors match the chart
  const deriveSolutionOrderAndColors = (data: any[] | null) => {
    const order: string[] = [];
    const seen = new Set<string>();
    if (!data) return { order, colorMap: {} as Record<string,string> };
    for (const p of data) {
      const sid = (p?.SolutionId || p?.SolutionID || p?.solutionId);
      if (!sid) continue;
      const k = String(sid);
      if (!seen.has(k)) {
        seen.add(k);
        order.push(k);
      }
    }
    const colorMap: Record<string,string> = {};
    for (let i = 0; i < order.length; i++) {
      colorMap[order[i]] = COLORS[i % COLORS.length];
    }
    return { order, colorMap };
  };

  const { order: solutionOrder, colorMap } = deriveSolutionOrderAndColors(displayedSpansData);

  // Filter displayed data for chart based on selectedSolutionId or selectedRoute
  const filteredForChart = (displayedSpansData || []).filter((p) => {
    if (selectedSolutionId) {
      const sid = (p?.SolutionId || p?.SolutionID || p?.solutionId);
      return String(sid) === selectedSolutionId;
    }
    if (selectedRoute) {
      const route = p?.SpanDC ?? p?.DataCenter ?? null;
      return route && String(route) === selectedRoute;
    }
    return true;
  });

  return (
    <div className="main-content fade-in">
      <div className="vso-form-container glow fiber-util-container" style={{ width: "92%", maxWidth: 1100 }}>
        <div className="banner-title">
          <span className="title-text">Fiber Span Utilization</span>
          <span className="title-sub">Span Utilization – Traffic Analysis</span>
          <div style={{ marginTop: 6 }} />
        </div>

        <div style={{ display: "flex", gap: 12, alignItems: "center", marginTop: 8 }} className="fiber-util-stats">
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
                {(() => {
                  const parts = (span || "").split(',').map((s) => (s || '').trim()).filter(Boolean);
                  if (parts.length === 0) return 'Tip: You can search multiple spans by entering comma-separated Span IDs.';
                  if (parts.length <= 20) return `Spans entered: ${parts.length} (max 20)`;
                  return `Spans entered: ${parts.length} — maximum is 20; remove ${parts.length - 20} to continue.`;
                })()}
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
        {!loading && submitted && displayedSpansData && displayedSpansData.length === 0 && (
          <MessageBar messageBarType={MessageBarType.info} isMultiline={false} styles={{ root: { marginTop: 12 } }}>
            No utilization data found for this span.
          </MessageBar>
        )}

        {!loading && displayedSpansData && displayedSpansData.length > 0 && (
          <div style={{ marginTop: 18 }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
              <div>
                {routes && routes.length > 0 && (
                  <div>
                    <Text styles={{ root: { color: "#dfefff", fontWeight: 700, fontSize: 15, marginBottom: 6 } }}>{routes.length === 1 ? 'Route:' : 'Routes:'}</Text>
                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, maxWidth: 760, alignItems: 'center' }}>
                      {(() => {
                        const maxShown = routesExpanded ? routes.length : 3;
                        const shown = routes.slice(0, maxShown);
                        return (
                          <>
                            {shown.map((r, idx) => (
                              <button
                                key={idx}
                                onClick={() => {
                                  // toggle selection by solutionId if available, otherwise by route
                                  if (r.solutionId) {
                                    setSelectedSolutionId(prev => prev === r.solutionId ? null : r.solutionId || null);
                                    setSelectedRoute(null);
                                  } else {
                                    setSelectedRoute(prev => prev === r.route ? null : r.route || null);
                                    setSelectedSolutionId(null);
                                  }
                                }}
                                style={{
                                  background: r.solutionId ? (colorMap[r.solutionId || ''] || '#0b2b34') : '#0b2b34',
                                  color: r.solutionId ? '#021216' : '#dff6fb',
                                  padding: '6px 10px',
                                  borderRadius: 12,
                                  fontSize: 12,
                                  fontWeight: 600,
                                  whiteSpace: 'nowrap',
                                  border: selectedSolutionId === r.solutionId || selectedRoute === r.route ? '2px solid #ffffff' : '1px solid #274648',
                                  cursor: 'pointer'
                                }}
                              >{formatRoute(r.route)}</button>
                            ))}
                            {routes.length > 3 && !routesExpanded && (
                              <button
                                onClick={() => setRoutesExpanded(true)}
                                style={{
                                  background: 'transparent',
                                  border: '1px dashed #274648',
                                  color: '#9fb3c6',
                                  padding: '6px 10px',
                                  borderRadius: 12,
                                  cursor: 'pointer',
                                  fontSize: 12,
                                  fontWeight: 600
                                }}
                                aria-label={`Show ${routes.length - 3} more routes`}
                              >{`+${routes.length - 3} more`}</button>
                            )}
                            {routesExpanded && routes.length > 3 && (
                              <button
                                onClick={() => setRoutesExpanded(false)}
                                style={{
                                  background: 'transparent',
                                  border: '1px dashed #274648',
                                  color: '#9fb3c6',
                                  padding: '6px 10px',
                                  borderRadius: 12,
                                  cursor: 'pointer',
                                  fontSize: 12,
                                  fontWeight: 600
                                }}
                                aria-label="Collapse routes list"
                              >Collapse</button>
                            )}
                          </>
                        );
                      })()}
                    </div>
                  </div>
                )}
              </div>
              <div style={{ textAlign: 'right', color: '#cfe3ff' }}>
                <div style={{ display: 'flex', gap: 12, alignItems: 'baseline', justifyContent: 'flex-end' }}>
                  <div style={{ display: 'inline-flex', gap: 6, alignItems: 'center', marginRight: 6 }}>
                    {[1,3,7].map((d) => (
                      <button
                        key={d}
                        onClick={() => {
                          setTimeframeDays(d);
                          if (fullSpansData && fullSpansData.length) {
                            setDisplayedSpansData(filterByDays(fullSpansData, d));
                          }
                        }}
                        style={{
                          background: timeframeDays === d ? '#0b4856' : 'transparent',
                          border: '1px solid #274648',
                          color: timeframeDays === d ? '#dff6fb' : '#9fb3c6',
                          padding: '6px 8px',
                          borderRadius: 6,
                          cursor: 'pointer',
                          fontSize: 12,
                          fontWeight: 600,
                        }}
                      >{`${d}D`}</button>
                    ))}
                  </div>
                  <div style={{ display: 'flex', gap: 12, alignItems: 'baseline', justifyContent: 'flex-end' }}>
                    <div className="fiber-util-stat-card">
                      <div className="fiber-util-stat-label">MAX</div>
                      <div className="fiber-util-stat-value">{maxVal !== null && !isNaN(maxVal) ? `${Number(maxVal).toFixed(2)} Gbps` : '—'}</div>
                    </div>
                    <div className="fiber-util-stat-card">
                      <div className="fiber-util-stat-label">MIN</div>
                      <div className="fiber-util-stat-value">{minVal !== null && !isNaN(minVal) ? `${Number(minVal).toFixed(2)} Gbps` : '—'}</div>
                    </div>
                  </div>
                </div>
              </div>
            </div>

            <div className="fiber-util-chart" style={{ padding: 12 }}>
              <TrafficChart data={filteredForChart} colorMap={colorMap} />
            </div>

            {solutionOrder && solutionOrder.length > 0 && (
              <div style={{ marginTop: 10, display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'center' }}>
                {solutionOrder.map((sid) => (
                  <button
                    key={sid}
                    onClick={() => { setSelectedSolutionId(prev => prev === sid ? null : sid); setSelectedRoute(null); }}
                    style={{
                      background: colorMap[sid] || '#0b2b34',
                      color: '#021216',
                      padding: '6px 10px',
                      borderRadius: 12,
                      fontSize: 12,
                      fontWeight: 700,
                      whiteSpace: 'nowrap',
                      border: selectedSolutionId === sid ? '2px solid #fff' : '1px solid rgba(255,255,255,0.08)',
                      cursor: 'pointer'
                    }}
                    aria-pressed={selectedSolutionId === sid}
                    title={sid}
                  >{sid}</button>
                ))}
              </div>
            )}

            <div style={{ marginTop: 12 }} className="fiber-util-disclaimer">
              <MessageBar messageBarType={MessageBarType.warning}>
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
