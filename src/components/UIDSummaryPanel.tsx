import React, { useMemo } from "react";

interface Props {
  data: any; // expects the same shape as viewData in UIDLookup (AExpansions, ZExpansions, KQLData, OLSLinks, AssociatedUIDs, GDCOTickets)
  currentUid?: string | null; // Prefer WF Status from AllWorkflowStatus map for this UID when available
  style?: React.CSSProperties; // optional absolute positioning override
  bare?: boolean; // render without outer panel wrapper so it can be composed
}

const normalize = (v: any) => (v == null ? "" : String(v));

const UIDSummaryPanel: React.FC<Props> = ({ data, currentUid, style, bare }) => {
  const summary = useMemo(() => {
    if (!data) return null;
    const rawLinks: any[] = Array.isArray(data?.OLSLinks) ? data.OLSLinks : [];
    // Fallback: if there are no link rows, synthesize one from KQLData DeviceA/DeviceZ so count reflects a link
    const kd = (data as any)?.KQLData || {};
    const hasKdDevices = !!(kd?.DeviceA || kd?.DeviceZ);
    const links: any[] = rawLinks.length ? rawLinks : (hasKdDevices ? [{ 'A Device': kd?.DeviceA || '', 'Z Device': kd?.DeviceZ || '' }] : []);
    const count = links.length;
    const incNum = data?.KQLData?.Increment != null && data?.KQLData?.Increment !== '' && !isNaN(Number(data?.KQLData?.Increment))
      ? Number(data?.KQLData?.Increment) : null;
    const parseSpeedFromText = (s: any): number | null => {
      const t = String(s ?? '').toLowerCase();
      if (!t) return null;
      if (/four\s*hundred|\b400\b/.test(t)) return 400;
      if (/\bhundred\b|\b100\b/.test(t)) return 100;
      if (/ten\s*g|10\s*g|\b10g\b|tengig/.test(t)) return 10;
      const m = t.match(/(\d+)\s*g/);
      if (m) return parseInt(m[1], 10);
      return null;
    };
    const perLinkGb = (row: any): number | null => {
      return (
        parseSpeedFromText(row?.APort) ??
        parseSpeedFromText(row?.ZPort) ??
        parseSpeedFromText(row?.ADevice || row?.['A Device'] || row?.DeviceA || row?.['Device A']) ??
        parseSpeedFromText(row?.ZDevice || row?.['Z Device'] || row?.DeviceZ || row?.['Device Z']) ??
        incNum
      );
    };
    const bucket: Record<string, number> = {};
    let totalGb: number | null = 0;
    for (const r of links) {
      const gb = perLinkGb(r);
      if (gb != null) {
        totalGb! += gb;
        bucket[String(gb)] = (bucket[String(gb)] || 0) + 1;
      }
    }
    if (Object.keys(bucket).length === 0) totalGb = null;

  // WF Status presentation: prefer AssociatedUID row for currentUid, then AllWorkflowStatus map, then KQLData.WorkflowStatus
  const map = (data as any)?.__WFStatusByUid as Record<string, string> | undefined;
  const assocRowsForWF: any[] = Array.isArray(data?.AssociatedUIDs) ? data.AssociatedUIDs : [];
  const assocMatchForWF = currentUid ? assocRowsForWF.find((r: any) => String(r?.UID ?? r?.Uid ?? r?.uid ?? '') === String(currentUid)) : null;
  const fromMap = currentUid && map ? map[String(currentUid)] : undefined;
  const rawWF = normalize(assocMatchForWF?.WorkflowStatus ?? assocMatchForWF?.Workflow ?? fromMap ?? data?.KQLData?.WorkflowStatus).trim();
    const isCancelled = /cancel|cancelled|canceled/i.test(rawWF);
    const isDecom = /decom/i.test(rawWF);
    const isFinished = /wffinished|wf finished|finished/i.test(rawWF);
    const isInProgress = /inprogress|in progress|in-progress|running/i.test(rawWF);
    const wfDisplay = isCancelled ? 'WF Cancelled' : isDecom ? 'DECOM' : isFinished ? 'WF Finished' : isInProgress ? 'WF In Progress' : (rawWF || '—');

    // Datacenters From/To: prefer AssociatedUIDs sites, fallback to DCLocation
    const assoc: any[] = Array.isArray(data?.AssociatedUIDs) ? data.AssociatedUIDs : [];
    const sitesA = Array.from(new Set(assoc.map(r => normalize(r['Site A'] || r['SiteA'] || r['siteA']).trim()).filter(Boolean)));
    const sitesZ = Array.from(new Set(assoc.map(r => normalize(r['Site Z'] || r['SiteZ'] || r['siteZ']).trim()).filter(Boolean)));
    const dcA = sitesA.length ? sitesA : [normalize(data?.AExpansions?.DCLocation)].filter(Boolean);
    const dcZ = sitesZ.length ? sitesZ : [normalize(data?.ZExpansions?.DCLocation)].filter(Boolean);

    // We intentionally do NOT include device counts or line type here.

    // Derive Type from Associated UIDs (Type column); default to Standard
    const typeVals = assoc
      .map((r) => String(r?.['Type'] ?? r?.['type'] ?? r?.['TYPE'] ?? '').trim())
      .filter(Boolean);
    const pick = typeVals.find((s) => /owned|hybrid/i.test(s)) || typeVals[0] || '';
    const normType = (() => {
      const t = (pick || '').replace(/_/g, '-').replace(/\s+/g, ' ').trim();
      if (!t) return 'Standard';
      if (/^owned$/i.test(t)) return 'Owned-OLS';
      if (/^hybrid$/i.test(t)) return 'Hybrid-OLS';
      // Normalize "Owned OLS" to "Owned-OLS"
      const m = t.replace(/\b(Owned|Hybrid)\b\s*[-_ ]?\s*\b(OLS)\b/i, (_m, a, _b) => `${a}-OLS`);
      return m || 'Standard';
    })();

    return { count, totalGb, wfDisplay, dcA, dcZ, bucket, typeLabel: normType } as const;
  }, [data, currentUid]);

  const capStr = (() => {
    if (!summary) return '';
    if (summary.totalGb == null) return `${summary.count} link${summary.count === 1 ? '' : 's'}`;
    // Build a distribution like "2 x 400G, 1 x 100G"
    const parts = Object.entries(summary.bucket)
      .map(([gb, cnt]) => [Number(gb), cnt] as [number, number])
      .sort((a, b) => b[0] - a[0])
      .map(([gb, cnt]) => `${cnt} x ${gb}G`);
    const dist = parts.join(', ');
    return dist ? `${dist} (${summary.totalGb}G total)` : `${summary.count} link${summary.count === 1 ? '' : 's'}`;
  })();

  const dcFrom = summary && summary.dcA.length ? summary.dcA.join(', ') : 'Unknown';
  const dcTo = summary && summary.dcZ.length ? summary.dcZ.join(', ') : 'Unknown';

  const leftSummary = useMemo(() => {
    if (!summary) return '';
    // Path, Capacity, WF Status, and Type (no explicit Links count in line)
    return `Path: ${dcFrom} → ${dcTo}\nCapacity: ${capStr}\nWF Status: ${summary.wfDisplay}\nType: ${summary.typeLabel}`;
  }, [dcFrom, dcTo, capStr, summary]);

  if (!summary) return null;

  // Prefer per-link optics/speed where available (these are merged from Base/Utilization in UIDLookup)
  const firstLink: any = Array.isArray((data as any)?.OLSLinks) && (data as any).OLSLinks.length ? (data as any).OLSLinks[0] : null;
  // Prefer associated UID row for optic/speed if it matches currentUid
  const assocRows: any[] = Array.isArray((data as any)?.AssociatedUIDs) ? (data as any).AssociatedUIDs : [];
  const assoc = currentUid ? assocRows.find(r => String(((r?.UID ?? r?.Uid ?? r?.uid) || '')) === String(currentUid)) : null;
    const extractOptic = (link: any, side: 'A' | 'Z') => {
    if (!link) return null;
    const dev = link[`${side}OpticalDevice`] ?? link[`${side} Optical Device`] ?? link[`${side}Optical`] ?? link[`${side}OptDev`] ?? link[`AOpticalDevice`];
    const port = link[`${side}OpticalPort`] ?? link[`${side} Optical Port`] ?? link[`${side}OptPort`] ?? link[`AOpticalPort`];
    const combined = (String(dev || '').trim() || '') + (port ? (port ? ` / ${String(port).trim()}` : '') : '');
    return combined || null;
  };
  const opticA = normalize((assoc?.OpticTypeA ?? assoc?.OpticA ?? extractOptic(firstLink, 'A') ?? (data as any)?.KQLData?.OpticTypeA)) || '—';
  const opticZ = normalize((assoc?.OpticTypeZ ?? assoc?.OpticZ ?? extractOptic(firstLink, 'Z') ?? (data as any)?.KQLData?.OpticTypeZ)) || '—';
  const incRaw = normalize((data as any)?.KQLData?.Increment);
    const speed = (() => {
    // Prefer per-link Speed/OpticalSpeed if present
      // 1) AssociatedUID increment/speed
      if (assoc) {
        const aS = assoc?.Increment ?? assoc?.increment ?? assoc?.OpticalSpeed ?? assoc?.Optical_Speed ?? assoc?.IncrementGb ?? assoc?.OpticalSpeedGb ?? assoc?.Speed ?? null;
        if (aS != null && String(aS).trim() !== '') {
          if (!isNaN(Number(aS))) return `${String(Number(aS)).replace(/\.0+$/, '')}G`;
          const ts = String(aS).toUpperCase();
          return /G$/.test(ts) ? ts : `${ts}`;
        }
      }
      if (firstLink) {
      const s = firstLink['Speed'] ?? firstLink['speed'] ?? firstLink['OpticalSpeed'] ?? firstLink['Optical_Speed'] ?? firstLink['OpticalSpeedGb'] ?? null;
      if (s != null && String(s).trim() !== '') {
        // If numeric (Gb), format as G
        if (!isNaN(Number(s))) return `${String(Number(s)).replace(/\.0+$/, '')}G`;
        const ts = String(s).toUpperCase();
        return /G$/.test(ts) ? ts : `${ts}`;
      }
    }
    if (!incRaw) return '—';
    const n = Number(incRaw);
    if (Number.isFinite(n)) return `${n}G`;
    return /g$/i.test(incRaw) ? incRaw.toUpperCase() : `${incRaw}G`;
  })();

  if (bare) {
    return (
      <div className="ai-summary-subpanel" style={style}>
        <div className="ai-summary-header">AI Summary</div>
        <div className="ai-summary-body">
          <div className="ai-chat">
            <div className="ai-chat-bubble fallback">
              <div
                style={{
                  display: 'grid',
                  gridTemplateColumns: 'max-content 1px 1fr',
                  columnGap: 12,
                  alignItems: 'start',
                }}
              >
                <div style={{ whiteSpace: 'pre-wrap' }}>{leftSummary}</div>
                <div style={{ width: 1, background: 'rgba(255,255,255,0.12)', height: '100%' }} />
                <div style={{ minWidth: 180, fontSize: 13 }}>
                  <div><span style={{ opacity: 0.85 }}>Router A Optic:</span> <b style={{ marginLeft: 6 }}>{opticA}</b></div>
                  <div><span style={{ opacity: 0.85 }}>Router Z Optic:</span> <b style={{ marginLeft: 6 }}>{opticZ}</b></div>
                  <div><span style={{ opacity: 0.85 }}>Speed:</span> <b style={{ marginLeft: 6 }}>{speed}</b></div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="ai-summary-panel" style={style}>
      <div className="ai-summary-header">AI Summary</div>
      <div className="ai-summary-body">
        <div className="ai-chat">
          <div className="ai-chat-bubble fallback" style={{ fontSize: 14 }}>
            <div
              style={{
                display: 'grid',
                gridTemplateColumns: 'max-content 1px 1fr',
                columnGap: 12,
                alignItems: 'start',
              }}
            >
              <div style={{ whiteSpace: 'pre-wrap' }}>{leftSummary}</div>
              <div style={{ width: 1, background: 'rgba(255,255,255,0.12)', height: '100%' }} />
              <div style={{ minWidth: 180 }}>
                <div><span style={{ opacity: 0.85 }}>Router A Optic:</span> <b style={{ marginLeft: 6 }}>{opticA}</b></div>
                <div><span style={{ opacity: 0.85 }}>Router Z Optic:</span> <b style={{ marginLeft: 6 }}>{opticZ}</b></div>
                <div><span style={{ opacity: 0.85 }}>Speed:</span> <b style={{ marginLeft: 6 }}>{speed}</b></div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default UIDSummaryPanel;
