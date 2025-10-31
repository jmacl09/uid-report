import React, { useMemo } from "react";

interface Props {
  data: any; // expects the same shape as viewData in UIDLookup (AExpansions, ZExpansions, KQLData, OLSLinks, AssociatedUIDs, GDCOTickets)
  style?: React.CSSProperties; // optional absolute positioning override
  bare?: boolean; // render without outer panel wrapper so it can be composed
}

const normalize = (v: any) => (v == null ? "" : String(v));

const UIDSummaryPanel: React.FC<Props> = ({ data, style, bare }) => {
  const summary = useMemo(() => {
    if (!data) return null;

    const links: any[] = Array.isArray(data?.OLSLinks) ? data.OLSLinks : [];
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

    // WF Status presentation
    const rawWF = normalize(data?.KQLData?.WorkflowStatus).trim();
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
  }, [data]);

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

  const localSummary = useMemo(() => {
    if (!summary) return '';
    const typeLine = `Type: ${summary.typeLabel}`;
    return `Path: ${dcFrom} → ${dcTo}\nCapacity: ${capStr}\nWF Status: ${summary.wfDisplay}\n${typeLine}`;
  }, [dcFrom, dcTo, capStr, summary]);

  if (!summary) return null;

  if (bare) {
    return (
      <div className="ai-summary-subpanel" style={style}>
        <div className="ai-summary-header">AI Summary</div>
        <div className="ai-summary-body">
          <div className="ai-chat">
            <div className="ai-chat-bubble fallback" style={{ whiteSpace: 'pre-wrap' }}>{localSummary}</div>
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
          <div className="ai-chat-bubble fallback" style={{ whiteSpace: 'pre-wrap', fontSize: 14 }}>{localSummary}</div>
        </div>
      </div>
    </div>
  );
};

export default UIDSummaryPanel;
