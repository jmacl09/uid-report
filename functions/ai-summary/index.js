/*
  AI Summary endpoint
  - If Azure OpenAI env is present and fetch succeeds, use it to generate a chat-like summary
  - Otherwise, fall back to a deterministic local summary

  Env required for AI path:
  - AZURE_OPENAI_ENDPOINT
  - AZURE_OPENAI_DEPLOYMENT
  - AZURE_OPENAI_API_KEY
  - AZURE_OPENAI_API_VERSION (e.g., 2024-10-01-preview)
*/

function toStringSafe(v) { return v == null ? '' : String(v); }

function buildLocalSummary(payload) {
  const data = payload || {};
  const links = Array.isArray(data.OLSLinks) ? data.OLSLinks : [];
  const count = links.length;

  // per-link capacity
  let perGb = null;
  const inc = data?.KQLData?.Increment;
  if (inc != null && inc !== '' && !isNaN(Number(inc))) perGb = Number(inc);
  if (perGb == null && links[0]?.APort) {
    const aPort = String(links[0].APort).toLowerCase();
    if (/four\s*hundred|400/.test(aPort)) perGb = 400;
    else if (/hundred|100/.test(aPort)) perGb = 100;
    else if (/ten\s*g|10\s*g|10g|tengig/.test(aPort)) perGb = 10;
    else {
      const m = aPort.match(/(\d+)\s*g/i);
      if (m) perGb = parseInt(m[1], 10);
    }
  }
  if (perGb == null && data?.KQLData?.DeviceA) {
    const da = String(data.KQLData.DeviceA).toLowerCase();
    if (/400/.test(da)) perGb = 400;
    else if (/100/.test(da)) perGb = 100;
    else if (/10/.test(da)) perGb = 10;
  }
  const totalGb = perGb != null ? count * perGb : null;

  // DC path: prefer Associated UIDs sites
  const assoc = Array.isArray(data?.AssociatedUIDs) ? data.AssociatedUIDs : [];
  const sitesA = Array.from(new Set(assoc.map(r => toStringSafe(r['Site A'] || r['SiteA'] || r['siteA']).trim()).filter(Boolean)));
  const sitesZ = Array.from(new Set(assoc.map(r => toStringSafe(r['Site Z'] || r['SiteZ'] || r['siteZ']).trim()).filter(Boolean)));
  const dcA = sitesA.length ? sitesA : [toStringSafe(data?.AExpansions?.DCLocation)].filter(Boolean);
  const dcZ = sitesZ.length ? sitesZ : [toStringSafe(data?.ZExpansions?.DCLocation)].filter(Boolean);

  // WF Status
  const rawWF = toStringSafe(data?.KQLData?.WorkflowStatus).trim();
  const isCancelled = /cancel|cancelled|canceled/i.test(rawWF);
  const isDecom = /decom/i.test(rawWF);
  const isFinished = /wffinished|wf finished|finished/i.test(rawWF);
  const isInProgress = /inprogress|in progress|in-progress|running/i.test(rawWF);
  const wfDisplay = isCancelled ? 'WF Cancelled' : isDecom ? 'DECOM' : isFinished ? 'WF Finished' : isInProgress ? 'WF In Progress' : (rawWF || '—');

  // Tickets: mention only in-progress (not cancelled/resolved)
  const tickets = Array.isArray(data?.GDCOTickets) ? data.GDCOTickets : [];
  const openTickets = tickets.filter((t) => {
    const state = toStringSafe(t?.State).toLowerCase();
    if (!state) return false;
    if (/(cancel|cancell)/.test(state)) return false;
    if (/(resolved|closed|complete|done)/.test(state)) return false;
    return true;
  });

  // Extra info: unique devices A/Z, top line type from MGFX
  const devA = Array.from(new Set((links || []).map(l => toStringSafe(l['A Device'] || l['ADevice']).trim()).filter(Boolean))));
  const devZ = Array.from(new Set((links || []).map(l => toStringSafe(l['Z Device'] || l['ZDevice']).trim()).filter(Boolean))));
  const mgfxA = Array.isArray(data?.MGFXA) ? data.MGFXA : [];
  const mgfxZ = Array.isArray(data?.MGFXZ) ? data.MGFXZ : [];
  const lines = [...mgfxA, ...mgfxZ].map(r => toStringSafe(r['Line'] || r['line']).trim()).filter(Boolean);
  const lineCounts = lines.reduce((m, v) => { m[v] = (m[v] || 0) + 1; return m; }, {});
  const topLine = Object.keys(lineCounts).sort((a,b) => lineCounts[b]-lineCounts[a])[0] || null;

  const capStr = totalGb != null && perGb != null ? `${count} x ${perGb}G (${totalGb}G total)` : `${count} link${count===1?'':'s'}`;
  const pathStr = `${(dcA && dcA.length) ? dcA.join(', ') : 'Unknown'} → ${(dcZ && dcZ.length) ? dcZ.join(', ') : 'Unknown'}`;

  const linesOut = [];
  linesOut.push(`Path: ${pathStr}`);
  linesOut.push(`Capacity: ${capStr}`);
  linesOut.push(`WF Status: ${wfDisplay}`);
  if (openTickets.length > 0) linesOut.push(`GDCO: ${openTickets.length} ticket${openTickets.length===1?'':'s'} in progress`);
  if (devA.length || devZ.length) linesOut.push(`Devices: A(${devA.length}) / Z(${devZ.length})`);
  if (topLine) linesOut.push(`Common Line: ${topLine}`);

  return linesOut.join('\n');
}

async function tryAzureOpenAI(input) {
  const endpoint = process.env.AZURE_OPENAI_ENDPOINT;
  const deployment = process.env.AZURE_OPENAI_DEPLOYMENT;
  const apiKey = process.env.AZURE_OPENAI_API_KEY;
  const apiVersion = process.env.AZURE_OPENAI_API_VERSION || '2024-10-01-preview';
  if (!endpoint || !deployment || !apiKey) return null;

  const sys = `You are a helpful network operations assistant. Summarize the fiber UID results.
Rules:
- Show path as "From → To" using site names if present (Associated UIDs Site A/Z); else DC location.
- Show capacity as "{count} x {perLink}G ({total}G total)" if per‑link known; else just link count.
- Show WF Status using friendly names (WF In Progress, WF Finished, WF Cancelled, DECOM) when applicable.
- Mention GDCO tickets only if there are tickets still in progress (ignore cancelled/resolved).
- Include a couple of extra insights: device counts A/Z, common line type (from MGFX) if clear.
- Keep it concise, 4–7 short lines.`;

  const user = JSON.stringify({
    OLSLinks: input?.OLSLinks || [],
    AssociatedUIDs: input?.AssociatedUIDs || [],
    GDCOTickets: input?.GDCOTickets || [],
    KQLData: input?.KQLData || {},
    AExpansions: input?.AExpansions || {},
    ZExpansions: input?.ZExpansions || {},
    MGFXA: input?.MGFXA || [],
    MGFXZ: input?.MGFXZ || []
  });

  const url = `${endpoint}/openai/deployments/${deployment}/chat/completions?api-version=${apiVersion}`;
  const payload = {
    messages: [
      { role: 'system', content: sys },
      { role: 'user', content: user }
    ],
    temperature: 0.2,
    max_tokens: 300
  };

  try {
    const res = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'api-key': apiKey
      },
      body: JSON.stringify(payload)
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const json = await res.json();
    const text = json?.choices?.[0]?.message?.content || '';
    if (!text) return null;
    return text.trim();
  } catch (e) {
    return null;
  }
}

module.exports = async function (context, req) {
  const input = req.body || {};

  // Try Azure OpenAI first (if configured)
  let answer = null;
  try {
    if (typeof fetch !== 'function') {
      // try to lazily polyfill fetch on older Node (best-effort)
      try { global.fetch = require('node-fetch'); } catch (e) {}
    }
    answer = await tryAzureOpenAI(input);
  } catch (e) {}

  // Fallback: deterministic local summary
  if (!answer) {
    answer = buildLocalSummary(input);
  }

  context.res = {
    headers: { 'Content-Type': 'application/json' },
    body: { message: answer }
  };
};
