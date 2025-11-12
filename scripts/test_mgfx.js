// Quick test script to validate MGFX grouping logic (order + optical prefix fallback)
const mgfx = [
  { "StartDevice": "gvx11-349-01xomt", "StartPort": "console", "EndDevice": "gvx11-0100-0206-01c0", "EndPort": "0/3/1", "EndSku": "Cisco-4351" },
  { "StartDevice": "gvx11-349-01xomt", "StartPort": "mgmt0", "EndDevice": "gvx11-0100-0206-01m0", "EndPort": "etp48", "EndSku": "Celestica-E1031-T48S4" },
  { "StartDevice": "gvx11-349-02xomt", "StartPort": "console", "EndDevice": "gvx11-0100-0206-01c0", "EndPort": "0/3/2", "EndSku": "Cisco-4351" },
  { "StartDevice": "gvx11-349-02xomt", "StartPort": "mgmt0", "EndDevice": "gvx11-0100-0206-02m0", "EndPort": "etp8", "EndSku": "Celestica-E1031-T48S4" },
  { "StartDevice": "gvx11-349-03xomt", "StartPort": "console", "EndDevice": "gvx11-0100-0206-01c0", "EndPort": "0/3/3", "EndSku": "Cisco-4351" },
  { "StartDevice": "gvx11-349-03xomt", "StartPort": "mgmt0", "EndDevice": "gvx11-0100-0206-02m0", "EndPort": "etp2", "EndSku": "Celestica-E1031-T48S4" },
  { "StartDevice": "gvx11-349-04xomt", "StartPort": "console", "EndDevice": "gvx11-0100-0206-01c0", "EndPort": "0/3/4", "EndSku": "Cisco-4351" },
  { "StartDevice": "gvx11-349-04xomt", "StartPort": "mgmt0", "EndDevice": "gvx11-0100-0206-02m0", "EndPort": "etp3", "EndSku": "Celestica-E1031-T48S4" },
  { "StartDevice": "gvx11-349-05xomt", "StartPort": "console", "EndDevice": "gvx11-0100-0206-01c0", "EndPort": "0/3/5", "EndSku": "Cisco-4351" },
  { "StartDevice": "gvx11-349-05xomt", "StartPort": "mgmt0", "EndDevice": "gvx11-0100-0206-02m0", "EndPort": "etp4", "EndSku": "Celestica-E1031-T48S4" },
  { "StartDevice": "gvx11-349-06xomt", "StartPort": "console", "EndDevice": "gvx11-0100-0206-01c0", "EndPort": "0/3/6", "EndSku": "Cisco-4351" },
  { "StartDevice": "gvx11-349-06xomt", "StartPort": "mgmt0", "EndDevice": "gvx11-0100-0206-02m0", "EndPort": "etp5", "EndSku": "Celestica-E1031-T48S4" },
  { "StartDevice": "osl20-349-01xomt", "StartPort": "console", "EndDevice": "osl20-0100-0102-01c0", "EndPort": "0/1/21", "EndSku": "Cisco-4351" },
  { "StartDevice": "osl20-349-01xomt", "StartPort": "mgmt0", "EndDevice": "osl20-0100-0102-01m0", "EndPort": "Ethernet1/16", "EndSku": "Nexus-3048" },
  { "StartDevice": "osl20-349-02xomt", "StartPort": "console", "EndDevice": "osl20-0100-0102-01c0", "EndPort": "0/1/22", "EndSku": "Cisco-4351" },
  { "StartDevice": "osl20-349-02xomt", "StartPort": "mgmt0", "EndDevice": "osl20-0100-0102-01m0", "EndPort": "Ethernet1/17", "EndSku": "Nexus-3048" },
  { "StartDevice": "osl20-349-03xomt", "StartPort": "console", "EndDevice": "osl20-0100-0102-01c0", "EndPort": "0/1/23", "EndSku": "Cisco-4351" },
  { "StartDevice": "osl20-349-03xomt", "StartPort": "mgmt0", "EndDevice": "osl20-0100-0102-01m0", "EndPort": "Ethernet1/18", "EndSku": "Nexus-3048" },
  { "StartDevice": "osl20-349-04xomt", "StartPort": "console", "EndDevice": "osl20-0100-0102-01c0", "EndPort": "0/2/0", "EndSku": "Cisco-4351" },
  { "StartDevice": "osl20-349-04xomt", "StartPort": "mgmt0", "EndDevice": "osl20-0100-0102-01m0", "EndPort": "Ethernet1/19", "EndSku": "Nexus-3048" },
  { "StartDevice": "osl20-349-05xomt", "StartPort": "console", "EndDevice": "osl20-0100-0102-01c0", "EndPort": "0/2/1", "EndSku": "Cisco-4351" },
  { "StartDevice": "osl20-349-05xomt", "StartPort": "mgmt0", "EndDevice": "osl20-0100-0102-01m0", "EndPort": "Ethernet1/20", "EndSku": "Nexus-3048" },
  { "StartDevice": "osl20-349-06xomt", "StartPort": "console", "EndDevice": "osl20-0100-0102-01c0", "EndPort": "0/2/2", "EndSku": "Cisco-4351" },
  { "StartDevice": "osl20-349-06xomt", "StartPort": "mgmt0", "EndDevice": "osl20-0100-0102-01m0", "EndPort": "Ethernet1/21", "EndSku": "Nexus-3048" }
];

// Simulated OLSLinks to provide optical device hints
const olsLinks = [
  { 'A Device': 'gvx11-349-02xomt', 'Z Device': 'osl20-349-02xomt' }
];

function groupAndAssign(arr, links) {
  const groups = new Map();
  for (const it of arr) {
    if (!it) continue;
    const start = String(it.StartDevice ?? it.StartDeviceName ?? it['StartDevice'] ?? '').trim();
    const end = String(it.EndDevice ?? it.EndDeviceName ?? it['EndDevice'] ?? '').trim();
    if (!start || !end) continue;
    const endsWithOlt = (s) => /olt$/i.test(s);
    if (endsWithOlt(start) || endsWithOlt(end)) continue;
    const list = groups.get(start) || [];
    list.push(it);
    groups.set(start, list);
  }

  const aRows = [];
  const zRows = [];

  const xomtPrefixKey = (s) => {
    const parts = String(s || '').split('-').filter(Boolean);
    if (parts.length >= 2) return (parts[0] + '-' + parts[1]).toLowerCase();
    return parts[0]?.toLowerCase() || String(s || '').toLowerCase();
  };
  const xomtBaseKey = (s) => {
    const parts = String(s || '').split('-').filter(Boolean);
    return parts[0]?.toLowerCase() || String(s || '').toLowerCase();
  };

  let firstPrefix = null;
  let optAPrefix = null;
  let optZPrefix = null;
  try {
    const pickDevice = (r, keys) => {
      for (const k of keys) {
        const v = r?.[k];
        if (v) return String(v).trim();
      }
      return null;
    };
    if (Array.isArray(links) && links.length) {
      const aDev = pickDevice(links.find(Boolean), ['A Optical Device', 'AOpticalDevice', 'A Optical Device', 'ADevice', 'A Device', 'A Device']);
      const zDev = pickDevice(links.find(Boolean), ['Z Optical Device', 'ZOpticalDevice', 'Z Optical Device', 'ZDevice', 'Z Device', 'Z Device']);
      if (aDev) optAPrefix = xomtBaseKey(aDev);
      if (zDev) optZPrefix = xomtBaseKey(zDev);
    }
  } catch(e){}

  const makeTargetFor = (x) => {
    const cur = xomtPrefixKey(x);
    const base = xomtBaseKey(x);
    if (!firstPrefix) firstPrefix = cur;
    try {
      const allPrefixes = Array.from(groups.keys()).map(k => xomtBaseKey(String(k))).filter(Boolean);
      const distinct = Array.from(new Set(allPrefixes));
      if (distinct.length > 2 && (optAPrefix || optZPrefix)) {
        if (optAPrefix && base === optAPrefix) return aRows;
        if (optZPrefix && base === optZPrefix) return zRows;
      }
    } catch(e){}
    return cur === firstPrefix ? aRows : zRows;
  };

  for (const [xomt, items] of Array.from(groups.entries())) {
    let c0Dev = '';
    let c0Port = '';
    let c0Sku = '';
    let m0Dev = '';
    let m0Port = '';
    for (const it of items) {
      const e = (it.EndDevice || it.end || '').toLowerCase();
      const endSku = String(it.EndSku ?? it.endSku ?? '').trim();
      if (/\bc0\b|c0$/i.test(e) || /-c0/i.test(e)) {
        c0Dev = it.EndDevice || it.end || '';
        c0Port = it.EndPort || it.endPort || c0Port;
        if (endSku) c0Sku = endSku;
      } else if (/\bm0\b|m0$/i.test(e) || /-m0/i.test(e)) {
        m0Dev = it.EndDevice || it.end || '';
        m0Port = it.EndPort || it.endPort || m0Port;
      } else {
        if ((it.EndDevice || '').toLowerCase().includes('c0') && !c0Dev) { c0Dev = it.EndDevice; c0Port = it.EndPort || c0Port; if (endSku) c0Sku = endSku; }
        if ((it.EndDevice || '').toLowerCase().includes('m0') && !m0Dev) { m0Dev = it.EndDevice; m0Port = it.EndPort || m0Port; }
      }
    }
    const row = {
      XOMT: xomt,
      'C0 Device': c0Dev || '',
      'C0 Port': c0Port || '',
      'Line': '',
      'M0 Device': m0Dev || '',
      'M0 Port': m0Port || '',
      'C0 DIFF': c0Dev ? `diff:${c0Dev}` : '',
      'M0 DIFF': m0Dev ? `diff:${m0Dev}` : '',
      StartHardwareSku: c0Sku || ''
    };
    const target = makeTargetFor(xomt);
    target.push(row);
  }

  return { aRows, zRows };
}

const out = groupAndAssign(mgfx, olsLinks);
console.log('A side rows:', out.aRows.length);
console.log(out.aRows.map(r=>r.XOMT));
console.log('Z side rows:', out.zRows.length);
console.log(out.zRows.map(r=>r.XOMT));
