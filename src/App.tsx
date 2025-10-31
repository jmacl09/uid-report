import React from "react";
import { BrowserRouter as Router, Routes, Route, useNavigate, useLocation } from "react-router-dom";
import { Nav } from "@fluentui/react";
import Dashboard from "./pages/Dashboard";
import UIDLookup from "./pages/UIDLookup";
import VSOAssistant from "./pages/VSOAssistant";
import DCATAssistant from "./pages/DCATAssistant";
import WirecheckAutomation from "./pages/WirecheckAutomation";
import SettingsPage from "./pages/Settings";
// Inline Suggestions page to avoid module resolution issues in some environments
import { Stack, Text, TextField, PrimaryButton, Dropdown, Checkbox } from "@fluentui/react";
import logo from "./assets/optical360-logo.png";
import "./Theme.css";

// Small UI piece: show logged-in Entra email/avatar in the top-right so users can confirm who is signed in.
const UserStatus: React.FC = () => {
  const [email, setEmail] = React.useState<string>("");
  React.useEffect(() => {
    try {
      const e = localStorage.getItem("loggedInEmail") || "";
      setEmail(e);
    } catch (err) {
      setEmail("");
    }
    const handler = (ev: any) => {
      try {
        const detail = ev?.detail || localStorage.getItem('loggedInEmail') || '';
        setEmail(detail);
      } catch (e) {}
    };
    window.addEventListener('loggedInEmailChanged', handler as EventListener);
    return () => window.removeEventListener('loggedInEmailChanged', handler as EventListener);
  }, []);
  return (
    <div className="user-status">
      <div className="user-avatar" title={email || 'Not signed in'}>
        {/* Microsoft square logo simplified as inline SVG */}
        <svg width="20" height="20" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" aria-hidden>
          <rect x="1" y="1" width="10.5" height="10.5" fill="#f35325" rx="1" />
          <rect x="12.5" y="1" width="10.5" height="10.5" fill="#81bc06" rx="1" />
          <rect x="1" y="12.5" width="10.5" height="10.5" fill="#05a6f0" rx="1" />
          <rect x="12.5" y="12.5" width="10.5" height="10.5" fill="#ffb900" rx="1" />
        </svg>
      </div>
      <div className="user-email">{email || <span style={{ color: '#888' }}>Not signed in</span>}</div>
    </div>
  );
};

const navLinks = [
  {
    links: [
      { name: "Home", key: "home", icon: "Home", url: "/" },
      { name: "UID Assistant", key: "uidAssistant", icon: "Search", url: "/uid" },
      { name: "VSO Assistant", key: "vsoAssistant", icon: "Robot", url: "/vso" },
      { name: "DCAT Assistant", key: "dcatAssistant", icon: "CalculatorAddition", url: "/dcat" },
      { name: "Wirecheck Automation", key: "wirecheck", icon: "Plug", url: "/wirecheck" },
      { name: "Suggestions", key: "suggestions", icon: "Megaphone", url: "/suggestions" },
      { name: "Settings", key: "settings", icon: "Settings", url: "/settings" },
    ],
  },
];

const SidebarNav: React.FC = () => {
  const navigate = useNavigate();
  const location = useLocation();

  return (
    <div className="sidebar dark-nav">
      <img src={logo} alt="Optical 360 Logo" className="logo-img" />

      <Nav
        groups={navLinks}
        onLinkClick={(e, item) => {
          e?.preventDefault();
          if (!item?.url || item.url === "#") return;
          if (item.url === "/uid" && location.pathname === "/uid") {
            const ts = Date.now();
            navigate(`/uid?reset=${ts}`);
          } else {
            navigate(item.url);
          }
        }}
        styles={{
          root: { border: "none", background: "transparent" },
          link: {
            color: "#ddd",
            borderRadius: "6px",
            selectors: {
              ":hover": { backgroundColor: "#2a2a2a" },
            },
          },
        }}
        selectedKey={
          location.pathname === "/"
            ? "home"
            : location.pathname.startsWith("/uid")
            ? "uidAssistant"
            : location.pathname.startsWith("/vso")
            ? "vsoAssistant"
            : location.pathname.startsWith("/dcat")
            ? "dcatAssistant"
            : location.pathname.startsWith("/wirecheck")
            ? "wirecheck"
            : location.pathname.startsWith("/suggestions")
            ? "suggestions"
            : location.pathname.startsWith("/settings")
            ? "settings"
            : undefined
        }
      />

      <div
        style={{
          marginTop: "auto",
          padding: "10px",
          color: "#666",
          fontSize: "12px",
          textAlign: "center",
        }}
      >
        Built by <b>Josh Maclean</b> | Microsoft
      </div>
    </div>
  );
};

// Lightweight inline Suggestions page matching the app's theme
type SuggestionItem = {
  id: string;
  ts: number;
  type: string;
  summary: string;
  description: string;
  status?: 'new' | 'inprogress' | 'completed';
  anonymous?: boolean;
  authorEmail?: string;
  authorAlias?: string;
};

const SUGGESTIONS_KEY = "uidSuggestions";

const getEmail = () => {
  try { return localStorage.getItem('loggedInEmail') || ''; } catch { return ''; }
};
const getAlias = (email?: string | null) => {
  const e = (email || '').trim();
  if (!e) return '';
  const at = e.indexOf('@');
  return at > 0 ? e.slice(0, at) : e;
};

const SuggestionsPageInline: React.FC = () => {
  const [items, setItems] = React.useState<SuggestionItem[]>(() => {
    try { const raw = localStorage.getItem(SUGGESTIONS_KEY); const arr = raw ? JSON.parse(raw) : []; return Array.isArray(arr) ? arr : []; } catch { return []; }
  });
  const [type, setType] = React.useState<string>('Improvement');
  const [summary, setSummary] = React.useState<string>('');
  const [description, setDescription] = React.useState<string>('');
  const [anonymous, setAnonymous] = React.useState<boolean>(false);

  React.useEffect(() => { try { localStorage.setItem(SUGGESTIONS_KEY, JSON.stringify(items)); } catch {} }, [items]);
  const email = getEmail();
  const alias = getAlias(email);

  const submit = () => {
    const s = summary.trim(); const d = description.trim(); if (!s || !d) return;
    const next: SuggestionItem = { id: `${Date.now()}-${Math.random().toString(36).slice(2,8)}`, ts: Date.now(), type, summary: s, description: d, status: 'new', anonymous, authorEmail: anonymous ? undefined : (email || undefined), authorAlias: anonymous ? undefined : (alias || undefined) };
    setItems([next, ...items]); setSummary(''); setDescription(''); setAnonymous(false);
  };
  const [expanded, setExpanded] = React.useState<string | null>(null);
  const sorted = React.useMemo(() => [...items].sort((a,b)=>b.ts-a.ts), [items]);

  return (
    <div style={{ maxWidth: 900, margin: '0 auto' }}>
      <div className="vso-form-container glow" style={{ width: '100%' }}>
        <div className="banner-title">
          <span className="title-text">Suggestions</span>
          <span className="title-sub">Share ideas, fixes, and improvements</span>
        </div>

        <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
          {/* Row 1: Type (dark) + Summary (full dark) */}
          <div style={{ display: 'flex', gap: 10, alignItems: 'center' }}>
            <div style={{ width: 260 }}>
              <Dropdown
                label="Type"
                options={[{ key:'Feature', text:'Feature' },{ key:'Improvement', text:'Improvement' },{ key:'Bug', text:'Bug' },{ key:'UI/UX', text:'UI/UX' },{ key:'Data', text:'Data' },{ key:'Other', text:'Other' }]}
                selectedKey={type}
                onChange={(_, opt) => setType(String(opt?.key || 'Improvement'))}
                styles={{
                  dropdown: { width: 260 },
                  title: { background: '#141414', color: '#fff', border: '1px solid #333', borderRadius: 8, height: 42 },
                  caretDown: { color: '#cfe3ff' },
                  callout: { background: '#141414', border: '1px solid #333' },
                  dropdownItems: { background: '#141414' },
                  dropdownItem: { background: '#141414', color: '#fff', selectors: { ':hover': { background: '#1a1a1a' } } },
                  dropdownItemSelected: { background: '#1f1f1f', color: '#fff' },
                  dropdownOptionText: { color: '#fff' },
                }}
              />
            </div>
            <div style={{ flex: 1 }}>
              <TextField
                label="Name / short summary"
                placeholder="e.g., Align export columns with CIS order"
                value={summary}
                onChange={(_, v)=>setSummary(v||'')}
                styles={{
                  root: { width: '100%' },
                  fieldGroup: { background: '#141414', border: '1px solid #333', borderRadius: 8, height: 42 },
                  field: { color: '#fff', selectors: { '::placeholder': { color: '#8ea6bf', opacity: 1 } } },
                }}
              />
            </div>
          </div>

          {/* Row 2: Description spanning the full suggestions box width (native textarea to avoid clipping) */}
          <div style={{ width: '100%' }}>
            <Text style={{ color: '#cfe3ff', fontWeight: 600, display: 'block', marginBottom: 4 }}>Description</Text>
            <textarea
              rows={6}
              placeholder="Describe the idea, why it helps, and any details"
              value={description}
              onChange={(e)=>setDescription(e.target.value)}
              style={{
                width: '100%',
                background: '#141414',
                color: '#fff',
                border: '1px solid #333',
                borderRadius: 8,
                padding: '10px 12px',
                lineHeight: '20px',
                boxSizing: 'border-box',
                resize: 'vertical',
                minHeight: 120,
              }}
            />
          </div>

          {/* Row 3: Anonymous checkbox (custom inline label) + Submit button */}
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              <Checkbox
                ariaLabel="Post anonymously"
                checked={anonymous}
                onChange={(_, c)=>setAnonymous(!!c)}
                boxSide="start"
                styles={{
                  root: { color: '#e6f1ff', display: 'inline-flex', alignItems: 'center', margin: 0 },
                  checkbox: { borderColor: '#5a6b7c', background: '#141414', width: 18, height: 18 },
                  checkmark: { color: '#00c853' },
                }}
              />
              <span style={{ color: '#e6f1ff', fontWeight: 600, whiteSpace: 'nowrap' }}>Post anonymously</span>
            </div>
            <PrimaryButton text="Submit suggestion" onClick={submit} disabled={!summary.trim() || !description.trim()} className="search-btn" />
          </div>
        </div>
      </div>

      <div className="notes-card" style={{ marginTop: 16 }}>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Text className="section-title">Community suggestions</Text>
          <span style={{ color: '#a6b7c6', fontSize: 12 }}>{sorted.length} total</span>
        </Stack>

        {sorted.length === 0 ? (
          <div className="note-empty">No suggestions yet. Be the first to post one.</div>
        ) : (
          <div className="notes-list">
            {sorted.map((s) => {
              const open = expanded === s.id;
              const status = s.status || 'new';
              const statusClass = status === 'completed' ? 'good' : status === 'inprogress' ? 'warning' : 'accent';
              return (
                <div key={s.id} className="note-item">
                  <div className="note-header" style={{ alignItems: 'center' }}>
                    <div className="note-meta" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                      <span className="wf-inprogress-badge" style={{ color: '#50b3ff', border: '1px solid rgba(80,179,255,0.28)', borderRadius: 8, padding: '2px 8px', fontWeight: 700, fontSize: 12 }}>{s.type}</span>
                      <span className="note-alias" style={{ color: '#e6f1ff' }}>{s.summary}</span>
                      <span className="note-dot">·</span>
                      <span className="note-time">{new Date(s.ts).toLocaleString()}</span>
                      <span className={`status-label ${statusClass}`} title={`Status: ${status.replace(/\b\w/g, c=>c.toUpperCase())}`}>{status === 'inprogress' ? 'In Progress' : status === 'completed' ? 'Completed' : 'New'}</span>
                      {!s.anonymous && (s.authorAlias || s.authorEmail) && (
                        <>
                          <span className="note-dot">·</span>
                          <span className="note-email">{s.authorAlias || s.authorEmail}</span>
                        </>
                      )}
                    </div>
                    <div className="note-controls">
                      <Dropdown
                        ariaLabel="Change status"
                        selectedKey={status}
                        options={[
                          { key: 'new', text: 'New' },
                          { key: 'inprogress', text: 'In Progress' },
                          { key: 'completed', text: 'Completed' },
                        ]}
                        styles={{
                          dropdown: { width: 160 },
                          title: { background: '#141414', color: '#fff', border: '1px solid #333', borderRadius: 8, height: 32 },
                          caretDown: { color: '#cfe3ff' },
                          callout: { background: '#141414', border: '1px solid #333' },
                          dropdownItems: { background: '#141414' },
                          dropdownItem: { background: '#141414', color: '#fff', selectors: { ':hover': { background: '#1a1a1a' } } },
                          dropdownItemSelected: { background: '#1f1f1f', color: '#fff' },
                          dropdownOptionText: { color: '#fff' },
                        }}
                        onChange={(_, opt) => {
                          const val = String(opt?.key || 'new') as 'new' | 'inprogress' | 'completed';
                          setItems(prev => prev.map(it => it.id === s.id ? { ...it, status: val } : it));
                        }}
                      />
                      <button className="note-btn" onClick={()=>setExpanded(open?null:s.id)} title={open? 'Collapse':'Expand'}>
                        {open ? 'Hide' : 'Show'}
                      </button>
                    </div>
                  </div>
                  {open && (
                    <div className="note-body">
                      <div className="note-text" style={{ whiteSpace: 'pre-wrap' }}>{s.description}</div>
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        )}
      </div>
    </div>
  );
};

function App() {
  // On app mount, try to fetch /.auth/me and persist the logged-in email for the whole app.
  React.useEffect(() => {
    const fetchAuth = async () => {
      try {
        const res = await fetch('/.auth/me', { credentials: 'include' });
        if (!res.ok) return;
        const data = await res.json();
        // Attempt same claim resolution as in VSOAssistant
        const identities = Array.isArray(data)
          ? data
          : data?.clientPrincipal
          ? [{ user_claims: data.clientPrincipal?.claims || [] }]
          : [];
        for (const id of identities) {
          const claims = id?.user_claims || [];
          const getClaim = (t: string) => claims.find((c: any) => c?.typ === t)?.val || '';
          const emailFromClaims =
            getClaim('http://schemas.xmlsoap.org/ws/2005/05/identity/claims/emailaddress') ||
            getClaim('emails') ||
            getClaim('preferred_username') ||
            getClaim('upn') ||
            '';
          // Fallback to clientPrincipal.userDetails (Static Web Apps often put the email there)
          const fallback = data?.clientPrincipal?.userDetails || '';
          const email = emailFromClaims || fallback;
          if (email) {
            try { localStorage.setItem('loggedInEmail', email); } catch (e) {}
            // notify other components
            try { window.dispatchEvent(new CustomEvent('loggedInEmailChanged', { detail: email })); } catch (e) {}
            return;
          }
        }
      } catch (e) {}
    };
    fetchAuth();
    // Re-check auth when the window regains focus or becomes visible (handles login via popup/tab)
    const tryRefresh = () => fetchAuth();
    window.addEventListener('focus', tryRefresh);
    document.addEventListener('visibilitychange', () => {
      if (document.visibilityState === 'visible') tryRefresh();
    });
    return () => {
      window.removeEventListener('focus', tryRefresh);
      document.removeEventListener('visibilitychange', () => {});
    };
  }, []);
  return (
    <Router>
      <div style={{ display: "flex", height: "100vh", backgroundColor: "#0a0a0a" }}>
        {/* Sidebar */}
        <SidebarNav />

        {/* Main Content */}
        <div
          className="main"
          style={{
            flex: 1,
            backgroundColor: "#111",
            overflowY: "auto",
            padding: "40px 30px 30px 260px",
            boxSizing: "border-box",
          }}
        >
          <div className="user-status-sticky-row">
            <UserStatus />
          </div>
          <Routes>
            <Route path="/" element={<Dashboard />} />
            <Route path="/uid" element={<UIDLookup />} />
            <Route path="/vso" element={<VSOAssistant />} />
            <Route path="/dcat" element={<DCATAssistant />} />
            <Route path="/wirecheck" element={<WirecheckAutomation />} />
            <Route path="/suggestions" element={<SuggestionsPageInline />} />
            <Route path="/settings" element={<SettingsPage />} />
          </Routes>
        </div>
      </div>
    </Router>
  );
}

export default App;
