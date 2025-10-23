import React from "react";
import { BrowserRouter as Router, Routes, Route, useNavigate, useLocation } from "react-router-dom";
import { Nav } from "@fluentui/react";
import Dashboard from "./pages/Dashboard";
import UIDLookup from "./pages/UIDLookup";
import VSOAssistant from "./pages/VSOAssistant";
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
      { name: "UID Lookup", key: "uidLookup", icon: "Search", url: "/uid" },
      { name: "VSO Assistant", key: "vsoAssistant", icon: "Robot", url: "/vso" },
      { name: "Fiber Spans", key: "fiberSpans", icon: "NetworkTower", url: "#" },
      { name: "Device Lookup", key: "deviceLookup", icon: "DeviceBug", url: "#" },
      { name: "Reports", key: "reports", icon: "BarChartVertical", url: "#" },
      { name: "Settings", key: "settings", icon: "Settings", url: "#" },
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
            : location.pathname === "/uid"
            ? "uidLookup"
            : location.pathname === "/vso"
            ? "vsoAssistant"
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
          <div style={{ position: 'relative' }}>
            <div style={{ position: 'absolute', right: 28, top: 12 }}>
              <UserStatus />
            </div>
          </div>
          <Routes>
            <Route path="/" element={<Dashboard />} />
            <Route path="/uid" element={<UIDLookup />} />
            <Route path="/vso" element={<VSOAssistant />} />
          </Routes>
        </div>
      </div>
    </Router>
  );
}

export default App;
