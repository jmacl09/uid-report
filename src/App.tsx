import React from "react";
import {
  BrowserRouter as Router,
  Routes,
  Route,
  useNavigate,
  useLocation
} from "react-router-dom";
import { Nav } from "@fluentui/react";

import {
  FiberSpanUtilization,
  UIDLookup,
  VSOAssistant,
  VSOAssistantDev,
  DCATAssistant,
  WirecheckAutomation,
  Dashboard
} from "./pages";

import SettingsPage from "./pages/Settings";
import SuggestionsPage from "./pages/SuggestionsPage";   // ⭐ Correct import

import logo from "./assets/optical360-logo.png";
import "./Theme.css";
// Removed unused Fluent UI controls to satisfy eslint

/* -------------------------------------------------------------
   USER STATUS COMPONENT
------------------------------------------------------------- */
const UserStatus: React.FC = () => {
  const [email, setEmail] = React.useState<string>("");

  React.useEffect(() => {
    try {
      const e = localStorage.getItem("loggedInEmail") || "";
      setEmail(e);
    } catch {}

    const handler = (ev: any) => {
      try {
        const detail = ev?.detail || localStorage.getItem("loggedInEmail") || "";
        setEmail(detail);
      } catch {}
    };

    window.addEventListener("loggedInEmailChanged", handler as EventListener);
    return () => window.removeEventListener("loggedInEmailChanged", handler as EventListener);
  }, []);

  return (
    <div className="user-status">
      <div className="user-avatar" title={email || "Not signed in"}>
        <svg width="20" height="20" viewBox="0 0 24 24" aria-hidden>
          <rect x="1" y="1" width="10.5" height="10.5" fill="#f35325" rx="1" />
          <rect x="12.5" y="1" width="10.5" height="10.5" fill="#81bc06" rx="1" />
          <rect x="1" y="12.5" width="10.5" height="10.5" fill="#05a6f0" rx="1" />
          <rect x="12.5" y="12.5" width="10.5" height="10.5" fill="#ffb900" rx="1" />
        </svg>
      </div>
      <div className="user-email">
        {email || <span style={{ color: "#888" }}>Not signed in</span>}
      </div>
    </div>
  );
};

/* -------------------------------------------------------------
   NAVIGATION SIDEBAR
------------------------------------------------------------- */
const navLinks = [
  {
    links: [
      { name: "Home", key: "home", icon: "Home", url: "/" },
      { name: "UID Assistant", key: "uidAssistant", icon: "Search", url: "/uid" },
      { name: "VSO Assistant", key: "vsoAssistant", icon: "Robot", url: "/vso" },
      { name: "Fiber Span Utilization", key: "fiberSpanUtil", icon: "Chart", url: "/fiber-span-utilization" },
      { name: "DCAT Assistant", key: "dcatAssistant", icon: "CalculatorAddition", url: "/dcat" },
      { name: "Wirecheck Automation", key: "wirecheck", icon: "Plug", url: "/wirecheck" },
      { name: "Suggestions", key: "suggestions", icon: "Megaphone", url: "/suggestions" },
      { name: "Settings", key: "settings", icon: "Settings", url: "/settings" }
    ]
  }
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
          if (!item?.url) return;

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
            borderRadius: 6,
            selectors: { ":hover": { backgroundColor: "#2a2a2a" } }
          }
        }}
        selectedKey={
          location.pathname === "/" ? "home" :
          location.pathname.startsWith("/uid") ? "uidAssistant" :
          location.pathname.startsWith("/vso") ? "vsoAssistant" :
          location.pathname.startsWith("/fiber-span-utilization") ? "fiberSpanUtil" :
          location.pathname.startsWith("/dcat") ? "dcatAssistant" :
          location.pathname.startsWith("/wirecheck") ? "wirecheck" :
          location.pathname.startsWith("/suggestions") ? "suggestions" :
          location.pathname.startsWith("/settings") ? "settings" :
          undefined
        }
      />

      <div style={{ marginTop: "auto", padding: 10, color: "#666", fontSize: 12, textAlign: "center" }}>
        Built by <b>Josh Maclean</b> | Microsoft
      </div>
    </div>
  );
};

/* -------------------------------------------------------------
   ROOT APP COMPONENT
------------------------------------------------------------- */

function App() {
  React.useEffect(() => {
    const isLocal = window.location.hostname === "localhost";

    const fetchAuth = async () => {
      if (isLocal) return;

      try {
        const res = await fetch("/.auth/me", { credentials: "include" });
        if (!res.ok) return;

        const data = await res.json();

        const identities = Array.isArray(data)
          ? data
          : data?.clientPrincipal
          ? [{ user_claims: data.clientPrincipal?.claims || [] }]
          : [];

        for (const id of identities) {
          const claims = id?.user_claims || [];
          const getClaim = (t: string) => claims.find((c: any) => c.typ === t)?.val || "";

          const emailFromClaims =
            getClaim("http://schemas.xmlsoap.org/ws/2005/05/identity/claims/emailaddress") ||
            getClaim("emails") ||
            getClaim("preferred_username") ||
            getClaim("upn") ||
            "";

          const fallback = data?.clientPrincipal?.userDetails || "";
          const email = emailFromClaims || fallback;

          if (email) {
            try {
              localStorage.setItem("loggedInEmail", email);
              window.dispatchEvent(new CustomEvent("loggedInEmailChanged", { detail: email }));
            } catch {}
            return;
          }
        }
      } catch {}
    };

    fetchAuth();

    if (!isLocal) {
      const refresh = () => fetchAuth();
      window.addEventListener("focus", refresh);
      document.addEventListener("visibilitychange", () => {
        if (document.visibilityState === "visible") refresh();
      });

      return () => {
        window.removeEventListener("focus", refresh);
        document.removeEventListener("visibilitychange", () => {});
      };
    }
  }, []);

  return (
    <Router>
      <div style={{ display: "flex", height: "100vh", backgroundColor: "#0a0a0a" }}>
        <SidebarNav />

        <div
          className="main"
          style={{
            flex: 1,
            backgroundColor: "#111",
            overflowY: "auto",
            padding: "40px 30px 30px 260px",
            boxSizing: "border-box"
          }}
        >
          <div style={{ display: "flex", justifyContent: "center" }}>
            <div className="global-banner" role="status" aria-live="polite">
              <div className="global-banner-inner">
                <strong>Site Migration In Progress</strong>
                <span className="global-banner-text">
                  {" "}— During this transition some features may be temporarily unavailable or behave unexpectedly.
                </span>
              </div>
            </div>
          </div>

          <div className="user-status-sticky-row">
            <UserStatus />
          </div>

          {/* ⭐ FIXED ROUTES — now using your real SuggestionsPage.tsx */}
          <Routes>
            <Route path="/" element={<Dashboard />} />
            <Route path="/uid" element={<UIDLookup />} />
            <Route path="/vso" element={<VSOAssistant />} />
            <Route path="/fiber-span-utilization" element={<FiberSpanUtilization />} />
            <Route path="/vso2" element={<VSOAssistantDev />} />
            <Route path="/dcat" element={<DCATAssistant />} />
            <Route path="/wirecheck" element={<WirecheckAutomation />} />
            <Route path="/suggestions" element={<SuggestionsPage />} /> {/* ⭐ Now correct */}
            <Route path="/settings" element={<SettingsPage />} />
          </Routes>
        </div>
      </div>
    </Router>
  );
}

export default App;
