import React from "react";
import { BrowserRouter as Router, Routes, Route, useNavigate, useLocation } from "react-router-dom";
import { Nav } from "@fluentui/react";
import Dashboard from "./pages/Dashboard";
import UIDLookup from "./pages/UIDLookup";
import VSOAssistant from "./pages/VSOAssistant";
import logo from "./assets/optical360-logo.png";
import "./Theme.css";

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
