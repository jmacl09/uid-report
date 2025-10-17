import React, { useState, useEffect } from "react";
import {
  initializeIcons,
  Stack,
  Text,
  TextField,
  PrimaryButton,
  Nav,
  Separator,
} from "@fluentui/react";

initializeIcons();

const navLinks = [
  {
    links: [
      { name: "UID Search", key: "uidSearch", icon: "Search", url: "#" },
      { name: "Fiber Spans", key: "fiberSpans", icon: "NetworkTower", url: "#" },
      { name: "Device Lookup", key: "deviceLookup", icon: "DeviceBug", url: "#" },
      { name: "Reports", key: "reports", icon: "BarChartVertical", url: "#" },
      { name: "Settings", key: "settings", icon: "Settings", url: "#" },
    ],
  },
];

export default function App() {
  const [uid, setUid] = useState<string>("");
  const [loading, setLoading] = useState<boolean>(false);
  const [user, setUser] = useState<any>(null);

  // 🔹 Fetch logged-in user info from Azure Static Web Apps
  useEffect(() => {
    const fetchUser = async () => {
      try {
        const res = await fetch("/.auth/me");
        if (res.ok) {
          const data = await res.json();
          const userInfo = data?.clientPrincipal;
          if (userInfo) setUser(userInfo);
        }
      } catch (err) {
        console.error("Error fetching user info:", err);
      }
    };
    fetchUser();
  }, []);

  // 🔹 Trigger Power Automate flow via the secure proxy route
  const handleSearch = async () => {
    if (!uid.trim()) {
      alert("Please enter a UID before searching.");
      return;
    }

    setLoading(true);

    try {
      const response = await fetch("/api/uid", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ UID: uid }),
      });

      if (response.ok) {
        alert(`✅ Flow triggered successfully for UID: ${uid}`);
      } else {
        const errorText = await response.text();
        alert(`❌ Failed to trigger flow (${response.status}): ${errorText}`);
      }
    } catch (error) {
      console.error(error);
      alert("⚠️ Network error while triggering the flow.");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{ display: "flex", height: "100vh", backgroundColor: "#f3f2f1" }}>
      {/* Sidebar */}
      <div
        style={{
          width: "260px",
          backgroundColor: "#002050",
          color: "white",
          padding: "20px",
          display: "flex",
          flexDirection: "column",
          justifyContent: "space-between",
        }}
      >
        <div>
          <Text
            variant="xLarge"
            styles={{ root: { color: "#fff", marginBottom: 20, fontWeight: 600 } }}
          >
            🔍 FiberTools
          </Text>
          <Nav
            groups={navLinks}
            styles={{
              root: {
                width: 240,
                boxSizing: "border-box",
                background: "#002050",
                color: "#ffffff",
              },
              linkText: { color: "#ffffff" },
              compositeLink: { selectors: { ":hover": { background: "#0078D4" } } },
            }}
          />
        </div>

        <div>
          <Separator styles={{ root: { borderColor: "#fff", marginTop: 20 } }} />
          {user ? (
            <Text
              variant="small"
              styles={{ root: { color: "#d0d0d0", marginTop: 10 } }}
            >
              Signed in as <strong>{user.userDetails}</strong>
            </Text>
          ) : (
            <Text
              variant="small"
              styles={{ root: { color: "#d0d0d0", marginTop: 10 } }}
            >
              Loading user info...
            </Text>
          )}
          <Text
            variant="small"
            styles={{ root: { color: "#d0d0d0", marginTop: 5 } }}
          >
            Built by Josh Maclean | Microsoft
          </Text>
        </div>
      </div>

      {/* Main Content */}
      <Stack
        tokens={{ childrenGap: 20 }}
        styles={{
          root: {
            flexGrow: 1,
            padding: "60px",
            background: "linear-gradient(135deg, #e6f0ff 0%, #ffffff 100%)",
          },
        }}
        verticalAlign="center"
        horizontalAlign="center"
      >
        <Text variant="xxLargePlus" styles={{ root: { color: "#002050" } }}>
          UID Lookup Portal
        </Text>
        <Text
          variant="mediumPlus"
          styles={{
            root: { color: "#555", textAlign: "center", maxWidth: 480, marginBottom: 20 },
          }}
        >
          Enter a UID below to retrieve network, fiber span, or optical device details.
        </Text>

        <Stack
          horizontal
          tokens={{ childrenGap: 10 }}
          horizontalAlign="center"
          styles={{ root: { marginTop: 10 } }}
        >
          <TextField
            placeholder="Enter UID (e.g., UID123456)"
            value={uid}
            onChange={(_e, v) => setUid(v ?? "")}
            styles={{
              fieldGroup: {
                width: 300,
                border: "1px solid #0078D4",
                borderRadius: "6px",
              },
            }}
          />
          <PrimaryButton
            text={loading ? "Triggering..." : "Search"}
            disabled={loading}
            onClick={handleSearch}
            styles={{
              root: {
                background: "#0078D4",
                borderRadius: "6px",
                padding: "0 24px",
              },
              rootHovered: { background: "#106EBE" },
            }}
          />
        </Stack>
      </Stack>
    </div>
  );
}
