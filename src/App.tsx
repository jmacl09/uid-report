import React, { useState } from "react";
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

  // ✅ Direct Logic App trigger (GET)
  const handleSearch = async () => {
    if (!uid.trim()) {
      alert("Please enter a UID before searching.");
      return;
    }

    setLoading(true);

    try {
      const response = await fetch(
        `https://fibertools-dsavavdcfdgnh2cm.westeurope-01.azurewebsites.net/api/fiberflow/triggers/When_an_HTTP_request_is_received/invoke?api-version=2022-05-01&sp=%2Ftriggers%2FWhen_an_HTTP_request_is_received%2Frun&sv=1.0&sig=8KqIymphhOqUAlnd7UGwLRaxP0ot5ZH30b7jWCEUedQ&UID=${encodeURIComponent(
          uid
        )}`,
        { method: "GET" }
      );

      if (response.ok) {
        alert(`✅ Flow triggered successfully for UID: ${uid}`);
      } else {
        const text = await response.text();
        alert(`❌ Flow failed (${response.status}): ${text}`);
      }
    } catch (err) {
      console.error(err);
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
        }}
      >
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
        <Separator styles={{ root: { borderColor: "#fff", marginTop: 20 } }} />
        <Text variant="small" styles={{ root: { color: "#d0d0d0", marginTop: 10 } }}>
          Built by Josh Maclean | Microsoft
        </Text>
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
