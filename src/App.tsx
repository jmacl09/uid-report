import React, { useState } from "react";
import {
  initializeIcons,
  Stack,
  Text,
  TextField,
  PrimaryButton,
  Nav,
  Separator,
  DetailsList,
  DetailsListLayoutMode,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  IconButton,
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
  const [data, setData] = useState<any>(null);
  const [error, setError] = useState<string | null>(null);

  // ---------------------------
  // Helper to build Fluent UI columns dynamically
  // ---------------------------
  const buildColumns = (objArray: any[]) =>
    Object.keys(objArray[0] || {}).map((key) => ({
      key,
      name: key,
      fieldName: key,
      minWidth: 100,
      maxWidth: 250,
      isResizable: true,
      onRender: (item: any) =>
        key.toLowerCase().includes("workflow") ||
        key.toLowerCase().includes("diff") ? (
          <a
            href={item[key]}
            target="_blank"
            rel="noopener noreferrer"
            style={{ color: "#3AA0FF", textDecoration: "none" }}
          >
            Open
          </a>
        ) : (
          <span style={{ color: "#d0d0d0" }}>{item[key]}</span>
        ),
    }));

  // ---------------------------
  // Utility: Copy text to clipboard
  // ---------------------------
  const copyToClipboard = (text: string) => {
    navigator.clipboard.writeText(text);
    alert("Copied to clipboard!");
  };

  // ---------------------------
  // Table export helper
  // ---------------------------
  const exportToExcel = (tableData: any[], title: string) => {
    const headers = Object.keys(tableData[0] || {});
    const csv = [
      headers.join(","),
      ...tableData.map((row) =>
        headers.map((h) => JSON.stringify(row[h] ?? "")).join(",")
      ),
    ].join("\n");

    const blob = new Blob([csv], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${title}.csv`;
    a.click();
  };

  const exportToOneNote = (tableData: any[], title: string) => {
    const headers = Object.keys(tableData[0] || {});
    const html = `
      <h2>${title}</h2>
      <table border="1" cellpadding="4" cellspacing="0">
        <tr>${headers.map((h) => `<th>${h}</th>`).join("")}</tr>
        ${tableData
          .map(
            (row) =>
              `<tr>${headers.map((h) => `<td>${row[h] ?? ""}</td>`).join("")}</tr>`
          )
          .join("")}
      </table>`;
    const blob = new Blob([html], { type: "text/html" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${title}.html`;
    a.click();
  };

  // ---------------------------
  // Search handler
  // ---------------------------
  const handleSearch = async () => {
    if (!uid.trim()) {
      alert("Please enter a UID before searching.");
      return;
    }

    setLoading(true);
    setError(null);
    setData(null);

    const triggerUrl = `https://fibertools-dsavavdcfdgnh2cm.westeurope-01.azurewebsites.net/api/fiberflow/triggers/When_an_HTTP_request_is_received/invoke?api-version=2022-05-01&sp=%2Ftriggers%2FWhen_an_HTTP_request_is_received%2Frun&sv=1.0&sig=8KqIymphhOqUAlnd7UGwLRaxP0ot5ZH30b7jWCEUedQ&UID=${encodeURIComponent(
      uid
    )}`;

    try {
      const start = await fetch(triggerUrl, { method: "GET" });

      if (start.status === 202) {
        const statusUrl = start.headers.get("location");
        if (!statusUrl) throw new Error("No status URL returned by Logic App.");

        let result = null;
        for (let i = 0; i < 30; i++) {
          await new Promise((r) => setTimeout(r, 1000));
          const poll = await fetch(statusUrl);
          if (poll.status === 200) {
            result = await poll.json();
            break;
          }
        }

        if (result) {
          setData(result);
        } else {
          throw new Error("Timed out waiting for Logic App to complete.");
        }
      } else if (start.ok) {
        const result = await start.json();
        setData(result);
      } else {
        const text = await start.text();
        throw new Error(`HTTP ${start.status}: ${text}`);
      }
    } catch (err: any) {
      console.error(err);
      setError(err.message || "Network or Logic App error occurred.");
    } finally {
      setLoading(false);
    }
  };

  // ---------------------------
  // Page-wide Copy
  // ---------------------------
  const copyPageData = () => {
    if (!data) return;
    copyToClipboard(JSON.stringify(data, null, 2));
  };

  // ---------------------------
  // UI
  // ---------------------------
  return (
    <div
      style={{
        display: "flex",
        height: "100vh",
        backgroundColor: "#1b1a19",
        color: "#ffffff",
      }}
    >
      {/* Sidebar */}
      <div
        style={{
          width: "260px",
          backgroundColor: "#0a0a0a",
          color: "white",
          padding: "20px",
          display: "flex",
          flexDirection: "column",
          boxShadow: "2px 0 8px rgba(0,0,0,0.5)",
        }}
      >
        <Text
          variant="xLarge"
          styles={{
            root: { color: "#50b3ff", marginBottom: 20, fontWeight: 700 },
          }}
        >
          ⚡ FiberTools
        </Text>
        <Nav
          groups={navLinks}
          styles={{
            root: {
              width: 240,
              background: "transparent",
              selectors: {
                ".ms-Nav-compositeLink.is-selected": {
                  backgroundColor: "#0078D4",
                },
                ".ms-Button-flexContainer": { color: "#ffffff" },
              },
            },
            link: {
              color: "#d0d0d0",
              selectors: {
                ":hover": { background: "#0078D4", color: "#fff" },
              },
            },
          }}
        />
        <Separator styles={{ root: { borderColor: "#3AA0FF", marginTop: 20 } }} />
        <Text
          variant="small"
          styles={{
            root: {
              color: "#aaaaaa",
              marginTop: "auto",
              textAlign: "center",
              borderTop: "1px solid #333",
              paddingTop: 10,
            },
          }}
        >
          Built by <b>Josh Maclean</b> | Microsoft  
          <br />
          <span style={{ color: "#50b3ff" }}>All rights reserved ©2025</span>
        </Text>
      </div>

      {/* Main Content */}
      <Stack
        tokens={{ childrenGap: 20 }}
        styles={{
          root: {
            flexGrow: 1,
            padding: "40px",
            background: "linear-gradient(135deg,#111,#1c1c1c)",
            overflowY: "auto",
          },
        }}
      >
        <Stack horizontal horizontalAlign="space-between">
          <Text
            variant="xxLargePlus"
            styles={{
              root: {
                color: "#50b3ff",
                fontWeight: 700,
                textShadow: "0 0 10px rgba(80,179,255,0.5)",
              },
            }}
          >
            UID Lookup Portal
          </Text>
          <PrimaryButton
            text="Copy Page"
            iconProps={{ iconName: "Copy" }}
            onClick={copyPageData}
            styles={{
              root: {
                background: "#0078D4",
                borderRadius: "6px",
              },
              rootHovered: { background: "#106EBE" },
            }}
          />
        </Stack>

        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <TextField
            placeholder="Enter UID (e.g., 20190610163)"
            value={uid}
            onChange={(_e, v) => setUid(v ?? "")}
            styles={{
              fieldGroup: {
                width: 300,
                border: "1px solid #50b3ff",
                borderRadius: "6px",
                background: "#2b2b2b",
              },
              field: { color: "#fff" },
            }}
          />
          <PrimaryButton
            text={loading ? "Loading..." : "Search"}
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

        {loading && <Spinner size={SpinnerSize.large} label="Fetching data..." />}
        {error && (
          <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>
        )}

        {data && (
          <>
            {/* ===== OLS Section ===== */}
            <Section
              title="OLS Optical Link Summary"
              tableData={data.OLSLinks || []}
              buildColumns={buildColumns}
              exportToExcel={exportToExcel}
              exportToOneNote={exportToOneNote}
            />

            {/* ===== Associated UIDs ===== */}
            <Section
              title="Associated UIDs"
              tableData={data.AssociatedUIDs || []}
              buildColumns={buildColumns}
              exportToExcel={exportToExcel}
              exportToOneNote={exportToOneNote}
            />

            {/* ===== MGFX Split View ===== */}
            <Stack horizontal tokens={{ childrenGap: 20 }}>
              <Stack grow>
                <Section
                  title="MGFX A-Side"
                  tableData={(data.MGFX || []).filter((r: any) => r.Side === "A")}
                  buildColumns={buildColumns}
                  exportToExcel={exportToExcel}
                  exportToOneNote={exportToOneNote}
                />
              </Stack>
              <Stack grow>
                <Section
                  title="MGFX Z-Side"
                  tableData={(data.MGFX || []).filter((r: any) => r.Side === "Z")}
                  buildColumns={buildColumns}
                  exportToExcel={exportToExcel}
                  exportToOneNote={exportToOneNote}
                />
              </Stack>
            </Stack>
          </>
        )}
      </Stack>
    </div>
  );
}

// ---------------------------
// Reusable Table Section Component
// ---------------------------
const Section = ({ title, tableData, buildColumns, exportToExcel, exportToOneNote }: any) => {
  if (!tableData?.length) return null;

  return (
    <div
      style={{
        background: "#222",
        borderRadius: "10px",
        padding: "20px",
        boxShadow: "0 0 10px rgba(0,0,0,0.5)",
      }}
    >
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
        <Text
          variant="xLarge"
          styles={{
            root: {
              color: "#50b3ff",
              fontWeight: 600,
              borderLeft: "4px solid #50b3ff",
              paddingLeft: 10,
            },
          }}
        >
          {title}
        </Text>
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <IconButton
            iconProps={{ iconName: "Copy" }}
            title="Copy Table"
            onClick={() => navigator.clipboard.writeText(JSON.stringify(tableData, null, 2))}
          />
          <IconButton
            iconProps={{ iconName: "ExcelDocument" }}
            title="Export to Excel"
            onClick={() => exportToExcel(tableData, title)}
          />
          <IconButton
            iconProps={{ iconName: "OneNoteLogo" }}
            title="Export to OneNote"
            onClick={() => exportToOneNote(tableData, title)}
          />
        </Stack>
      </Stack>

      <DetailsList
        items={tableData}
        columns={buildColumns(tableData)}
        layoutMode={DetailsListLayoutMode.justified}
      />
    </div>
  );
};
