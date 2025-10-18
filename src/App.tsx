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

  const buildColumns = (objArray: any[]) =>
    Object.keys(objArray[0] || {}).map((key) => ({
      key,
      name: key,
      fieldName: key,
      minWidth: 100,
      maxWidth: 250,
      isResizable: true,
      styles: { root: { color: "#fff" } },
      onRender: (item: any) =>
        key.toLowerCase().includes("workflow") ||
        key.toLowerCase().includes("diff") ||
        key.toLowerCase().includes("ticketlink") ? (
          <a
            href={item[key]}
            target="_blank"
            rel="noopener noreferrer"
            style={{ color: "#3AA0FF", textDecoration: "none" }}
          >
            Open
          </a>
        ) : (
          <span style={{ color: "#ccc" }}>{item[key]}</span>
        ),
    }));

  const copyToClipboard = (text: string) => {
    navigator.clipboard.writeText(text);
    alert("Copied to clipboard!");
  };

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

        if (result) setData(result);
        else throw new Error("Timed out waiting for Logic App to complete.");
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

  const copyPageData = () => {
    if (!data) return;
    const allData = {
      OLSLinks: data.OLSLinks,
      AssociatedUIDs: data.AssociatedUIDs,
      MGFXA: data.MGFXA,
      MGFXZ: data.MGFXZ,
      GDCOTickets: data.GDCOTickets,
    };
    copyToClipboard(JSON.stringify(allData, null, 2));
  };

  return (
    <div
      style={{
        display: "flex",
        height: "100vh",
        backgroundColor: "#111",
        color: "#fff",
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
            background: "radial-gradient(circle at top left,#1a1a1a 0%,#111 100%)",
            overflowY: "auto",
          },
        }}
      >
        {/* Header */}
        <Stack horizontal horizontalAlign="space-between">
          <Text
            variant="xxLargePlus"
            styles={{
              root: {
                color: "#50b3ff",
                fontWeight: 700,
                textShadow: "0 0 10px rgba(80,179,255,0.6)",
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

        {/* Search Input */}
        <Stack
          horizontalAlign="center"
          tokens={{ childrenGap: 10 }}
          style={{ marginTop: 10 }}
        >
          <Stack horizontal tokens={{ childrenGap: 10 }}>
            <TextField
              placeholder="Enter UID (e.g., 20190610163)"
              value={uid}
              onChange={(_e, v) => setUid(v ?? "")}
              styles={{
                fieldGroup: {
                  width: 320,
                  border: "1px solid #50b3ff",
                  borderRadius: "8px",
                  background: "#1c1c1c",
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
                  borderRadius: "8px",
                  padding: "0 24px",
                },
                rootHovered: { background: "#106EBE" },
              }}
            />
          </Stack>
          {loading && (
            <div
              style={{
                display: "flex",
                alignItems: "center",
                marginTop: 20,
                gap: 12,
              }}
            >
              <Spinner
                size={SpinnerSize.large}
                label="Fetching data..."
                styles={{
                  label: { color: "#50b3ff", fontSize: 16 },
                }}
              />
              <div
                style={{
                  width: 200,
                  height: 8,
                  borderRadius: 4,
                  background: "linear-gradient(90deg,#0078D4,#50b3ff,#0078D4)",
                  animation: "pulse 2s infinite linear",
                }}
              />
            </div>
          )}
        </Stack>

        {error && (
          <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>
        )}

        {data && (
          <>
            <Section title="OLS Optical Link Summary" data={data.OLSLinks} buildColumns={buildColumns} exportToExcel={exportToExcel} exportToOneNote={exportToOneNote}/>
            <Section title="Associated UIDs" data={data.AssociatedUIDs} buildColumns={buildColumns} exportToExcel={exportToExcel} exportToOneNote={exportToOneNote}/>
            <Stack horizontal tokens={{ childrenGap: 20 }}>
              <Stack grow>
                <Section title="MGFX A-Side" data={data.MGFXA} buildColumns={buildColumns} exportToExcel={exportToExcel} exportToOneNote={exportToOneNote}/>
              </Stack>
              <Stack grow>
                <Section title="MGFX Z-Side" data={data.MGFXZ} buildColumns={buildColumns} exportToExcel={exportToExcel} exportToOneNote={exportToOneNote}/>
              </Stack>
            </Stack>
            <Section title="GDCO Tickets" data={data.GDCOTickets} buildColumns={buildColumns} exportToExcel={exportToExcel} exportToOneNote={exportToOneNote}/>
          </>
        )}
      </Stack>
    </div>
  );
}

const Section = ({ title, data, buildColumns, exportToExcel, exportToOneNote }: any) => {
  if (!data?.length) return null;
  return (
    <div
      style={{
        background: "#181818",
        borderRadius: "12px",
        padding: "20px",
        boxShadow: "0 0 15px rgba(0,0,0,0.6)",
        border: "1px solid #333",
        transition: "0.2s ease",
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
          <IconButton iconProps={{ iconName: "Copy" }} title="Copy Table" onClick={() => navigator.clipboard.writeText(JSON.stringify(data, null, 2))}/>
          <IconButton iconProps={{ iconName: "ExcelDocument" }} title="Export to Excel" onClick={() => exportToExcel(data, title)}/>
          <IconButton iconProps={{ iconName: "OneNoteLogo" }} title="Export to OneNote" onClick={() => exportToOneNote(data, title)}/>
        </Stack>
      </Stack>
      <DetailsList
        items={data}
        columns={buildColumns(data)}
        layoutMode={DetailsListLayoutMode.justified}
        styles={{
          root: {
            marginTop: 10,
            background: "#181818",
          },
          headerWrapper: {
            background: "#222",
          },
          contentWrapper: {
            selectors: {
              ".ms-DetailsRow": {
                backgroundColor: "#181818",
                color: "#fff",
              },
              ".ms-DetailsRow:hover": {
                backgroundColor: "#242424",
                boxShadow: "0 0 10px rgba(80,179,255,0.3)",
              },
              ".ms-DetailsRow:nth-child(even)": {
                backgroundColor: "#202020",
              },
            },
          },
        }}
      />
    </div>
  );
};
