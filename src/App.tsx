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
import "./App.css";

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
  const [summary, setSummary] = useState<string>("Awaiting UID lookup...");
  const [showAllOLS, setShowAllOLS] = useState<boolean>(false);

  const naturalSort = (a: string, b: string) =>
    a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" });

  const buildColumns = (objArray: any[]) =>
    Object.keys(objArray[0] || {}).map((key) => ({
      key,
      name: key,
      fieldName: key,
      minWidth: 70,
      maxWidth: 180,
      isResizable: true,
      isMultiline: false,
      onRender: (item: any) => {
        const val = item[key];
        if (
          key.toLowerCase().includes("workflow") ||
          key.toLowerCase().includes("diff") ||
          key.toLowerCase().includes("ticketlink")
        ) {
          return (
            <a
              href={val}
              target="_blank"
              rel="noopener noreferrer"
              style={{ color: "#3AA0FF", textDecoration: "none" }}
            >
              Open
            </a>
          );
        }
        return <span style={{ color: "#d0d0d0", whiteSpace: "nowrap" }}>{val}</span>;
      },
    }));

  const handleSearch = async () => {
    if (!uid.trim()) {
      alert("Please enter a UID before searching.");
      return;
    }

    setLoading(true);
    setError(null);
    setData(null);
    setShowAllOLS(false);
    setSummary("Analyzing data...");

    const triggerUrl = `https://fibertools-dsavavdcfdgnh2cm.westeurope-01.azurewebsites.net/api/fiberflow/triggers/When_an_HTTP_request_is_received/invoke?api-version=2022-05-01&sp=%2Ftriggers%2FWhen_an_HTTP_request_is_received%2Frun&sv=1.0&sig=8KqIymphhOqUAlnd7UGwLRaxP0ot5ZH30b7jWCEUedQ&UID=${encodeURIComponent(
      uid
    )}`;

    try {
      const res = await fetch(triggerUrl);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const result = await res.json();

      result.OLSLinks?.sort((a: any, b: any) => naturalSort(a.APort, b.APort));
      result.AssociatedUIDs?.sort(
        (a: any, b: any) => parseInt(b.Uid) - parseInt(a.Uid)
      );
      result.MGFXA?.sort((a: any, b: any) =>
        a.XOMT.localeCompare(b.XOMT, undefined, { numeric: true })
      );
      result.MGFXZ?.sort((a: any, b: any) =>
        a.XOMT.localeCompare(b.XOMT, undefined, { numeric: true })
      );

      setData(result);
      setTimeout(() => makeSummary(result), 800);
    } catch (err: any) {
      setError(err.message || "Network error occurred.");
      setSummary("Error retrieving data.");
    } finally {
      setLoading(false);
    }
  };

  const makeSummary = (d: Record<string, any>) => {
    if (!d) return;
    const links = d.OLSLinks?.length || 0;
    const uids = d.AssociatedUIDs?.length || 0;
    const mgfxA = d.MGFXA?.length || 0;
    const mgfxZ = d.MGFXZ?.length || 0;
    const tickets = d.GDCOTickets?.length || 0;
    setSummary(
      `Found ${links} active optical paths, ${uids} associated UIDs, ${mgfxA + mgfxZ
      } MGFX fiber ends, and ${tickets} related GDCO tickets.`
    );
  };

  const Section = ({ title, rows, highlightUid }: any) => {
    if (!rows?.length) return null;
    const filtered = rows.map((r: any) => {
      const copy = { ...r };
      delete copy.Side;
      return copy;
    });

    return (
      <div
        className="compact-table"
        style={{
          background: "#181818",
          borderRadius: 8,
          padding: "8px 10px",
          border: "1px solid #2b2b2b",
          marginBottom: 14,
        }}
      >
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Text
            variant="large"
            styles={{
              root: {
                color: "#50b3ff",
                fontWeight: 600,
                marginBottom: 6,
              },
            }}
          >
            {title}
          </Text>
          <Stack horizontal tokens={{ childrenGap: 6 }}>
            <IconButton
              iconProps={{ iconName: "Copy" }}
              title="Copy JSON"
              onClick={() =>
                navigator.clipboard.writeText(JSON.stringify(filtered, null, 2))
              }
            />
            <IconButton
              iconProps={{ iconName: "OneNoteLogo" }}
              title="Export to OneNote"
              onClick={() => exportToOneNote(filtered, title)}
            />
          </Stack>
        </Stack>

        <DetailsList
          items={filtered}
          columns={buildColumns(filtered)}
          layoutMode={DetailsListLayoutMode.fixedColumns}
          compact={true}
          styles={{
            root: {
              background: "#181818",
              overflowX: "hidden",
              maxWidth: "fit-content",
            },
          }}
          onRenderRow={(props, defaultRender) => {
            if (!props) return null;
            const isHighlight = highlightUid && props.item.Uid === highlightUid;
            return (
              <div
                style={{
                  boxShadow: isHighlight
                    ? "0 0 8px rgba(80,179,255,0.7)"
                    : "none",
                  borderRadius: isHighlight ? 4 : 0,
                }}
              >
                {defaultRender?.(props)}
              </div>
            );
          }}
        />
      </div>
    );
  };

  const exportToOneNote = (tableData: any[], title: string) => {
    const headers = Object.keys(tableData[0] || {});
    const html = `
      <div style="font-family:Segoe UI;background:#1b1b1b;color:#fff;padding:10px">
        <h2 style="color:#fff;background:linear-gradient(90deg,#0078D4,#3AA0FF);padding:4px 10px;border-radius:4px">${title}</h2>
        <table border="1" cellspacing="0" cellpadding="4" style="width:auto;border-collapse:collapse;border-color:#333">
          <tr style="background:linear-gradient(90deg,#0078D4,#3AA0FF);color:#fff;font-weight:600">${headers
            .map((h) => `<th>${h}</th>`)
            .join("")}</tr>
          ${tableData
            .map(
              (row, i) =>
                `<tr style="background:${i % 2 === 0 ? "#181818" : "#202020"}">${headers
                  .map((h) => `<td>${row[h] ?? ""}</td>`)
                  .join("")}</tr>`
            )
            .join("")}
        </table>
      </div>`;
    const blob = new Blob([html], { type: "text/html" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${title}.html`;
    a.click();
  };

  const OLSSection = ({ title, rows }: any) => {
    if (!rows?.length) return null;
    const displayed = showAllOLS ? rows : rows.slice(0, 10);
    return (
      <>
        <Section title={title} rows={displayed} />
        {rows.length > 10 && !showAllOLS && (
          <PrimaryButton
            text="Show More"
            onClick={() => setShowAllOLS(true)}
            styles={{
              root: {
                marginTop: 6,
                background: "#0078D4",
                borderRadius: 6,
                padding: "0 16px",
              },
              rootHovered: { background: "#106EBE" },
            }}
          />
        )}
      </>
    );
  };

  return (
    <div style={{ display: "flex", height: "100vh", backgroundColor: "#111" }}>
      {/* Sidebar */}
      <div
        style={{
          width: 240,
          backgroundColor: "#0a0a0a",
          padding: 20,
          display: "flex",
          flexDirection: "column",
        }}
      >
        <Text
          variant="xLarge"
          styles={{ root: { color: "#50b3ff", marginBottom: 20 } }}
        >
          ⚡ FiberTools
        </Text>
        <Nav
          groups={navLinks}
          styles={{
            link: {
              color: "#ccc",
              selectors: { ":hover": { background: "#0078D4", color: "#fff" } },
            },
          }}
        />
        <Separator styles={{ root: { borderColor: "#3AA0FF", marginTop: 20 } }} />
        <Text
          variant="small"
          styles={{
            root: {
              color: "#999",
              marginTop: "auto",
              textAlign: "center",
              borderTop: "1px solid #333",
              paddingTop: 8,
            },
          }}
        >
          Built by <b>Josh Maclean</b> | Microsoft
          <br />
          <span style={{ color: "#50b3ff" }}>All rights reserved ©2025</span>
        </Text>
      </div>

      {/* Main */}
      <Stack
        tokens={{ childrenGap: 18 }}
        styles={{
          root: {
            flexGrow: 1,
            padding: 30,
            overflowY: "auto",
            alignItems: "flex-start",
          },
        }}
      >
        <Text
          variant="xxLargePlus"
          styles={{
            root: {
              textAlign: "center",
              color: "#50b3ff",
              fontWeight: 700,
              textShadow: "0 0 10px rgba(80,179,255,0.6)",
              marginBottom: 10,
              width: "100%",
            },
          }}
        >
          UID Lookup Portal
        </Text>

        {/* UID Input */}
        <Stack horizontalAlign="center" tokens={{ childrenGap: 10 }}>
          <Stack horizontal tokens={{ childrenGap: 10 }}>
            <TextField
              placeholder="Enter UID (e.g., 20190610163)"
              value={uid}
              onChange={(_e, v) => setUid(v ?? "")}
              styles={{
                fieldGroup: {
                  width: 300,
                  border: "1px solid #50b3ff",
                  borderRadius: 8,
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
                  borderRadius: 8,
                  padding: "0 24px",
                },
                rootHovered: { background: "#106EBE" },
              }}
            />
          </Stack>
          {loading && (
            <Spinner
              size={SpinnerSize.large}
              label="Fetching data..."
              styles={{ label: { color: "#50b3ff", fontSize: 14 } }}
            />
          )}
        </Stack>

        <div
          style={{
            marginTop: 8,
            textAlign: "center",
            color: "#50b3ff",
            fontWeight: 500,
            fontSize: 14,
          }}
        >
          {summary}
        </div>

        {error && (
          <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>
        )}

        {data && (
          <>
            <OLSSection title="OLS Optical Link Summary" rows={data.OLSLinks} />
            <Stack horizontal tokens={{ childrenGap: 20 }}>
              <Section
                title="Associated UIDs"
                rows={data.AssociatedUIDs}
                highlightUid={uid}
              />
              <Section title="GDCO Tickets" rows={data.GDCOTickets} />
            </Stack>
            <Stack horizontal tokens={{ childrenGap: 20 }}>
              <Section title="MGFX A-Side" rows={data.MGFXA} />
              <Section title="MGFX Z-Side" rows={data.MGFXZ} />
            </Stack>
          </>
        )}
      </Stack>
    </div>
  );
}
