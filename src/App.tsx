import React, { useState, useEffect } from "react";
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
  const [showScroll, setShowScroll] = useState<boolean>(false);
  const [summary, setSummary] = useState<string>("Awaiting UID lookup...");

  useEffect(() => {
    const onScroll = () => setShowScroll(window.scrollY > 300);
    window.addEventListener("scroll", onScroll);
    return () => window.removeEventListener("scroll", onScroll);
  }, []);

  const buildColumns = (objArray: any[]) =>
    Object.keys(objArray[0] || {}).map((key) => ({
      key,
      name: key,
      fieldName: key,
      minWidth: 80,
      maxWidth: 200,
      isResizable: true,
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
          <span style={{ color: "#d0d0d0" }}>{item[key]}</span>
        ),
    }));

  const handleSearch = async () => {
    if (!uid.trim()) {
      alert("Please enter a UID before searching.");
      return;
    }

    setLoading(true);
    setError(null);
    setData(null);
    setSummary("Analyzing data...");

    const triggerUrl = `https://fibertools-dsavavdcfdgnh2cm.westeurope-01.azurewebsites.net/api/fiberflow/triggers/When_an_HTTP_request_is_received/invoke?api-version=2022-05-01&sp=%2Ftriggers%2FWhen_an_HTTP_request_is_received%2Frun&sv=1.0&sig=8KqIymphhOqUAlnd7UGwLRaxP0ot5ZH30b7jWCEUedQ&UID=${encodeURIComponent(
      uid
    )}`;

    try {
      const res = await fetch(triggerUrl);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const result = await res.json();

      // Sort MGFX by XOMT
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

  // ---------- Smart Summary ----------
  const makeSummary = (d: Record<string, any>) => {
    if (!d) return;
    const links = d.OLSLinks?.length || 0;
    const uids = d.AssociatedUIDs?.length || 0;
    const mgfxA = d.MGFXA?.length || 0;
    const mgfxZ = d.MGFXZ?.length || 0;
    const tickets = d.GDCOTickets?.length || 0;
    const active = links ? "active optical paths" : "no OLS data";
    const msg = `Found ${links} ${active}, ${uids} associated UIDs, ${mgfxA + mgfxZ
      } MGFX fiber ends, and ${tickets} related GDCO tickets.`;
    setSummary(msg);
  };

  // ---------- Export Helpers ----------
  const exportToOneNote = (tableData: any[], title: string) => {
    const headers = Object.keys(tableData[0] || {});
    const html = `
      <div style="font-family:Segoe UI;background:#1b1b1b;color:#fff;padding:10px">
        <h2 style="color:#50b3ff;border-left:4px solid #50b3ff;padding-left:6px">${title}</h2>
        <table border="1" cellspacing="0" cellpadding="4" style="width:100%;border-collapse:collapse;border-color:#333">
          <tr style="background:#222;color:#50b3ff">${headers
            .map((h) => `<th>${h}</th>`)
            .join("")}</tr>
          ${tableData
            .map(
              (row, i) =>
                `<tr style="background:${
                  i % 2 === 0 ? "#181818" : "#202020"
                }">${headers.map((h) => `<td>${row[h] ?? ""}</td>`).join("")}</tr>`
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

  const copyPageHTML = () => {
    if (!data) return;
    const sections = [
      { title: "OLS Optical Link Summary", rows: data.OLSLinks },
      { title: "Associated UIDs", rows: data.AssociatedUIDs },
      { title: "MGFX A-Side", rows: data.MGFXA },
      { title: "MGFX Z-Side", rows: data.MGFXZ },
      { title: "GDCO Tickets", rows: data.GDCOTickets },
    ];
    let html = `<div style="font-family:Segoe UI;background:#1b1b1b;color:#fff">`;
    for (const s of sections) {
      if (!s.rows?.length) continue;
      const headers = Object.keys(s.rows[0]);
      html += `
        <h2 style="color:#50b3ff;border-left:4px solid #50b3ff;padding-left:6px">${s.title}</h2>
        <table border="1" cellspacing="0" cellpadding="4" style="width:100%;border-collapse:collapse;border-color:#333;margin-bottom:15px">
          <tr style="background:#222;color:#50b3ff">${headers
            .map((h) => `<th>${h}</th>`)
            .join("")}</tr>
          ${s.rows
            .map(
              (row: Record<string, any>, i: number) =>
                `<tr style="background:${
                  i % 2 === 0 ? "#181818" : "#202020"
                }">${headers.map((h) => `<td>${row[h] ?? ""}</td>`).join("")}</tr>`
            )
            .join("")}
        </table>`;
    }
    html += "</div>";
    navigator.clipboard.writeText(html);
    alert("All tables copied as formatted HTML for OneNote ✅");
  };

  // ---------- Section ----------
  const Section = ({ title, rows }: any) => {
    if (!rows?.length) return null;
    const filtered = rows.map((r: any) => {
      const copy = { ...r };
      delete copy.Side;
      return copy;
    });
    return (
      <div
        style={{
          background: "#181818",
          borderRadius: "10px",
          padding: "10px 14px",
          border: "1px solid #2b2b2b",
        }}
      >
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Text
            variant="large"
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
          <Stack horizontal tokens={{ childrenGap: 6 }}>
            <IconButton
              iconProps={{ iconName: "Copy" }}
              title="Copy Table"
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
          layoutMode={DetailsListLayoutMode.justified}
          styles={{
            root: { marginTop: 6, background: "#181818" },
            headerWrapper: { background: "#222" },
            contentWrapper: {
              selectors: {
                ".ms-DetailsRow": {
                  backgroundColor: "#181818",
                  minHeight: 26,
                },
                ".ms-DetailsRow:hover": {
                  backgroundColor: "#242424",
                  boxShadow: "0 0 6px rgba(80,179,255,0.4)",
                },
                ".ms-DetailsRow:nth-child(even)": { backgroundColor: "#202020" },
              },
            },
          }}
        />
      </div>
    );
  };

  // ---------- Main Layout ----------
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
        styles={{ root: { flexGrow: 1, padding: 30, overflowY: "auto" } }}
      >
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <div style={{ flex: 1 }} />
          <Text
            variant="xxLargePlus"
            styles={{
              root: {
                textAlign: "center",
                color: "#50b3ff",
                fontWeight: 700,
                textShadow: "0 0 10px rgba(80,179,255,0.6)",
              },
            }}
          >
            UID Lookup Portal
          </Text>
          <div
            style={{
              width: 260,
              background: "#181818",
              borderRadius: 8,
              padding: "8px 12px",
              border: "1px solid #333",
              color: "#ccc",
              fontSize: 13,
            }}
          >
            <b>AI Summary:</b>
            <div style={{ color: "#50b3ff", marginTop: 4 }}>{summary}</div>
          </div>
        </Stack>

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
            <div style={{ marginTop: 10 }}>
              <Spinner
                size={SpinnerSize.large}
                label="Fetching data..."
                styles={{ label: { color: "#50b3ff", fontSize: 14 } }}
              />
            </div>
          )}
        </Stack>

        {error && (
          <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>
        )}

        {data && (
          <>
            <Section title="OLS Optical Link Summary" rows={data.OLSLinks} />
            <Stack horizontal tokens={{ childrenGap: 20 }}>
              <Stack grow>
                <Section title="Associated UIDs" rows={data.AssociatedUIDs} />
              </Stack>
              <Stack grow>
                <Section title="GDCO Tickets" rows={data.GDCOTickets} />
              </Stack>
            </Stack>
            <Stack horizontal tokens={{ childrenGap: 20 }}>
              <Stack grow>
                <Section title="MGFX A-Side" rows={data.MGFXA} />
              </Stack>
              <Stack grow>
                <Section title="MGFX Z-Side" rows={data.MGFXZ} />
              </Stack>
            </Stack>
          </>
        )}
      </Stack>

      {/* Scroll to Top */}
      {showScroll && (
        <button
          onClick={() => window.scrollTo({ top: 0, behavior: "smooth" })}
          style={{
            position: "fixed",
            bottom: 30,
            right: 30,
            background: "#0078D4",
            border: "none",
            color: "#fff",
            padding: "10px 14px",
            borderRadius: "50%",
            fontSize: 18,
            boxShadow: "0 0 10px rgba(80,179,255,0.6)",
            cursor: "pointer",
          }}
        >
          ↑
        </button>
      )}

      {/* Copy Page Button */}
      <PrimaryButton
        text="Copy Page"
        onClick={copyPageHTML}
        iconProps={{ iconName: "Copy" }}
        styles={{
          root: {
            position: "fixed",
            top: 30,
            right: 30,
            background: "#0078D4",
            borderRadius: "8px",
          },
          rootHovered: { background: "#106EBE" },
        }}
      />
    </div>
  );
}
