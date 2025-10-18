import React, { useState, useEffect } from "react";
import {
  initializeIcons,
  Stack,
  Text,
  TextField,
  PrimaryButton,
  Nav,
  Separator,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  IconButton,
} from "@fluentui/react";
import { saveAs } from "file-saver";
import * as XLSX from "xlsx";
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
  const [history, setHistory] = useState<string[]>([]);
  const [loading, setLoading] = useState<boolean>(false);
  const [data, setData] = useState<any>(null);
  const [error, setError] = useState<string | null>(null);
  const [summary, setSummary] = useState<string>("Awaiting UID lookup...");

  useEffect(() => {
    const saved = JSON.parse(localStorage.getItem("uidHistory") || "[]");
    setHistory(saved);
  }, []);

  useEffect(() => {
    localStorage.setItem("uidHistory", JSON.stringify(history.slice(0, 10)));
  }, [history]);

  const naturalSort = (a: string, b: string) =>
    a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" });

  const handleSearch = async (searchUid?: string) => {
    const query = searchUid || uid;
    if (!query.trim()) {
      alert("Please enter a UID before searching.");
      return;
    }

    setLoading(true);
    setError(null);
    setData(null);
    setSummary("Analyzing data...");

    const triggerUrl = `https://fibertools-dsavavdcfdgnh2cm.westeurope-01.azurewebsites.net/api/fiberflow/triggers/When_an_HTTP_request_is_received/invoke?api-version=2022-05-01&sp=%2Ftriggers%2FWhen_an_HTTP_request_is_received%2Frun&sv=1.0&sig=8KqIymphhOqUAlnd7UGwLRaxP0ot5ZH30b7jWCEUedQ&UID=${encodeURIComponent(
      query
    )}`;

    try {
      const res = await fetch(triggerUrl);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const result = await res.json();

      result.OLSLinks?.sort((a: any, b: any) => naturalSort(a.APort, b.APort));
      result.MGFXA?.sort((a: any, b: any) => naturalSort(a.XOMT, b.XOMT));
      result.MGFXZ?.sort((a: any, b: any) => naturalSort(a.XOMT, b.XOMT));

      setData(result);
      setSummary(
        `Found ${result.OLSLinks?.length || 0} active optical paths, ${
          result.AssociatedUIDs?.length || 0
        } associated UIDs, and ${result.GDCOTickets?.length || 0} related GDCO tickets.`
      );

      if (!history.includes(query)) setHistory([query, ...history]);
    } catch (err: any) {
      setError(err.message || "Network error occurred.");
      setSummary("Error retrieving data.");
    } finally {
      setLoading(false);
    }
  };

  const copyTableText = (title: string, rows: any[], headers: string[]) => {
    const text =
      `${title}\n` +
      headers.join("\t") +
      "\n" +
      rows.map((r) => Object.values(r).join("\t")).join("\n") +
      "\n\n";
    navigator.clipboard.writeText(text);
    alert(`Copied ${title} (plain text) ✅`);
  };

  const copyAllTables = () => {
    if (!data) return;

    const sections = [
      {
        title: "Optical Link Summary",
        rows: data.OLSLinks,
        headers: [
          "A Device",
          "A Port",
          "Z Device",
          "Z Port",
          "A Optical Device",
          "Z Optical Device",
          "Z Optical Port",
          "Workflow",
        ],
      },
      {
        title: "Associated UIDs",
        rows: data.AssociatedUIDs,
        headers: [
          "UID",
          "SRLG ID",
          "Action",
          "Type",
          "Device A",
          "Device Z",
          "Site A",
          "Site Z",
          "Lag A",
          "Lag Z",
        ],
      },
      {
        title: "GDCO Tickets",
        rows: data.GDCOTickets,
        headers: [
          "Ticket ID",
          "Datacenter Code",
          "Title",
          "State",
          "Assigned To",
          "Ticket Link",
        ],
      },
      {
        title: "MGFX A-Side",
        rows: data.MGFXA?.map(({ Side, ...keep }: any) => keep),
        headers: [
          "XOMT",
          "CO Device",
          "CO Port",
          "MO Device",
          "MO Port",
          "CO Diff",
          "MO Diff",
        ],
      },
      {
        title: "MGFX Z-Side",
        rows: data.MGFXZ?.map(({ Side, ...keep }: any) => keep),
        headers: [
          "XOMT",
          "CO Device",
          "CO Port",
          "MO Device",
          "MO Port",
          "CO Diff",
          "MO Diff",
        ],
      },
    ];

    let combinedText = "";
    for (const s of sections) {
      if (s.rows && s.rows.length > 0) {
        combinedText += `${s.title}\n${s.headers.join("\t")}\n${s.rows
          .map((r) => Object.values(r).join("\t"))
          .join("\n")}\n\n`;
      }
    }

    navigator.clipboard.writeText(combinedText.trim());
    alert("Copied all tables (plain text) ✅");
  };

  const exportExcel = () => {
    if (!data || !uid) return;
    const wb = XLSX.utils.book_new();

    const sections = {
      "OLS Optical Link Summary": data.OLSLinks,
      "Associated UIDs": data.AssociatedUIDs,
      "GDCO Tickets": data.GDCOTickets,
      "MGFX A-Side": data.MGFXA,
      "MGFX Z-Side": data.MGFXZ,
    };

    for (const [title, rows] of Object.entries(sections)) {
      if (!Array.isArray(rows) || !rows.length) continue;
      const ws = XLSX.utils.json_to_sheet(rows);
      XLSX.utils.book_append_sheet(wb, ws, title.slice(0, 31));
    }

    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = new Blob([wbout], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    saveAs(blob, `UID_Report_${uid}.xlsx`);
  };

  const exportOneNote = () => {
    if (!data || !uid) return;
    const tables = document.querySelectorAll(".data-table");
    let html = `
      <html><head><meta charset="utf-8">
      <title>UID Report ${uid}</title>
      <style>
        body {font-family:'Segoe UI';background:#fff;color:#000;}
        h1 {color:#0078d4;}
        table {border-collapse:collapse;margin-bottom:20px;width:100%;background:#f9f9f9;border-radius:6px;overflow:hidden;}
        th {background:#0078d4;color:#fff;padding:6px 10px;text-align:left;}
        td {padding:5px 10px;border-bottom:1px solid #ddd;}
        tr:nth-child(even){background:#f2f2f2;}
      </style>
      </head><body><h1>UID Report ${uid}</h1>`;

    tables.forEach((tbl: any) => (html += tbl.outerHTML));
    html += "</body></html>";

    const blob = new Blob([html], { type: "multipart/related" });
    const fileName = `UID_Report_${uid}_OneNote.mht`;
    saveAs(blob, fileName);

    // Attempt to open automatically in OneNote
    setTimeout(() => {
      window.location.href = `onenote:${fileName}`;
    }, 800);
  };

  const Table = ({ title, headers, rows, highlightUid }: any) => {
    if (!rows?.length) return null;

    return (
      <div className="table-container">
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Text className="section-title">{title}</Text>
          <Stack horizontal tokens={{ childrenGap: 6 }}>
            <IconButton
              iconProps={{ iconName: "Copy" }}
              title="Copy Table (Plain Text)"
              onClick={() => copyTableText(title, rows, headers)}
            />
          </Stack>
        </Stack>
        <table className="data-table">
          <thead>
            <tr>
              {headers.map((h: string, i: number) => (
                <th key={i}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((row: any, i: number) => {
              const keys = Object.keys(row);
              const highlight =
                highlightUid && row.Uid?.toString() === highlightUid;
              return (
                <tr key={i} className={highlight ? "highlight-row" : ""}>
                  {keys.map((key, j) => {
                    const val = row[key];
                    if (
                      key.toLowerCase().includes("workflow") ||
                      key.toLowerCase().includes("diff") ||
                      key.toLowerCase().includes("ticketlink")
                    ) {
                      return (
                        <td key={j}>
                          <button
                            className="open-btn"
                            onClick={() => window.open(val, "_blank")}
                          >
                            Open
                          </button>
                        </td>
                      );
                    }
                    return <td key={j}>{val}</td>;
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    );
  };

  return (
    <div style={{ display: "flex", height: "100vh", backgroundColor: "#111" }}>
      <div className="sidebar">
        <Text variant="xLarge" className="logo">
          ⚡ FiberTools
        </Text>
        <Nav groups={navLinks} />
        <Separator />
        <Text className="footer">
          Built by <b>Josh Maclean</b> | Microsoft
        </Text>
      </div>

      <Stack className="main">
        {/* Top Header */}
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Text className="portal-title">UID Lookup Portal</Text>
          <Stack horizontal tokens={{ childrenGap: 10 }}>
            <IconButton
              iconProps={{ iconName: "Copy" }}
              title="Copy All Tables (Plain Text)"
              className="copy-all-btn"
              onClick={copyAllTables}
            />
            <IconButton
              iconProps={{ iconName: "ExcelLogo" }}
              title="Export to Excel"
              className="excel-btn"
              onClick={exportExcel}
            />
            <IconButton
              iconProps={{ iconName: "OneNoteLogo" }}
              title="Export to OneNote"
              className="onenote-btn"
              onClick={exportOneNote}
            />
          </Stack>
        </Stack>

        {/* Search */}
        <Stack horizontalAlign="center" tokens={{ childrenGap: 10 }}>
          <Stack horizontal tokens={{ childrenGap: 10 }}>
            <TextField
              placeholder="Enter UID (e.g., 20190610163)"
              value={uid}
              onChange={(_e, v) => setUid(v ?? "")}
              className="input-field"
            />
            <PrimaryButton
              text={loading ? "Loading..." : "Search"}
              disabled={loading}
              onClick={() => handleSearch()}
              className="search-btn"
            />
          </Stack>
          {loading && <Spinner size={SpinnerSize.large} label="Fetching data..." />}
        </Stack>

        {/* UID History */}
        {history.length > 0 && (
          <div className="uid-history">
            Recent:{" "}
            {history.map((item, i) => (
              <span key={i} onClick={() => handleSearch(item)}>
                {item}
              </span>
            ))}
          </div>
        )}

        <div className="summary">{summary}</div>

        {error && (
          <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>
        )}

        {data && (
          <>
            <Table
              title="Optical Link Summary"
              headers={[
                "A Device",
                "A Port",
                "Z Device",
                "Z Port",
                "A Optical Device",
                "Z Optical Device",
                "Z Optical Port",
                "Workflow",
              ]}
              rows={data.OLSLinks}
            />
            <Stack horizontal tokens={{ childrenGap: 20 }}>
              <Table
                title="Associated UIDs"
                headers={[
                  "UID",
                  "SRLG ID",
                  "Action",
                  "Type",
                  "Device A",
                  "Device Z",
                  "Site A",
                  "Site Z",
                  "Lag A",
                  "Lag Z",
                ]}
                rows={data.AssociatedUIDs}
                highlightUid={uid}
              />
              <Table
                title="GDCO Tickets"
                headers={[
                  "Ticket ID",
                  "Datacenter Code",
                  "Title",
                  "State",
                  "Assigned To",
                  "Ticket Link",
                ]}
                rows={data.GDCOTickets}
              />
            </Stack>
            <Stack horizontal tokens={{ childrenGap: 20 }}>
              <Table
                title="MGFX A-Side"
                headers={[
                  "XOMT",
                  "CO Device",
                  "CO Port",
                  "MO Device",
                  "MO Port",
                  "CO Diff",
                  "MO Diff",
                ]}
                rows={data.MGFXA?.map(({ Side, ...keep }: any) => keep)}
              />
              <Table
                title="MGFX Z-Side"
                headers={[
                  "XOMT",
                  "CO Device",
                  "CO Port",
                  "MO Device",
                  "MO Port",
                  "CO Diff",
                  "MO Diff",
                ]}
                rows={data.MGFXZ?.map(({ Side, ...keep }: any) => keep)}
              />
            </Stack>
          </>
        )}
      </Stack>
    </div>
  );
}
