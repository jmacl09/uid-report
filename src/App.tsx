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

      if (!history.includes(query)) setHistory([query, ...history]);
    } catch (err: any) {
      setError(err.message || "Network error occurred.");
    } finally {
      setLoading(false);
    }
  };

  const pad = (text: string, width: number) => {
    text = text == null ? "" : String(text);
    return text.padEnd(width, " ");
  };

  const copyTableText = (title: string, rows: Record<string, any>[], headers: string[]) => {
    if (!rows?.length) return;
    const colWidths = headers.map((h, i) =>
      Math.max(
        h.length,
        ...rows.map((r) => String(Object.values(r)[i] ?? "").length)
      ) + 2
    );

    let output = `${title}\n`;
    output += headers.map((h, i) => pad(h, colWidths[i])).join("") + "\n";
    output += "-".repeat(colWidths.reduce((a, b) => a + b, 0)) + "\n";

    for (const r of rows) {
      const vals = Object.values(r);
      output += vals.map((v, i) => pad(v, colWidths[i])).join("") + "\n";
    }

    navigator.clipboard.writeText(output.trimEnd());
    alert(`Copied ${title} as structured table ✅`);
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

  const Table = ({ title, headers, rows, highlightUid }: any) => {
    if (!rows?.length) return null;
    const scrollable: React.CSSProperties = {};
    if ((title === "GDCO Tickets" || title === "Associated UIDs") && rows.length > 5) {
      scrollable.maxHeight = 230;
      scrollable.overflowY = "auto";
    }

    return (
      <div className="table-container" style={scrollable}>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Text className="section-title">{title}</Text>
          <Stack horizontal tokens={{ childrenGap: 6 }}>
            <IconButton
              iconProps={{ iconName: "Copy" }}
              title="Copy Table"
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
              const highlight = highlightUid && row.Uid?.toString() === highlightUid;
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
          ⚡ Optical 360
        </Text>
        <Nav groups={navLinks} />
        <Separator />
        <Text className="footer">
          Built by <b>Josh Maclean</b> | Microsoft
        </Text>
      </div>

      <Stack className="main">
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Text className="portal-title">UID Lookup Portal</Text>
          <Stack horizontal tokens={{ childrenGap: 10 }}>
            <IconButton
              iconProps={{ iconName: "ExcelLogo" }}
              title="Export to Excel"
              className="excel-btn"
              onClick={exportExcel}
            />
          </Stack>
        </Stack>

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

        {error && (
          <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>
        )}

        {data && (
          <>
            <Stack tokens={{ childrenGap: 10 }}>
              <Stack
                horizontal
                horizontalAlign="space-between"
                verticalAlign="center"
                className="optical-summary-header"
              >
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                  <Text className="section-title">Optical Link Summary</Text>

                  <div className="side-buttons left">
                    <Text className="side-label">A Side:</Text>
                    <button
                      className="sleek-btn wan"
                      onClick={() => window.open(data?.AExpansions?.AUrl, "_blank")}
                    >
                      WAN Checker
                    </button>
                    <button
                      className="sleek-btn optical"
                      onClick={() => window.open(data?.AExpansions?.AOpticalUrl, "_blank")}
                    >
                      Optical Validator
                    </button>
                  </div>

                  <div className="side-buttons right">
                    <Text className="side-label">Z Side:</Text>
                    <button
                      className="sleek-btn wan"
                      onClick={() => window.open(data?.ZExpansions?.ZUrl, "_blank")}
                    >
                      WAN Checker
                    </button>
                    <button
                      className="sleek-btn optical"
                      onClick={() => window.open(data?.ZExpansions?.ZOpticalUrl, "_blank")}
                    >
                      Optical Validator
                    </button>
                  </div>
                </Stack>
              </Stack>

              <Table
                title=""
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
            </Stack>

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
                  "C0 Device",
                  "C0 Port",
                  "M0 Device",
                  "M0 Port",
                  "C0 Diff",
                  "M0 Diff",
                ]}
                rows={data.MGFXA?.map(({ Side, ...keep }: any) => keep)}
              />
              <Table
                title="MGFX Z-Side"
                headers={[
                  "XOMT",
                  "C0 Device",
                  "C0 Port",
                  "M0 Device",
                  "M0 Port",
                  "C0 Diff",
                  "M0 Diff",
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
