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
import logo from "./assets/optical360-logo.png";

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
  const [lastSearched, setLastSearched] = useState<string>("");
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

    setUid(query); // ✅ ensure the input field updates when clicking a linked UID
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

      if (Array.isArray(result.AssociatedUIDs)) {
        result.AssociatedUIDs.sort((a: any, b: any) => {
          const uidA = String(a?.UID || a?.Uid || a?.uid || "");
          const uidB = String(b?.UID || b?.Uid || b?.uid || "");
          return uidB.localeCompare(uidA, undefined, { numeric: true });
        });
      }

      setData(result);
      setLastSearched(query);
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
      Math.max(h.length, ...rows.map((r) => String(Object.values(r)[i] ?? "").length)) + 2
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
      "Link Summary": data.OLSLinks,
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
    return (
      <div className="table-container">
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Text className="section-title">{title}</Text>
          <IconButton
            iconProps={{ iconName: "Copy" }}
            title="Copy Table"
            onClick={() => copyTableText(title, rows, headers)}
          />
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
              const highlight = highlightUid && row.Uid?.toString() === highlightUid;
              return (
                <tr key={i} className={highlight ? "highlight-row" : ""}>
                  {Object.entries(row).map(([key, val]: [string, any], j) => {
                    if (
                      key.toLowerCase().includes("workflow") ||
                      key.toLowerCase().includes("diff") ||
                      key.toLowerCase().includes("ticketlink")
                    ) {
                      return (
                        <td key={j}>
                          <button className="open-btn" onClick={() => window.open(val, "_blank")}>
                            Open
                          </button>
                        </td>
                      );
                    }
                    if (title === "Associated UIDs" && key.toLowerCase() === "uid") {
                      return (
                        <td key={j}>
                          <span
                            className="uid-click"
                            onClick={() => handleSearch(val)} // ✅ fixed: triggers correct UID search
                            title={`Search UID ${val}`}
                          >
                            {val}
                          </span>
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
      <div className="sidebar dark-nav">
        <img src={logo} alt="Optical 360 Logo" className="logo-img" />
        <Nav groups={navLinks} />
        <Separator />
        <Text className="footer">
          Built by <b>Josh Maclean</b> | Microsoft
        </Text>
      </div>

      <Stack className="main">
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <div />
          <IconButton
            iconProps={{ iconName: "ExcelLogo" }}
            title="Export to Excel"
            className="excel-btn"
            onClick={exportExcel}
          />
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
          {lastSearched && (
            <Text className="last-searched">
              Last searched:{" "}
              <span className="uid-click" onClick={() => handleSearch(lastSearched)}>
                {lastSearched}
              </span>
            </Text>
          )}
          {loading && <Spinner size={SpinnerSize.large} label="Fetching data..." />}
        </Stack>

        <div className="last-searched-gap" />

        {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}

        {data && (
          <>
            {/* Details Section */}
            <div className="table-container details-fit">
              <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                <Text className="section-title">Details</Text>
              </Stack>
              <table className="data-table details-table">
                <thead>
                  <tr>
                    <th>SRLGID</th>
                    <th>SRLG</th>
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    <td>{data?.AExpansions?.SRLGID}</td>
                    <td>{data?.AExpansions?.SRLG}</td>
                  </tr>
                </tbody>
              </table>
              <div className="details-buttons">
                <button
                  className="sleek-btn repo"
                  onClick={() => window.open(data?.AExpansions?.DocumentRepository, "_blank")}
                >
                  WAN Capacity Repository
                </button>
                <button
                  className="sleek-btn kmz"
                  onClick={() =>
                    window.open(
                      `https://fiberplanner.cloudg.is/?srlg=${data?.AExpansions?.SrlgId}`,
                      "_blank"
                    )
                  }
                >
                  {data?.AExpansions?.DCLocation || "Open"} KMZ Route
                </button>
              </div>
            </div>

            {/* Added spacing before WAN Buttons ✅ */}
            <div style={{ marginTop: "12px" }} />

            {/* WAN Buttons */}
            <div className="button-header-align-left">
              <div className="side-buttons">
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
                <Text className="side-label" style={{ marginLeft: "40px" }}>
                  Z Side:
                </Text>
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
            </div>

            {/* Tables */}
            <Table
              title="Link Summary"
              headers={[
                "A Device",
                "A Port",
                "Z Device",
                "Z Port",
                "A Optical Device",
                "A Optical Port",
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
                  "SrlgId",
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
                headers={["Ticket Id", "DC Code", "Title", "State", "Assigned To", "Link"]}
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
                  "C0 DIFF",
                  "M0 DIFF",
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
                  "C0 DIFF",
                  "M0 DIFF",
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
