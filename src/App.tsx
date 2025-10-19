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
  const [lastUid, setLastUid] = useState<string | null>(null);
  const [loading, setLoading] = useState<boolean>(false);
  const [data, setData] = useState<any>(null);
  const [error, setError] = useState<string | null>(null);
  const [inputError, setInputError] = useState<string | null>(null);

  const handleSearch = async (searchUid?: string) => {
    const query = searchUid || uid.trim();
    if (!query) return setInputError("Please enter a UID before searching.");
    if (!/^\d{11}$/.test(query)) {
      setInputError("UID must contain exactly 11 digits (numbers only).");
      return;
    }

    setInputError(null);
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
      setData(result);
      setLastUid(query);
    } catch (err: any) {
      setError(err.message || "Network error occurred.");
    } finally {
      setLoading(false);
    }
  };

  const exportExcel = () => {
    if (!data || !uid) return;
    const wb = XLSX.utils.book_new();
    const sections = {
      Details: [
        {
          SRLGID: data?.AExpansions?.SRLGID,
          SRLG: data?.AExpansions?.SRLG,
          DocumentRepository: data?.AExpansions?.DocumentRepository,
        },
      ],
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

  const copyTableText = (title: string, rows: Record<string, any>[], headers: string[]) => {
    if (!rows?.length) return;
    const colWidths = headers.map(
      (h, i) => Math.max(h.length, ...rows.map((r) => String(Object.values(r)[i] ?? "").length)) + 2
    );
    const pad = (text: string, width: number) => (text ?? "").toString().padEnd(width, " ");
    let output = `${title}\n${headers.map((h, i) => pad(h, colWidths[i])).join("")}\n`;
    output += "-".repeat(colWidths.reduce((a, b) => a + b, 0)) + "\n";
    for (const r of rows)
      output += Object.values(r)
        .map((v, i) => pad(v, colWidths[i]))
        .join("") + "\n";
    navigator.clipboard.writeText(output.trimEnd());
    alert(`Copied ${title} as structured table ✅`);
  };

  const Table = ({ title, headers, rows, highlightUid, noScroll }: any) => {
    if (!rows?.length) return null;
    const wrapperStyle = noScroll ? { overflow: "hidden" } : {};
    return (
      <div className="table-container" style={wrapperStyle}>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Text className="section-title">{title}</Text>
          <IconButton
            iconProps={{ iconName: "Copy" }}
            title="Copy Table"
            onClick={() => copyTableText(title, rows, headers)}
          />
        </Stack>
        <table className="data-table" style={{ width: "100%" }}>
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
                  {Object.entries(row).map(([key, val], j) => {
                    if (typeof val === "object" && val !== null) val = "";
                    if (key.toLowerCase().includes("workflow")) {
                      return (
                        <td key={j}>
                          <button
                            className="open-btn"
                            onClick={() => window.open(String(val), "_blank")}
                          >
                            Open
                          </button>
                        </td>
                      );
                    }
                    if (title === "Associated UIDs" && key.toLowerCase() === "uid") {
                      return (
                        <td key={j}>
                          <button
                            className="link-btn-noline"
                            onClick={() => {
                              const newUid = String(val);
                              setUid(newUid);
                              handleSearch(newUid);
                            }}
                          >
                            {String(val)}
                          </button>
                        </td>
                      );
                    }
                    return <td key={j}>{String(val)}</td>;
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
      {/* Sidebar */}
      <div className="sidebar">
        <img src={logo} alt="Optical 360 Logo" className="logo-img" />
        <Nav
          groups={navLinks}
          styles={{
            root: { background: "transparent" },
            link: {
              color: "#ddd",
              fontSize: 14,
              borderRadius: 6,
              margin: "2px 0",
              selectors: {
                ":hover": {
                  backgroundColor: "#1b1b1b",
                  color: "#50b3ff",
                },
              },
            },
          }}
        />
        <Separator />
        <Text className="footer">
          Built by <b>Josh Maclean</b> | Microsoft
        </Text>
      </div>

      {/* Main Content */}
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

        {/* Search Input */}
        <Stack horizontalAlign="center" tokens={{ childrenGap: 10 }}>
          <Stack horizontal tokens={{ childrenGap: 10 }}>
            <TextField
              placeholder="Enter UID (11 digits)"
              value={uid}
              onChange={(_e, v) => setUid(v ?? "")}
              className="input-field"
              errorMessage={inputError || undefined}
            />
            <PrimaryButton
              text={loading ? "Loading..." : "Search"}
              disabled={loading}
              onClick={() => handleSearch()}
              className="search-btn"
            />
          </Stack>

          {/* Clickable last searched UID */}
          {lastUid && (
            <Text
              className="last-uid-link"
              onClick={() => handleSearch(lastUid)}
            >
              Last searched UID: {lastUid}
            </Text>
          )}

          {loading && <Spinner size={SpinnerSize.large} label="Fetching data..." />}
        </Stack>

        <div className="table-spacing" />

        {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}

        {data && (
          <>
            {/* Details (No scroll, fixed width) */}
            <Table
              title="Details"
              noScroll
              headers={["SRLG ID", "SRLG", "Repository", "KMZ Route"]}
              rows={[
                {
                  "SRLG ID": data?.AExpansions?.SRLGID || "N/A",
                  SRLG: data?.AExpansions?.SRLG || "N/A",
                  Repository: (
                    <button
                      className="sleek-btn green"
                      onClick={() =>
                        window.open(String(data?.AExpansions?.DocumentRepository), "_blank")
                      }
                    >
                      WAN Capacity Repository
                    </button>
                  ),
                  "KMZ Route": (
                    <button
                      className="sleek-btn green"
                      onClick={() =>
                        window.open(
                          `https://fiberplanner.cloudg.is/?srlg=${encodeURIComponent(
                            data?.AExpansions?.SRLGID || ""
                          )}`,
                          "_blank"
                        )
                      }
                    >
                      {data?.AExpansions?.SRLGID
                        ? `${data?.AExpansions?.SRLGID} KMZ Route`
                        : "Open KMZ Route"}
                    </button>
                  ),
                },
              ]}
            />

            {/* A/Z Side Buttons */}
            <div className="button-header-align-left">
              <div className="side-buttons">
                <Text className="side-label">A Side:</Text>
                <button
                  className="sleek-btn wan"
                  onClick={() => window.open(String(data?.AExpansions?.AUrl), "_blank")}
                >
                  WAN Checker
                </button>
                <button
                  className="sleek-btn optical"
                  onClick={() =>
                    window.open(String(data?.AExpansions?.AOpticalUrl), "_blank")
                  }
                >
                  Optical Validator
                </button>

                <Text className="side-label" style={{ marginLeft: "40px" }}>
                  Z Side:
                </Text>
                <button
                  className="sleek-btn wan"
                  onClick={() => window.open(String(data?.ZExpansions?.ZUrl), "_blank")}
                >
                  WAN Checker
                </button>
                <button
                  className="sleek-btn optical"
                  onClick={() =>
                    window.open(String(data?.ZExpansions?.ZOpticalUrl), "_blank")
                  }
                >
                  Optical Validator
                </button>
              </div>
            </div>

            {/* Link Summary */}
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

            {/* Associated UIDs + GDCO Tickets */}
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

            {/* MGFX A/Z Sides */}
            <Stack horizontal tokens={{ childrenGap: 20 }}>
              <Table
                title="MGFX A-Side"
                headers={["XOMT", "C0 Device", "C0 Port", "M0 Device", "M0 Port", "C0 DIFF", "M0 DIFF"]}
                rows={data.MGFXA?.map(({ Side, ...keep }: any) => keep)}
              />
              <Table
                title="MGFX Z-Side"
                headers={["XOMT", "C0 Device", "C0 Port", "M0 Device", "M0 Port", "C0 DIFF", "M0 DIFF"]}
                rows={data.MGFXZ?.map(({ Side, ...keep }: any) => keep)}
              />
            </Stack>
          </>
        )}
      </Stack>
    </div>
  );
}
