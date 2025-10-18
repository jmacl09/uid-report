import React, { useState, useEffect } from "react";
import {
  initializeIcons,
  Stack,
  Text,
  TextField,
  PrimaryButton,
  IconButton,
} from "@fluentui/react";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";

initializeIcons();

export default function App() {
  const [uid, setUid] = useState<string>("");
  const [data, setData] = useState<any>(null);
  const [summary, setSummary] = useState<string>("Awaiting UID lookup...");
  const [uidHistory, setUidHistory] = useState<string[]>([]);

  useEffect(() => {
    const history = JSON.parse(localStorage.getItem("uidHistory") || "[]");
    setUidHistory(history);
  }, []);

  const addToHistory = (uidVal: string) => {
    let history = JSON.parse(localStorage.getItem("uidHistory") || "[]");
    if (!history.includes(uidVal)) {
      history.unshift(uidVal);
      history = history.slice(0, 10);
      localStorage.setItem("uidHistory", JSON.stringify(history));
      setUidHistory(history);
    }
  };

  const handleSearch = async () => {
    if (!uid.trim()) {
      alert("Please enter a UID before searching.");
      return;
    }

    setSummary("Fetching data...");
    setData(null);

    const triggerUrl = `https://fibertools-dsavavdcfdgnh2cm.westeurope-01.azurewebsites.net/api/fiberflow/triggers/When_an_HTTP_request_is_received/invoke?api-version=2022-05-01&sp=%2Ftriggers%2FWhen_an_HTTP_request_is_received%2Frun&sv=1.0&sig=8KqIymphhOqUAlnd7UGwLRaxP0ot5ZH30b7jWCEUedQ&UID=${encodeURIComponent(
      uid
    )}`;

    try {
      const res = await fetch(triggerUrl);
      const result = await res.json();

      // Sort OLS by APort ascending
      result.OLSLinks?.sort((a: any, b: any) =>
        a.APort.localeCompare(b.APort, undefined, { numeric: true })
      );

      // Sort MGFX A/Z by XOMT numeric order
      const sortXOMT = (arr: any[]) =>
        arr?.sort(
          (a, b) =>
            parseInt(a.XOMT.replace(/\D/g, "")) -
            parseInt(b.XOMT.replace(/\D/g, ""))
        );

      sortXOMT(result.MGFXA);
      sortXOMT(result.MGFXZ);

      setData(result);
      setSummary(
        `Found ${result.OLSLinks?.length || 0} active optical paths, ${
          result.AssociatedUIDs?.length || 0
        } associated UIDs, and ${
          result.GDCOTickets?.length || 0
        } related GDCO tickets.`
      );
      addToHistory(uid);
    } catch (err) {
      setSummary("Error retrieving data.");
    }
  };

  const openLinks = (url: string) => (
    <a
      href={url}
      target="_blank"
      rel="noopener noreferrer"
      className="open-btn"
    >
      Open
    </a>
  );

  // ---------- Export Helpers ----------
  const copyHTML = () => {
    if (!data) return;
    const tables = document.querySelectorAll(".data-table");
    let html = "";
    tables.forEach((tbl) => (html += tbl.outerHTML));
    navigator.clipboard.writeText(html);
    alert("All tables copied as formatted HTML ✅");
  };

  const exportExcel = () => {
    if (!data || !uid) return;
    const wb = XLSX.utils.book_new();

    const addSheet = (rows: any[], name: string) => {
      if (!rows || !rows.length) return;
      const ws = XLSX.utils.json_to_sheet(rows);
      XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
    };

    addSheet(data.OLSLinks, "OLS Optical Summary");
    addSheet(data.AssociatedUIDs, "Associated UIDs");
    addSheet(data.GDCOTickets, "GDCO Tickets");
    addSheet(data.MGFXA, "MGFX A-Side");
    addSheet(data.MGFXZ, "MGFX Z-Side");

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
      </head>
      <body style='font-family:Segoe UI;background:#1b1b1b;color:#fff;'>
        <h1 style='color:#50b3ff;'>UID Report ${uid}</h1>
    `;
    tables.forEach((tbl: any) => (html += tbl.outerHTML));
    html += "</body></html>";

    const blob = new Blob([html], { type: "text/html" });
    saveAs(blob, `UID_Report_${uid}.html`);
  };

  // ---------- Table Renderer ----------
  const renderTable = (title: string, rows: any[], highlightField?: string) => {
    if (!rows?.length) return null;

    const headers = Object.keys(rows[0]);

    return (
      <div className="section">
        <Stack
          horizontal
          horizontalAlign="space-between"
          verticalAlign="center"
          className="section-header"
        >
          <Text className="section-title">{title}</Text>
          <IconButton
            iconProps={{ iconName: "Copy" }}
            title="Copy Table"
            onClick={() =>
              navigator.clipboard.writeText(JSON.stringify(rows, null, 2))
            }
          />
        </Stack>
        <table className="data-table">
          <thead>
            <tr>
              {headers.map((h) => (
                <th key={h}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((r, i) => (
              <tr
                key={i}
                className={
                  highlightField && r[highlightField] === uid ? "highlight" : ""
                }
              >
                {headers.map((h) => (
                  <td key={h}>
                    {typeof r[h] === "string" && r[h].startsWith("http") ? (
                      openLinks(r[h])
                    ) : (
                      <span>{r[h]}</span>
                    )}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  };

  return (
    <div className="app-container">
      {/* Sidebar */}
      <div className="sidebar">
        <Text variant="xLarge" className="logo">
          ⚡ FiberTools
        </Text>
        <div className="footer">
          Built by <b>Josh Maclean</b> | Microsoft <br />
          <span>©2025</span>
        </div>
      </div>

      {/* Main */}
      <div className="main">
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Text className="main-title">UID Lookup Portal</Text>

          {/* Action Buttons */}
          <Stack horizontal tokens={{ childrenGap: 8 }}>
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
            <IconButton
              iconProps={{ iconName: "Copy" }}
              title="Copy HTML"
              className="copy-btn"
              onClick={copyHTML}
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
              className="uid-input"
            />
            <PrimaryButton text="Search" onClick={handleSearch} />
          </Stack>

          {uidHistory.length > 0 && (
            <div className="uid-history">
              Recent:{" "}
              {uidHistory.map((u, i) => (
                <span key={i} onClick={() => setUid(u)}>
                  {u}
                </span>
              ))}
            </div>
          )}
          <Text className="summary-text">{summary}</Text>
        </Stack>

        {/* Tables */}
        {data && (
          <>
            {renderTable("OLS Optical Link Summary", data.OLSLinks)}
            <Stack horizontal tokens={{ childrenGap: 20 }}>
              <Stack grow>
                {renderTable("Associated UIDs", data.AssociatedUIDs, "Uid")}
              </Stack>
              <Stack grow>
                {renderTable("GDCO Tickets", data.GDCOTickets)}
              </Stack>
            </Stack>
            <Stack horizontal tokens={{ childrenGap: 20 }}>
              <Stack grow>{renderTable("MGFX A-Side", data.MGFXA)}</Stack>
              <Stack grow>{renderTable("MGFX Z-Side", data.MGFXZ)}</Stack>
            </Stack>
          </>
        )}
      </div>
    </div>
  );
}
