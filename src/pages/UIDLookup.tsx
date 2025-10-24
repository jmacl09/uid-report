import React, { useState, useEffect } from "react";
import { useLocation, useNavigate } from "react-router-dom";
import {
  initializeIcons,
  Stack,
  Text,
  TextField,
  PrimaryButton,
  IconButton,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
} from "@fluentui/react";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";

export default function UIDLookup() {
  initializeIcons();
  const location = useLocation();
  const navigate = useNavigate();

  const [uid, setUid] = useState<string>("");
  const [data, setData] = useState<any>(null);
  const [loading, setLoading] = useState<boolean>(false);
  const [history, setHistory] = useState<string[]>(() => {
    try {
      const raw = localStorage.getItem("uidHistory");
      return raw ? JSON.parse(raw) : [];
    } catch {
      return [];
    }
  });
  const [lastSearched, setLastSearched] = useState<string>("");
  const [error, setError] = useState<string | null>(null);
  useEffect(() => {
    localStorage.setItem("uidHistory", JSON.stringify(history.slice(0, 10)));
  }, [history]);

  // Reset to landing view when sidebar forces a reset param
  useEffect(() => {
    const params = new URLSearchParams(location.search);
    if (params.has("reset")) {
      setUid("");
      setLastSearched("");
      setData(null);
      setError(null);
      setLoading(false);
      navigate("/uid", { replace: true });
    }
  }, [location.search, navigate]);

  const naturalSort = (a: string, b: string) =>
    a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" });

  const handleSearch = async (searchUid?: string) => {
    const query = (searchUid || uid || "").toString();
    if (!/^\d{11}$/.test(query)) {
      setError("Invalid UID. It must contain exactly 11 numbers.");
      return;
    }
    
    setUid(query);
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

  const formatTableText = (
    title: string,
    rows: Record<string, any>[] | undefined,
    headers: string[]
  ): string => {
    if (!rows || !rows.length) return "";
    const colWidths = headers.map((h, i) =>
      Math.max(h.length, ...rows.map((r) => String(Object.values(r)[i] ?? "").length)) + 2
    );
    let out = `${title}\n`;
    out += headers.map((h, i) => pad(h, colWidths[i])).join("") + "\n";
    out += "-".repeat(colWidths.reduce((a, b) => a + b, 0)) + "\n";
    for (const r of rows) {
      const vals = Object.values(r);
      out += vals.map((v, i) => pad(String(v ?? ""), colWidths[i])).join("") + "\n";
    }
    return out.trimEnd() + "\n\n";
  };

  // Build full plain‑text export of all sections
  const buildAllText = (): string => {
    if (!data) return "";
    let text = "";

    // Details
    try {
      const detailsHeaders = ["SRLGID", "SRLG"];
      const detailsRows = [
        { SRLGID: String(data?.AExpansions?.SRLGID ?? ""), SRLG: String(data?.AExpansions?.SRLG ?? "") },
      ].map((r) => Object.values(r).reduce((acc: any, v: any, i: number) => ({ ...acc, [detailsHeaders[i]]: v }), {}));
      text += formatTableText("Details", detailsRows as any, detailsHeaders);
    } catch {}

    // Link Summary
    text += formatTableText(
      "Link Summary",
      data.OLSLinks,
      [
        "A Device",
        "A Port",
        "Z Device",
        "Z Port",
        "A Optical Device",
        "A Optical Port",
        "Z Optical Device",
        "Z Optical Port",
        "Workflow",
      ]
    );

    // Associated UIDs
    text += formatTableText(
      "Associated UIDs",
      data.AssociatedUIDs,
      ["UID", "SrlgId", "Action", "Type", "Device A", "Device Z", "Site A", "Site Z", "Lag A", "Lag Z"]
    );

    // GDCO Tickets
    text += formatTableText("GDCO Tickets", data.GDCOTickets, ["Ticket Id", "DC Code", "Title", "State", "Assigned To", "Link"]);

    // MGFX A/Z (remove Side key if present)
    const mgfxA = (data.MGFXA || []).map(({ Side, ...keep }: any) => keep);
    const mgfxZ = (data.MGFXZ || []).map(({ Side, ...keep }: any) => keep);
    const mgfxHeaders = ["XOMT", "C0 Device", "C0 Port", "M0 Device", "M0 Port", "C0 DIFF", "M0 DIFF"];
    text += formatTableText("MGFX A-Side", mgfxA, mgfxHeaders);
    text += formatTableText("MGFX Z-Side", mgfxZ, mgfxHeaders);

    return text.trimEnd();
  };

  const copyAll = async () => {
    const text = buildAllText();
    if (!text) return;
    await navigator.clipboard.writeText(text);
    alert("All sections copied to clipboard as plain text.");
  };

  const exportOneNote = async () => {
    const text = buildAllText();
    if (text) {
      try { await navigator.clipboard.writeText(text); } catch {}
    }
    // Open OneNote (web quick note). Content is on clipboard for immediate paste; no alerts shown.
    // If the Windows app is registered, this deep link may open it on some systems:
    // window.location.href = 'onenote:';
    window.open("https://www.onenote.com/quicknote?auth=1", "_blank");
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
                            onClick={() => handleSearch(val)}
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

  const isInitialView = !lastSearched && !loading && !data;

  return (
    <Stack className="main">
      {isInitialView ? (
      <div className="vso-form-container glow" style={{ width: "80%", maxWidth: 800 }}>
        <div className="banner-title">
          <span className="title-text">UID Lookup</span>
          <span className="title-sub">The Ulitimate UID Lookup Tool</span>
        </div>

        <div style={{ display: "flex", gap: 10, alignItems: "center", justifyContent: "center" }}>
          <TextField
            placeholder="Enter UID (e.g., 20190610163)"
            value={uid}
            onChange={(_e, v) => {
              const cleaned = (v ?? "").replace(/\D/g, "").slice(0, 11);
              setUid(cleaned);
            }}
            className="input-field"
            inputMode="numeric"
            pattern="[0-9]*"
          />
          <PrimaryButton
            text="Search"
            disabled={loading}
            onClick={() => handleSearch()}
            className="search-btn"
            style={{ marginLeft: 20 }}
          />
        </div>

        <div style={{ marginTop: 8, textAlign: "center", fontSize: 12, color: "#aaa" }}>
          First time here?{' '}
          <span className="uid-click" onClick={() => handleSearch('20190610161')}>
            Try now
          </span>
        </div>

        {lastSearched && (
          <Text className="last-searched" style={{ marginTop: 6 }}>
            Last searched:{" "}
            <span className="uid-click" onClick={() => handleSearch(lastSearched)}>
              {lastSearched}
            </span>
          </Text>
        )}

        {history.length > 0 && (
          <div style={{ marginTop: 6, color: "#aaa", fontSize: 12 }}>
            Recent: {history.slice(0, 5).map((h, i) => (
              <span
                key={h}
                className="uid-click"
                style={{ marginLeft: i === 0 ? 0 : 10 }}
                onClick={() => handleSearch(h)}
              >
                {h}
              </span>
            ))}
          </div>
        )}


      </div>
      ) : (
        <div className="uid-compact-bar" style={{ margin: "4px 0 10px", display: "flex", justifyContent: "center" }}>
          <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
            <div style={{ width: 220 }}>
              <TextField
                placeholder="Enter UID (e.g., 20190610163)"
                value={uid}
                onChange={(_e, v) => {
                  const cleaned = (v ?? "").replace(/\D/g, "").slice(0, 11);
                  setUid(cleaned);
                }}
                className="input-field"
                inputMode="numeric"
                pattern="[0-9]*"
                styles={{ fieldGroup: { width: 220 } }}
              />
            </div>
            <PrimaryButton
              text="Search"
              disabled={loading}
              onClick={() => handleSearch()}
              className="search-btn"
              style={{ marginLeft: 20 }}
            />
          </div>
        </div>
      )}

      {loading && (
        <div style={{ textAlign: "center", margin: '6px 0 12px' }}>
          <Spinner size={SpinnerSize.large} label="Fetching data..." />
        </div>
      )}

      <div className="last-searched-gap" />

      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}

      {data && (
        <>
          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, marginBottom: 8 }}>
            <IconButton iconProps={{ iconName: 'Copy' }} title="Copy All" ariaLabel="Copy All" onClick={copyAll} />
            <IconButton iconProps={{ iconName: 'ExcelLogo' }} title="Export to Excel" ariaLabel="Export to Excel" onClick={exportExcel} />
            <IconButton iconProps={{ iconName: 'OneNoteLogo' }} title="Export to OneNote" ariaLabel="Export to OneNote" onClick={exportOneNote} />
          </div>
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
                  <th>Repository</th>
                  <th>Fiber Planner</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td>{data?.AExpansions?.SRLGID}</td>
                  <td>{data?.AExpansions?.SRLG}</td>
                  <td>
                    <button
                      className="sleek-btn repo"
                      onClick={() => window.open(data?.AExpansions?.DocumentRepository, "_blank")}
                    >
                      WAN Capacity Repository
                    </button>
                  </td>
                  <td>
                    <button
                      className="sleek-btn kmz"
                      onClick={() =>
                        window.open(
                          `https://fiberplanner.cloudg.is/?srlg=${data?.AExpansions?.SRLGID}`,
                          "_blank"
                        )
                      }
                    >
                      {data?.AExpansions?.DCLocation || "Open"} KMZ Route
                    </button>
                  </td>
                </tr>
              </tbody>
            </table>
          </div>

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

          <Stack horizontal tokens={{ childrenGap: 20 }} styles={{ root: { width: '100%' } }}>
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

          <Stack horizontal tokens={{ childrenGap: 20 }} styles={{ root: { width: '100%' } }}>
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
  );
}










