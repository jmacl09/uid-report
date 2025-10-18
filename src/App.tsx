import React, { useState } from "react";
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

      setData(result);
      setSummary(
        `Found ${result.OLSLinks?.length || 0} active optical paths, ${
          result.AssociatedUIDs?.length || 0
        } associated UIDs, and ${result.GDCOTickets?.length || 0} related GDCO tickets.`
      );
    } catch (err: any) {
      setError(err.message || "Network error occurred.");
      setSummary("Error retrieving data.");
    } finally {
      setLoading(false);
    }
  };

  const Table = ({ title, headers, rows }: any) => {
    if (!rows?.length) return null;
    return (
      <div className="table-container">
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Text className="section-title">{title}</Text>
          <Stack horizontal tokens={{ childrenGap: 6 }}>
            <IconButton
              iconProps={{ iconName: "Copy" }}
              title="Copy JSON"
              onClick={() =>
                navigator.clipboard.writeText(JSON.stringify(rows, null, 2))
              }
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
            {rows.map((row: any, i: number) => (
              <tr key={i}>
                {headers.map((h: string, j: number) => {
                  const key = Object.keys(row)[j];
                  const value = row[key];
                  if (
                    key.toLowerCase().includes("workflow") ||
                    key.toLowerCase().includes("diff") ||
                    key.toLowerCase().includes("ticketlink")
                  ) {
                    return (
                      <td key={j}>
                        <a
                          href={value}
                          target="_blank"
                          rel="noopener noreferrer"
                        >
                          Open
                        </a>
                      </td>
                    );
                  }
                  return <td key={j}>{value}</td>;
                })}
              </tr>
            ))}
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
        <Text className="portal-title">UID Lookup Portal</Text>

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
              onClick={handleSearch}
              className="search-btn"
            />
          </Stack>
          {loading && <Spinner size={SpinnerSize.large} label="Fetching data..." />}
        </Stack>

        <div className="summary">{summary}</div>

        {error && (
          <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>
        )}

        {data && (
          <>
            <Table
              title="OLS Optical Link Summary"
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
              />
              <Table
                title="GDCO Tickets"
                headers={[
                  "Ticket ID",
                  "Datacenter Code",
                  "Clean Title",
                  "State",
                  "Clean Assigned To",
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
                rows={data.MGFXA.map(
                  ({ Side, ...keep }: Record<string, any>) => keep
                )}
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
                rows={data.MGFXZ.map(
                  ({ Side, ...keep }: Record<string, any>) => keep
                )}
              />
            </Stack>
          </>
        )}
      </Stack>
    </div>
  );
}
