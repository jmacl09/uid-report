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
  IColumn
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

  const naturalSort = (a: string, b: string) =>
    a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" });

  const buildColumns = (objArray: any[]): IColumn[] => {
    if (!objArray?.length) return [];
    return Object.keys(objArray[0]).map((key) => ({
      key,
      name: key,
      fieldName: key,
      minWidth: 80,
      maxWidth: 160,
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
        return <span>{val}</span>;
      },
    }));
  };

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

      result.OLSLinks?.sort((a: any, b: any) => naturalSort(a.APort, b.APort));
      result.AssociatedUIDs?.sort(
        (a: any, b: any) => parseInt(b.Uid) - parseInt(a.Uid)
      );

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

  const Section = ({ title, rows, highlightUid }: any) => {
    if (!rows?.length) return null;

    return (
      <div className="table-section">
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

        <DetailsList
          items={rows}
          columns={buildColumns(rows)}
          compact={true}
          layoutMode={DetailsListLayoutMode.fixedColumns}
          styles={{
            root: { background: "#181818", borderRadius: 4, paddingTop: 4 },
            contentWrapper: { overflowX: "hidden" },
          }}
        />
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
            <Section title="OLS Optical Link Summary" rows={data.OLSLinks} />
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
