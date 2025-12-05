import React, { useEffect, useMemo, useState } from "react";
import { useNavigate } from "react-router-dom";
import {
  Stack,
  Text,
  DetailsList,
  IColumn,
  DetailsListLayoutMode,
  SelectionMode,
  Shimmer,
  ShimmerElementType,
  Dropdown,
  IDropdownOption,
  DatePicker,
  PrimaryButton,
  Icon,
  Separator,
  useTheme
} from "@fluentui/react";
import { LineChart, IChartProps } from "@fluentui/react-charting";
import { logAction } from "../api/log";
import "../Theme.css";

interface ActivityLogEntity {
  partitionKey: string;
  rowKey: string;
  email?: string;
  owner?: string;
  action?: string;
  title?: string;
  description?: string;
  category?: string;
  timestamp: string;
  savedAt?: string;
  metadata?: string;
}

const CET_LOCALE = "de-CH";
const CET_TZ = "Europe/Zurich";

const formatCET = (dateString: string) => {
  if (!dateString) return "-";
  try {
    return new Date(dateString).toLocaleString(CET_LOCALE, {
      timeZone: CET_TZ,
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit"
    });
  } catch {
    return dateString;
  }
};

const parseDetails = (details?: string) => {
  if (!details) return null;
  try {
    const parsed = JSON.parse(details);
    return typeof parsed === "object" ? parsed : details;
  } catch {
    return details;
  }
};

const renderParsedDetails = (obj: any) => {
  if (!obj || typeof obj !== "object") return null;
  return (
    <Stack tokens={{ childrenGap: 2 }}>
      {Object.entries(obj).map(([key, value]) => {
        let display: string;
        if (Array.isArray(value)) {
          display = value.length ? value.join(", ") : "None";
        } else if (value === "" || value === null || value === undefined) {
          display = "(empty)";
        } else {
          display = String(value);
        }
        return (
          <Text variant="xSmall" key={key}>
            <strong>{key}:</strong> {display}
          </Text>
        );
      })}
    </Stack>
  );
};

const ADMIN_EMAIL = "joshmaclean@microsoft.com";

const Logs: React.FC = () => {
  const [email, setEmail] = useState<string | null>(null);
  const [authorized, setAuthorized] = useState<boolean | null>(null);
  const [loadingUser, setLoadingUser] = useState(true);

  const [items, setItems] = useState<ActivityLogEntity[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string | null>(null);

  const [userFilter, setUserFilter] = useState<string | undefined>();
  const [actionFilter, setActionFilter] = useState<string | undefined>();
  const [fromDate, setFromDate] = useState<Date | null>(null);
  const [toDate, setToDate] = useState<Date | null>(null);

  const navigate = useNavigate();
  const theme = useTheme();

  /* -------------------------------------
     AUTH ↦ ADMIN ONLY
  -------------------------------------- */
  useEffect(() => {
    let cancelled = false;
    const fetchUser = async () => {
      try {
        const res = await fetch("/.auth/me", { credentials: "include" });
        const data = await res.json();
        const principal = data?.clientPrincipal;
        const userEmail = principal?.userDetails;

        if (!cancelled) {
          setEmail(userEmail || null);
          const isAdmin = userEmail?.toLowerCase() === ADMIN_EMAIL.toLowerCase();
          setAuthorized(isAdmin);
          setLoadingUser(false);

          if (isAdmin) {
            logAction(userEmail || "", "View Logs Page");
          } else {
            navigate("/", { replace: true });
          }
        }
      } catch {
        if (!cancelled) {
          setAuthorized(false);
          setLoadingUser(false);
          navigate("/", { replace: true });
        }
      }
    };
    fetchUser();
    return () => {
      cancelled = true;
    };
  }, [navigate]);

  /* -------------------------------------
     LOAD LOG DATA
  -------------------------------------- */
  useEffect(() => {
    if (!authorized) return;
    let cancelled = false;

    const load = async () => {
      setLoading(true);
      setError(null);

      try {
        const params: string[] = ["limit=500"];
        if (fromDate) params.push("dateFrom=" + fromDate.toISOString());
        if (toDate) {
          const end = new Date(toDate);
          end.setHours(23, 59, 59, 999);
          params.push("dateTo=" + end.toISOString());
        }

        const res = await fetch(`/api/log?${params.join("&")}`);
        const data = await res.json();

        const mapped = (data.items || []).map((e: any) => ({
          partitionKey: e.partitionKey,
          rowKey: e.rowKey,
          email: e.email || e.owner,
          owner: e.owner,
          action: e.action || e.title,
          title: e.title,
          description: e.description,
          category: e.category,
          timestamp: e.timestamp || e.Timestamp || e.savedAt,
          savedAt: e.savedAt,
          metadata: e.metadata || e.description
        }));

        if (!cancelled) setItems(mapped);
      } catch (err: any) {
        if (!cancelled) setError(err.message || "Failed to load logs");
      } finally {
        if (!cancelled) setLoading(false);
      }
    };

    load();
    return () => {
      cancelled = true;
    };
  }, [authorized, fromDate, toDate]);

  /* -------------------------------------
     FILTERED VIEW
  -------------------------------------- */
  const filteredItems = useMemo(() => {
    return items.filter((i) => {
      if (userFilter && i.email !== userFilter) return false;
      if (actionFilter && i.action !== actionFilter) return false;
      return true;
    });
  }, [items, userFilter, actionFilter]);

  const totalVisitsToday = useMemo(() => {
    const d = new Date();
    return filteredItems.filter((it) => {
      const t = new Date(it.timestamp);
      return (
        t.getFullYear() === d.getFullYear() &&
        t.getMonth() === d.getMonth() &&
        t.getDate() === d.getDate()
      );
    }).length;
  }, [filteredItems]);

  const uniqueUsers = useMemo(
    () => new Set(filteredItems.map((i) => i.email)).size,
    [filteredItems]
  );

  const totalActions = filteredItems.length;

  /* -------------------------------------
     TABLE COLUMNS
  -------------------------------------- */
  const columns: IColumn[] = [
    {
      key: "timestamp",
      name: "Time (CET)",
      fieldName: "timestamp",
      minWidth: 170,
      maxWidth: 220,
      onRender: (item) => <Text>{formatCET(item.timestamp)}</Text>
    },
    {
      key: "user",
      name: "User",
      fieldName: "email",
      minWidth: 180,
      onRender: (item) => <Text>{item.email || item.owner || "-"}</Text>
    },
    {
      key: "action",
      name: "Action",
      fieldName: "action",
      minWidth: 200,
      onRender: (item) => (
        <Text styles={{ root: { fontWeight: 600, color: "#e5e7eb" } }}>
          {item.action || item.title}
        </Text>
      )
    },
    {
      key: "category",
      name: "Category",
      fieldName: "category",
      minWidth: 140
    },
    {
      key: "details",
      name: "Details",
      fieldName: "description",
      minWidth: 260,
      onRender: (item) => {
        const parsed = parseDetails(item.metadata);
        if (!parsed) return <Text>(none)</Text>;
        if (typeof parsed === "string") return <Text>{parsed}</Text>;
        return renderParsedDetails(parsed);
      }
    }
  ];

  /* -------------------------------------
     CHART DATA
  -------------------------------------- */
  const chartData: IChartProps = useMemo(() => {
    const buckets = new Map<string, number>();
    for (const it of filteredItems) {
      const d = new Date(it.timestamp);
      const key = d.toISOString().substring(0, 10);
      buckets.set(key, (buckets.get(key) || 0) + 1);
    }

    const points = Array.from(buckets.entries())
      .sort(([a], [b]) => (a < b ? -1 : 1))
      .map(([key, count], idx) => ({
        x: idx + 1,
        y: count,
        xAxisCalloutData: key,
        yAxisCalloutData: `${count} actions`
      }));

    return {
      chartTitle: "Activity Over Time",
      lineChartData: [
        {
          legend: "Actions",
          data: points,
          color: "#3b82f6"
        }
      ]
    };
  }, [filteredItems]);

  /* -------------------------------------
     RENDER
  -------------------------------------- */

  if (loadingUser || authorized === null) {
    return (
      <div className="page-root">
        <Stack horizontalAlign="center">
          <Text variant="xLarge">Loading admin dashboard…</Text>
          <Shimmer />
        </Stack>
      </div>
    );
  }

  if (!authorized) return null;

  return (
    <div className="page-root" style={{ maxWidth: 1500, margin: "0 auto" }}>
      <Stack tokens={{ childrenGap: 28 }}>
        
        {/* HEADER */}
        <Stack horizontal horizontalAlign="space-between">
          <Stack>
            <Text variant="xxLarge" style={{ fontWeight: 700 }}>
              Activity Logs
            </Text>
            <Text variant="small" style={{ color: "#9ca3af" }}>
              Internal audit trail for all user interactions across Optical360.
            </Text>
          </Stack>

          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }}>
            <Icon iconName="Contact" styles={{ root: { color: "#3b82f6" } }} />
            <Text>{email}</Text>
          </Stack>
        </Stack>

        {/* METRICS */}
        <Stack horizontal wrap tokens={{ childrenGap: 16 }}>
          <Metric title="Total Visits Today" value={totalVisitsToday} />
          <Metric title="Unique Users" value={uniqueUsers} />
          <Metric title="Total Actions Logged" value={totalActions} />
        </Stack>

        {/* FILTER PANEL */}
        <Stack className="card-surface logs-filters" tokens={{ childrenGap: 16 }}>
          <Stack horizontal wrap tokens={{ childrenGap: 20 }} verticalAlign="end">
            <Dropdown
              label="User"
              options={Array.from(new Set(items.map(i => i.email).filter((v): v is string => !!v))).map(v => ({ key: v, text: v }))}
              placeholder="All users"
              selectedKey={userFilter}
              onChange={(_, opt) => setUserFilter(opt?.key as string)}
              styles={{ dropdown: { minWidth: 220 } }}
            />
            <Dropdown
              label="Action"
              options={Array.from(new Set(items.map(i => i.action).filter((v): v is string => !!v))).map(v => ({ key: v, text: v }))}
              placeholder="All actions"
              selectedKey={actionFilter}
              onChange={(_, opt) => setActionFilter(opt?.key as string)}
              styles={{ dropdown: { minWidth: 220 } }}
            />
            <DatePicker label="From" value={fromDate || undefined} onSelectDate={(d) => setFromDate(d || null)} />
            <DatePicker label="To" value={toDate || undefined} onSelectDate={(d) => setToDate(d || null)} />

            <PrimaryButton
              text="Refresh"
              iconProps={{ iconName: "Refresh" }}
              onClick={() => {
                setFromDate(fromDate ? new Date(fromDate) : null);
                setToDate(toDate ? new Date(toDate) : null);
              }}
            />
          </Stack>

          {error && <Text style={{ color: "red" }}>{error}</Text>}
        </Stack>

        {/* TIMELINE + CHART */}
        <Stack horizontal wrap tokens={{ childrenGap: 20 }}>
          
          {/* TIMELINE */}
          <Stack className="card-surface" style={{ minWidth: 340, maxWidth: 420 }} tokens={{ childrenGap: 8 }}>
            <SectionHeader icon="TimelineProgress" title="Recent Activity" />

            {loading ? (
              <Shimmer />
            ) : (
              <Stack tokens={{ childrenGap: 12 }}>
                {filteredItems.slice(0, 12).map((item) => (
                  <Stack key={item.rowKey} horizontal tokens={{ childrenGap: 12 }}>
                    <div className="timeline-dot" />
                    <Stack>
                      <Text variant="xSmall" style={{ color: "#9ca3af" }}>
                        {formatCET(item.timestamp)}
                      </Text>
                      <Text variant="small" style={{ fontWeight: 600 }}>
                        {item.action}
                      </Text>
                      <Text variant="xSmall" style={{ color: "#6b7280" }}>
                        {item.email}
                      </Text>

                      {/* Parsed details */}
                      {(() => {
                        const parsed = parseDetails(item.metadata);
                        if (parsed && typeof parsed === "object") {
                          return (
                            <Stack className="timeline-details">
                              {renderParsedDetails(parsed)}
                            </Stack>
                          );
                        }
                        return null;
                      })()}

                    </Stack>
                  </Stack>
                ))}
              </Stack>
            )}
          </Stack>

          {/* CHART */}
          <Stack className="card-surface" grow tokens={{ childrenGap: 8 }}>
            <SectionHeader icon="AreaChart" title="Activity Over Time" />

            {loading ? <Shimmer /> : (
              <LineChart
                data={chartData}
                height={260}
                hideLegend={false}
                wrapXAxisLables
              />
            )}
          </Stack>
        </Stack>

        {/* TABLE */}
        <Stack className="card-surface" tokens={{ childrenGap: 12 }}>
          <SectionHeader icon="Table" title="All Activity" />
          <Text variant="xSmall" style={{ color: "#9ca3af" }}>
            {filteredItems.length} entries
          </Text>

          {loading ? (
            <Shimmer />
          ) : (
            <DetailsList
              items={filteredItems}
              columns={columns}
              layoutMode={DetailsListLayoutMode.justified}
              selectionMode={SelectionMode.none}
              setKey="logsTable"
            />
          )}
        </Stack>

        <Separator />

        <Text variant="xSmall" style={{ color: "#9ca3af" }}>
          Activity logging is scoped to Optical360 internal usage and visible only to the platform admin.
        </Text>
      </Stack>
    </div>
  );
};


/* --------------------------------------------------
   REUSABLE UI COMPONENTS
-------------------------------------------------- */

const Metric = ({ title, value }: { title: string; value: number }) => (
  <Stack className="metric-card" tokens={{ childrenGap: 4 }}>
    <Text className="metric-label">{title}</Text>
    <Text variant="xxLarge" className="metric-value">{value}</Text>
  </Stack>
);

const SectionHeader = ({ icon, title }: { icon: string; title: string }) => (
  <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
    <Icon iconName={icon} styles={{ root: { color: "#3b82f6" } }} />
    <Text variant="mediumPlus" style={{ fontWeight: 600 }}>{title}</Text>
  </Stack>
);

export default Logs;
