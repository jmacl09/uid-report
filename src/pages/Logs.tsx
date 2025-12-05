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
  useTheme,
  IDetailsRowProps,
  DetailsRow,
  IRenderFunction
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
    if (parsed && typeof parsed === "object") {
      return parsed;
    }
    return details;
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
          display = value.length === 0 ? "None" : value.join(", ");
        } else if (value === "") {
          display = "(empty)";
        } else if (value === null || value === undefined) {
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

  // Fetch signed-in user via SWA auth
  useEffect(() => {
    let cancelled = false;
    const fetchUser = async () => {
      try {
        const res = await fetch("/.auth/me", { credentials: "include" });
        if (!res.ok) throw new Error("Unable to read auth state");
        const data = await res.json();
        const principal = Array.isArray(data) ? null : data?.clientPrincipal;
        const userEmail: string | undefined = principal?.userDetails;
        if (!cancelled) {
          setEmail(userEmail || null);
          const isAdmin = !!userEmail && userEmail.toLowerCase() === ADMIN_EMAIL.toLowerCase();
          setAuthorized(isAdmin);
          setLoadingUser(false);
          if (userEmail) {
            try {
              localStorage.setItem("loggedInEmail", userEmail);
              window.dispatchEvent(new CustomEvent("loggedInEmailChanged", { detail: userEmail }));
            } catch {}
          }
          if (!isAdmin) {
            navigate("/", { replace: true });
          } else {
            logAction(userEmail || "", "View Logs Page");
          }
        }
      } catch {
        if (!cancelled) {
          setLoadingUser(false);
          setAuthorized(false);
          navigate("/", { replace: true });
        }
      }
    };
    fetchUser();
    return () => {
      cancelled = true;
    };
  }, [navigate]);

  // Load activity log data
  useEffect(() => {
    if (!authorized) return;
    let cancelled = false;
    const load = async () => {
      setLoading(true);
      setError(null);
      try {
        const params: string[] = ["limit=500"];
        if (fromDate) {
          params.push("dateFrom=" + encodeURIComponent(fromDate.toISOString()));
        }
        if (toDate) {
          const end = new Date(toDate.getTime());
          end.setHours(23, 59, 59, 999);
          params.push("dateTo=" + encodeURIComponent(end.toISOString()));
        }
        const query = params.length ? `?${params.join("&")}` : "";
        const res = await fetch(`/api/log${query}`);
        if (!res.ok) throw new Error("Failed to load logs");
        const data = await res.json();
        const raw: any[] = data?.items || [];
        const mapped: ActivityLogEntity[] = raw.map((e) => ({
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
        if (!cancelled) {
          setItems(mapped);
        }
      } catch (err: any) {
        if (!cancelled) {
          setError(err?.message || "Failed to load logs");
        }
      } finally {
        if (!cancelled) {
          setLoading(false);
        }
      }
    };
    load();
    return () => {
      cancelled = true;
    };
  }, [authorized, fromDate, toDate]);

  const filteredItems = useMemo(() => {
    return items.filter((it) => {
      if (userFilter && it.email !== userFilter) return false;
      if (actionFilter && it.action !== actionFilter) return false;
      return true;
    });
  }, [items, userFilter, actionFilter]);

  const totalVisitsToday = useMemo(() => {
    const today = new Date();
    const y = today.getFullYear();
    const m = today.getMonth();
    const d = today.getDate();
    return filteredItems.filter((it) => {
      const dt = new Date(it.timestamp);
      return dt.getFullYear() === y && dt.getMonth() === m && dt.getDate() === d;
    }).length;
  }, [filteredItems]);

  const uniqueUsers = useMemo(() => {
    const set = new Set(filteredItems.map((i) => i.email));
    return set.size;
  }, [filteredItems]);

  const totalActions = filteredItems.length;

  const mostUsedFeature = useMemo(() => {
    if (!filteredItems.length) return "-";
    const counts = new Map<string, number>();
    for (const it of filteredItems) {
      const key = it.action || "Unknown";
      counts.set(key, (counts.get(key) || 0) + 1);
    }
    let maxKey = "-";
    let maxVal = 0;
    counts.forEach((v, k) => {
      if (v > maxVal) {
        maxVal = v;
        maxKey = k;
      }
    });
    return maxKey;
  }, [filteredItems]);

  const columns: IColumn[] = [
    {
      key: "timestamp",
      name: "Time (CET)",
      fieldName: "timestamp",
      minWidth: 170,
      maxWidth: 220,
      isResizable: true,
      onRender: (item: ActivityLogEntity) => (
        <Text variant="small">{formatCET(item.timestamp)}</Text>
      )
    },
    {
      key: "email",
      name: "User",
      fieldName: "email",
      minWidth: 160,
      maxWidth: 260,
      isResizable: true,
      onRender: (item: ActivityLogEntity) => (
        <Text variant="small">{item.email || item.owner || "-"}</Text>
      )
    },
    {
      key: "action",
      name: "Action",
      fieldName: "action",
      minWidth: 200,
      maxWidth: 280,
      isResizable: true,
      onRender: (item: ActivityLogEntity) => (
        <Text variant="small" styles={{ root: { fontWeight: 500 } }}>
          {item.action || item.title || "Activity"}
        </Text>
      )
    },
    {
      key: "category",
      name: "Category",
      fieldName: "category",
      minWidth: 120,
      maxWidth: 160,
      isResizable: true
    },
    {
      key: "details",
      name: "Details",
      fieldName: "description",
      minWidth: 260,
      isResizable: true,
      onRender: (item: ActivityLogEntity) => {
        const parsed = parseDetails(item.metadata || item.description || "");
        if (!parsed) {
          return <Text variant="xSmall">(none)</Text>;
        }
        if (typeof parsed === "string") {
          return <Text variant="xSmall">{parsed}</Text>;
        }
        return renderParsedDetails(parsed);
      }
    }
  ];

  const userOptions: IDropdownOption[] = useMemo(() => {
    const set = new Set((items || []).map((i) => i.email || i.owner).filter((v): v is string => !!v));
    return Array.from(set).map((e) => ({ key: e as string, text: e as string }));
  }, [items]);

  const actionOptions: IDropdownOption[] = useMemo(() => {
    const set = new Set((items || []).map((i) => i.action).filter((v): v is string => !!v));
    return Array.from(set).map((e) => ({ key: e as string, text: e as string }));
  }, [items]);

  const shimmer = (
    <Shimmer
      shimmerElements={[
        { type: ShimmerElementType.line, height: 16 },
        { type: ShimmerElementType.gap, width: 8 },
        { type: ShimmerElementType.line, height: 16 },
        { type: ShimmerElementType.gap, width: 8 },
        { type: ShimmerElementType.line, height: 16 }
      ]}
    />
  );

  const chartData: IChartProps = useMemo(() => {
    const buckets = new Map<string, number>();
    for (const it of filteredItems) {
      const dt = new Date(it.timestamp);
      const key = `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, "0")}-${String(dt.getDate()).padStart(2, "0")}`;
      buckets.set(key, (buckets.get(key) || 0) + 1);
    }
    const points = Array.from(buckets.entries())
      .sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0))
      .map(([k, v], idx) => ({ x: idx + 1, y: v, xAxisCalloutData: k, yAxisCalloutData: `${v} actions` }));
    return {
      chartTitle: "Activity Over Time",
      lineChartData: [
        {
          legend: "Actions",
          data: points,
          color: theme.palette.themePrimary
        }
      ]
    } as IChartProps;
  }, [filteredItems, theme.palette.themePrimary]);

  if (loadingUser || authorized === null) {
    return (
      <div className="page-root" style={{ paddingTop: 80 }}>
        <Stack horizontalAlign="center" tokens={{ childrenGap: 16 }}>
          <Text variant="xLarge" style={{ color: theme.palette.neutralLight }}>Loading admin dashboard…</Text>
          {shimmer}
        </Stack>
      </div>
    );
  }

  if (!authorized) {
    return null;
  }

  return (
    <div className="page-root" style={{ maxWidth: 1400, margin: "0 auto" }}>
      <Stack tokens={{ childrenGap: 24 }}>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Stack tokens={{ childrenGap: 4 }}>
            <Text variant="xLarge" style={{ color: theme.palette.neutralLight, fontWeight: 600 }}>Activity Logs</Text>
            <Text variant="small" style={{ color: theme.palette.neutralTertiary }}>
              Internal audit trail for key user actions across Optical360.
            </Text>
          </Stack>
          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }}>
            <Icon
              iconName="History"
              styles={{ root: { fontSize: 20, color: theme.palette.themePrimary } }}
            />
            <Text variant="small" style={{ color: theme.palette.neutralTertiary }}>{email}</Text>
          </Stack>
        </Stack>

        {/* Metrics */}
        <Stack horizontal wrap tokens={{ childrenGap: 16 }}>
          <Stack className="metric-card" tokens={{ childrenGap: 4 }}>
            <Text variant="xSmall" className="metric-label">Total Visits Today</Text>
            <Text variant="xxLarge" className="metric-value">{totalVisitsToday}</Text>
            <Text variant="xSmall" styles={{ root: { color: theme.palette.neutralTertiary } }}>For current day (CET)</Text>
          </Stack>
          <Stack className="metric-card" tokens={{ childrenGap: 4 }}>
            <Text variant="xSmall" className="metric-label">Unique Users</Text>
            <Text variant="xxLarge" className="metric-value">{uniqueUsers}</Text>
            <Text variant="xSmall" styles={{ root: { color: theme.palette.neutralTertiary } }}>Distinct accounts in view</Text>
          </Stack>
          <Stack className="metric-card" tokens={{ childrenGap: 4 }}>
            <Text variant="xSmall" className="metric-label">Total Actions Logged</Text>
            <Text variant="xxLarge" className="metric-value">{totalActions}</Text>
            <Text variant="xSmall" styles={{ root: { color: theme.palette.neutralTertiary } }}>Filtered result set</Text>
          </Stack>
        </Stack>

        {/* Filters */}
        <Stack className="card-surface logs-filters" tokens={{ childrenGap: 12 }}>
          <Stack horizontal wrap tokens={{ childrenGap: 16 }} verticalAlign="end">
            <Dropdown
              label="User"
              placeholder="All users"
              options={userOptions}
              selectedKey={userFilter}
              onChange={(_, opt) => setUserFilter(opt ? String(opt.key) : undefined)}
              styles={{ dropdown: { minWidth: 220 } }}
            />
            <Dropdown
              label="Action"
              placeholder="All actions"
              options={actionOptions}
              selectedKey={actionFilter}
              onChange={(_, opt) => setActionFilter(opt ? String(opt.key) : undefined)}
              styles={{ dropdown: { minWidth: 260 } }}
            />
            <Dropdown
              label="Category"
              placeholder="All categories"
              options={Array.from(new Set((items || []).map(i => i.category).filter((v): v is string => !!v))).map(c => ({ key: c, text: c }))}
              selectedKey={actionFilter /* placeholder: could add dedicated categoryFilter state */}
              disabled
              styles={{ dropdown: { minWidth: 200 } }}
            />
            <DatePicker
              label="From"
              placeholder="Start date"
              value={fromDate || undefined}
              onSelectDate={(d) => setFromDate(d || null)}
            />
            <DatePicker
              label="To"
              placeholder="End date"
              value={toDate || undefined}
              onSelectDate={(d) => setToDate(d || null)}
            />
            <PrimaryButton
              text="Refresh"
              iconProps={{ iconName: "Refresh" }}
              onClick={() => {
                // trigger reload via state change
                setFromDate(fromDate ? new Date(fromDate) : null);
                setToDate(toDate ? new Date(toDate) : null);
              }}
            />
          </Stack>
          {error && (
            <Text variant="small" style={{ color: theme.palette.red }}>{error}</Text>
          )}
        </Stack>

        <Stack horizontal wrap tokens={{ childrenGap: 16 }}>
          {/* Timeline - left column */}
          <Stack grow className="card-surface" style={{ minWidth: 300, maxWidth: 420 }} tokens={{ childrenGap: 8 }}>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
              <Icon iconName="TimelineProgress" />
              <Text variant="mediumPlus">Recent Activity</Text>
            </Stack>
            {loading ? (
              <Shimmer />
            ) : (
              <Stack tokens={{ childrenGap: 8 }}>
                {filteredItems.slice(0, 12).map((it) => (
                  <Stack key={it.rowKey} horizontal tokens={{ childrenGap: 12 }} verticalAlign="center">
                    <div className="timeline-dot" />
                    <Stack>
                      <Text variant="xSmall" styles={{ root: { color: theme.palette.neutralTertiary } }}>
                        {formatCET(it.timestamp)}
                      </Text>
                      <Text variant="small" styles={{ root: { color: theme.palette.neutralLight, fontWeight: 500 } }}>
                        {it.action || it.title || "Activity"}
                      </Text>
                      <Text variant="xSmall" styles={{ root: { color: theme.palette.neutralSecondary } }}>
                        {it.email || it.owner}
                      </Text>
                      {parseDetails(it.metadata || it.description || "") && typeof parseDetails(it.metadata || it.description || "") !== "string" && (
                        <Stack className="timeline-details" tokens={{ childrenGap: 2 }}>
                          {renderParsedDetails(parseDetails(it.metadata || it.description || ""))}
                        </Stack>
                      )}
                    </Stack>
                  </Stack>
                ))}
                {!filteredItems.length && (
                  <Text variant="small" style={{ color: "#888" }}>No activity yet for the selected filters.</Text>
                )}
              </Stack>
            )}
          </Stack>

          {/* Chart - right column */}
          <Stack grow className="card-surface" tokens={{ childrenGap: 8 }}>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
              <Icon iconName="AreaChart" />
              <Text variant="mediumPlus">Activity Over Time</Text>
            </Stack>
            {loading ? (
              <Shimmer />
            ) : (
              <LineChart
                data={chartData}
                height={220}
                hideLegend={false}
                wrapXAxisLables
              />
            )}
          </Stack>
        </Stack>

        {/* Table */}
        <Stack className="card-surface" tokens={{ childrenGap: 12 }}>
          <Stack horizontal verticalAlign="center" horizontalAlign="space-between">
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
              <Icon iconName="Table" />
              <Text variant="mediumPlus">All Activity</Text>
            </Stack>
            <Text variant="xSmall" style={{ color: theme.palette.neutralTertiary }}>{filteredItems.length} entries</Text>
          </Stack>
          {loading ? (
            <Shimmer />
          ) : (
            <DetailsList
              items={filteredItems}
              columns={columns}
              layoutMode={DetailsListLayoutMode.justified}
              selectionMode={SelectionMode.none}
              setKey="logsTable"
              isHeaderVisible
            />
          )}
        </Stack>

        <Separator />
        <Text variant="xSmall" style={{ color: theme.palette.neutralTertiary }}>
          Activity logging is scoped to internal Optical360 usage only and is visible exclusively to the admin.
        </Text>
      </Stack>
    </div>
  );
};

export default Logs;
