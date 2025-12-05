import React, { useEffect, useState } from "react";
import { Stack, Text, Shimmer } from "@fluentui/react";
import "../Theme.css";

interface VisitorStat {
  email: string;
  daysVisited: number;
}

const TopVisitors: React.FC = () => {
  const [loading, setLoading] = useState(true);
  const [stats, setStats] = useState<VisitorStat[]>([]);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    const load = async () => {
      setLoading(true);
      setError(null);
      try {
        const res = await fetch(`/api/log?limit=500`);
        const data = await res.json();
        const items = data.items || [];

        // Map email -> set of dates (yyyy-mm-dd)
        const map = new Map<string, Set<string>>();

        for (const it of items) {
          const email = (it.email || it.owner || "").toLowerCase();
          if (!email) continue;
          const ts = it.timestamp || it.Timestamp || it.savedAt;
          if (!ts) continue;
          const d = new Date(ts);
          const key = d.toISOString().substring(0, 10);
          if (!map.has(email)) map.set(email, new Set());
          map.get(email)!.add(key);
        }

        const arr: VisitorStat[] = Array.from(map.entries()).map(([email, days]) => ({
          email,
          daysVisited: days.size
        })).sort((a, b) => b.daysVisited - a.daysVisited);

        if (!cancelled) {
          setStats(arr);
        }
      } catch (err: any) {
        if (!cancelled) setError(err.message || "Failed to load visitors");
      } finally {
        if (!cancelled) setLoading(false);
      }
    };

    load();
    return () => { cancelled = true; };
  }, []);

  return (
    <div className="page-root">
      <Stack tokens={{ childrenGap: 18 }}>
        <Stack horizontal horizontalAlign="space-between">
          <Stack>
            <Text variant="xxLarge" style={{ fontWeight: 700, color: "#ffffff" }}>
              Top Visitors
            </Text>
            <Text variant="small" style={{ color: "#9cb3d8" }}>
              Visitors by distinct days visited and their email addresses.
            </Text>
          </Stack>
        </Stack>

        <div className="card-surface">
          {loading ? (
            <Shimmer />
          ) : error ? (
            <Text style={{ color: "#e06666" }}>{error}</Text>
          ) : (
            <Stack tokens={{ childrenGap: 12 }}>
              {stats.length === 0 && <Text style={{ color: "#9cb3d8" }}>No visitors found.</Text>}

              <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
                {stats.slice(0, 50).map((s) => (
                  <div key={s.email} className="metric-card">
                    <Text className="metric-label">{s.email}</Text>
                    <Text variant="xLarge" className="metric-value">{s.daysVisited}</Text>
                    <Text variant="small" style={{ color: "#99b2d6" }}>distinct days visited</Text>
                  </div>
                ))}
              </Stack>
            </Stack>
          )}
        </div>
      </Stack>
    </div>
  );
};

export default TopVisitors;
