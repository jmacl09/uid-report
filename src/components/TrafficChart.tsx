import React from "react";
import {
  ResponsiveContainer,
  LineChart,
  Line,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
} from "recharts";

type RawPoint = {
  TIMESTAMP: string | number;
  in_gbps: number;
  SolutionId?: string;
  [k: string]: any;
};

type Props = {
  data: RawPoint[];
  height?: number;
  colorMap?: Record<string, string>;
};

// Small color palette for up to ~8 lines; it will cycle if more are present.
const COLORS = ["#60A5FA", "#34D399", "#F59E0B", "#F472B6", "#A78BFA", "#FB7185", "#60C8E8", "#FCD34D"];

function formatTimestampLabel(val: string | number) {
  try {
    const d = new Date(val);
    if (isNaN(d.getTime())) return String(val);
    return d.toLocaleString();
  } catch {
    return String(val);
  }
}

const TrafficChart: React.FC<Props> = ({ data, height = 340, colorMap }) => {
  if (!data || data.length === 0) return null;

  // Normalize: ensure each point has a timestamp string and numeric value
  const points: RawPoint[] = data.map((p) => ({ ...p, TIMESTAMP: p.TIMESTAMP, in_gbps: Number(p.in_gbps || 0) }));

  // Determine unique solution ids and timestamps and convert timestamps to numeric ms
  const solutionSet = new Set<string>();
  const tsSet = new Set<number>();
  for (const p of points) {
    const sid = (p.SolutionId || "_default").toString();
    solutionSet.add(sid);
    const dt = new Date(p.TIMESTAMP as any);
    const ms = isNaN(dt.getTime()) ? Date.now() : dt.getTime();
    tsSet.add(ms);
  }

  const solutionIds = Array.from(solutionSet);
  const timestamps = Array.from(tsSet).sort((a, b) => a - b);

  // Build series: one aggregated object per timestamp (use numeric __ts field)
  const seriesMap: Record<number, any> = {};
  for (const ts of timestamps) {
    seriesMap[ts] = { __ts: ts };
  }

  for (const p of points) {
    const dt = new Date(p.TIMESTAMP as any);
    const ts = isNaN(dt.getTime()) ? Date.now() : dt.getTime();
    const sid = (p.SolutionId || "_default").toString();
    const obj = seriesMap[ts] || (seriesMap[ts] = { __ts: ts });
    // If multiple points for same sid/timestamp exist, accumulate
    obj[sid] = (obj[sid] || 0) + Number(p.in_gbps || 0);
  }

  const chartData = timestamps.map((t) => seriesMap[t]);

  const tooltipFormatter = (value: any, name: string) => {
    return [Number(value).toLocaleString(), `${name} Gbps`];
  };

  const TooltipContent: React.FC<any> = ({ active, payload, label }) => {
    if (!active || !payload || !payload.length) return null;
    return (
      <div
        style={{
          background: "#071821",
          color: "#fff",
          padding: 10,
          borderRadius: 8,
          boxShadow: "0 6px 18px rgba(0,0,0,0.6)",
          fontSize: 12,
          minWidth: 160,
        }}
      >
        <div style={{ fontWeight: 700, marginBottom: 6 }}>{formatTimestampLabel(label)}</div>
        {payload.map((pl: any, idx: number) => (
          <div
            key={idx}
            style={{ display: "flex", justifyContent: "space-between", gap: 8, alignItems: "center", marginTop: 6 }}
          >
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <span
                style={{
                  width: 10,
                  height: 10,
                  background: pl.color || "#888",
                  borderRadius: 6,
                  display: "inline-block",
                  boxShadow: "0 0 0 1px rgba(0,0,0,0.4) inset",
                }}
              />
              <div style={{ color: "#e6f1ff" }}>{pl.name}</div>
            </div>
            <div style={{ fontWeight: 700, color: "#fff" }}>{Number(pl.value).toLocaleString()} Gbps</div>
          </div>
        ))}
      </div>
    );
  };

  return (
    <div className="traffic-chart-card" style={{ width: "100%", height, minHeight: 120 }}>
      <ResponsiveContainer width="100%" height={height}>
        <LineChart data={chartData} margin={{ top: 12, right: 24, left: 8, bottom: 8 }}>
          <CartesianGrid strokeDasharray="3 3" stroke="#1f2937" />
          <XAxis
            dataKey="__ts"
            type="number"
            domain={["dataMin", "dataMax"]}
            tickFormatter={(v: any) => {
              try {
                const d = new Date(Number(v));
                return d.toLocaleString(undefined, { hour: "2-digit", minute: "2-digit", month: "short", day: "numeric" });
              } catch {
                return String(v);
              }
            }}
            stroke="#9ca3af"
            minTickGap={20}
          />
          <YAxis stroke="#9ca3af" />
          <Tooltip content={<TooltipContent />} formatter={tooltipFormatter} labelFormatter={(v:any)=>formatTimestampLabel(v)} />
          <Legend wrapperStyle={{ color: '#cbd5e1' }} />
          {solutionIds.map((sid, idx) => (
            <Line
              key={sid}
              type="monotone"
              dataKey={sid}
              name={sid}
              stroke={colorMap && colorMap[sid] ? colorMap[sid] : COLORS[idx % COLORS.length]}
              strokeWidth={2}
              dot={false}
              activeDot={{ r: 6 }}
              connectNulls
            />
          ))}
        </LineChart>
      </ResponsiveContainer>
    </div>
  );
};

export default TrafficChart;
