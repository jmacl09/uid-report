import React, { useEffect, useMemo, useState } from "react";
import { logAction } from "../api/log";

type Stage = "VSO_Details" | "Email_Template" | "Card_3";

type VsoDetailsRequest = {
  FacilityCodeA: string;
  Diversity?: string;
  SpliceRackA?: string;
  Stage: Stage;
};

type SpanRow = {
  SpanID: string;
  Diversity: string;
  IDF_A: string;
  SpliceRackA: string;
  WiringScope: string;
  Status: string;
  Color: string;
  OpticalLink: string;
  FormattedSpans: string;
  prodSpans?: number;
  maintSpans?: number;
  prodPct?: number;
};

type VsoDetailsResponse = {
  Spans: SpanRow[];
  DataCenter: string;
};

export default function VSO2() {
  const [facility, setFacility] = useState("");
  const [diversity, setDiversity] = useState("N");
  const [spliceRack, setSpliceRack] = useState("N");
  const [stage, setStage] = useState<Stage>("VSO_Details");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [data, setData] = useState<VsoDetailsResponse | null>(null);

  useEffect(() => {
    const email = (() => {
      try {
        return localStorage.getItem("loggedInEmail") || "";
      } catch {
        return "";
      }
    })();
    logAction(email, "View VSO 2");
  }, []);

  const canSubmit = useMemo(
    () => !!stage && (stage !== "VSO_Details" || !!facility),
    [stage, facility]
  );

  const submit = async () => {
    const email = (() => {
      try {
        return localStorage.getItem("loggedInEmail") || "";
      } catch {
        return "";
      }
    })();
    logAction(email, "Submit VSO2 Request", {
      facility,
      diversity,
      spliceRack,
      stage,
    });

    setLoading(true);
    setError(null);
    setData(null);

    try {
      const payload: VsoDetailsRequest = {
        FacilityCodeA: facility,
        Diversity: diversity,
        SpliceRackA: spliceRack,
        Stage: stage,
      };

      const res = await fetch("/api/LogicAppProxy", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ type: "VSO", ...payload }),
      });

      const raw = await res.text();
      let json: any;

      try {
        json = JSON.parse(raw);
      } catch (e) {
        throw new Error(
          `Backend did not return JSON (HTTP ${res.status}). Body: ${raw.slice(
            0,
            240
          )}`
        );
      }

      if (!res.ok) {
        throw new Error(
          json?.error ||
            json?.message ||
            `Request failed (${res.status})`
        );
      }

      setData(json as VsoDetailsResponse);
    } catch (e: any) {
      setError(e?.message || "Unknown error");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{ padding: 16 }}>
      <h2>VSO 2</h2>
      <div style={{ display: "grid", gap: 12, maxWidth: 720 }}>
        <label>
          Stage
          <select
            value={stage}
            onChange={(e) => setStage(e.target.value as Stage)}
            style={{ marginLeft: 8 }}
          >
            <option value="VSO_Details">VSO_Details</option>
            <option value="Email_Template">Email_Template</option>
          </select>
        </label>

        {stage === "VSO_Details" && (
          <>
            <label>
              FacilityCodeA
              <input
                value={facility}
                onChange={(e) => setFacility(e.target.value)}
                placeholder="e.g., DJW1"
                style={{ marginLeft: 8 }}
              />
            </label>

            <label>
              Diversity
              <input
                value={diversity}
                onChange={(e) => setDiversity(e.target.value)}
                placeholder="N for no filter"
                style={{ marginLeft: 8 }}
              />
            </label>

            <label>
              SpliceRackA
              <input
                value={spliceRack}
                onChange={(e) => setSpliceRack(e.target.value)}
                placeholder="N for no filter"
                style={{ marginLeft: 8 }}
              />
            </label>
          </>
        )}

        <button disabled={!canSubmit || loading} onClick={submit}>
          {loading
            ? "Running..."
            : stage === "VSO_Details"
            ? "Run Query"
            : "Send Email"}
        </button>

        {error && (
          <div style={{ color: "#b10" }}>
            <strong>Error:</strong> {error}
          </div>
        )}

        {data && (
          <div>
            <h3>Data Center: {data.DataCenter}</h3>

            <div style={{ overflowX: "auto" }}>
              <table
                style={{
                  borderCollapse: "collapse",
                  minWidth: 960,
                }}
              >
                <thead>
                  <tr>
                    {[
                      "Diversity",
                      "SpanID",
                      "IDF_A",
                      "SpliceRackA",
                      "WiringScope",
                      "Status",
                      "prodSpans",
                      "maintSpans",
                      "prodPct",
                    ].map((h) => (
                      <th
                        key={h}
                        style={{
                          textAlign: "left",
                          padding: "6px 10px",
                          borderBottom: "1px solid #ddd",
                        }}
                      >
                        {h}
                      </th>
                    ))}
                  </tr>
                </thead>

                <tbody>
                  {data.Spans.map((r: SpanRow) => (
                    <tr key={r.SpanID}>
                      <td
                        style={{
                          padding: "6px 10px",
                          borderBottom: "1px solid #eee",
                        }}
                      >
                        {r.Diversity}
                      </td>
                      <td
                        style={{
                          padding: "6px 10px",
                          borderBottom: "1px solid #eee",
                        }}
                      >
                        <a
                          href={r.OpticalLink}
                          target="_blank"
                          rel="noreferrer"
                        >
                          {r.SpanID}
                        </a>
                      </td>
                      <td
                        style={{
                          padding: "6px 10px",
                          borderBottom: "1px solid #eee",
                        }}
                      >
                        {r.IDF_A}
                      </td>
                      <td
                        style={{
                          padding: "6px 10px",
                          borderBottom: "1px solid #eee",
                        }}
                      >
                        {r.SpliceRackA}
                      </td>
                      <td
                        style={{
                          padding: "6px 10px",
                          borderBottom: "1px solid #eee",
                        }}
                      >
                        {r.WiringScope}
                      </td>
                      <td
                        style={{
                          padding: "6px 10px",
                          borderBottom: "1px solid #eee",
                        }}
                      >
                        {r.Status}
                      </td>
                      <td
                        style={{
                          padding: "6px 10px",
                          borderBottom: "1px solid #eee",
                        }}
                      >
                        {r.prodSpans ?? ""}
                      </td>
                      <td
                        style={{
                          padding: "6px 10px",
                          borderBottom: "1px solid #eee",
                        }}
                      >
                        {r.maintSpans ?? ""}
                      </td>
                      <td
                        style={{
                          padding: "6px 10px",
                          borderBottom: "1px solid #eee",
                        }}
                      >
                        {r.prodPct ?? ""}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
