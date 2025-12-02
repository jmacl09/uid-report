// Wrapper for posting to UID + VSO Logic Apps through LogicAppProxy

export type LogicAppResponse = {
  Spans?: any[] | any;
  DataCenter?: string;
  [k: string]: any;
};

/**
 * Generic POST helper that always calls the Azure Function API:
 * /api/LogicAppProxy
 */
async function callProxy(payload: object): Promise<LogicAppResponse> {
  const res = await fetch("/api/LogicAppProxy", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });

  if (!res.ok) {
    const txt = await res.text().catch(() => "");
    throw new Error(`LogicAppProxy failed: ${res.status} — ${txt}`);
  }

  return (await res.json()) as LogicAppResponse;
}

/**
 * -------------------------
 * VSO Logic App (Stage-based)
 * -------------------------
 */
export async function vsoPostStage(stage: string | number, payload: object) {
  return callProxy({
    type: "VSO",
    Stage: stage,
    ...payload,
  });
}

/**
 * Fiber Span Utilization — Stage 11
 * Called exactly like the VSO Assistant
 */
export async function getSpanUtilization(span: string, days?: number) {
  if (!span) throw new Error("Span is required");

  const payload: any = {
    type: "VSO",
    Stage: "11",
    Span: span,
  };

  if (typeof days === "number") {
    payload.Days = days;
  }

  return callProxy(payload);
}

/**
 * -------------------------
 * UID Logic App functions
 * -------------------------
 */
export async function uidPostStage(stage: string | number, payload: object) {
  return callProxy({
    type: "UID",
    Stage: stage,
    ...payload,
  });
}

/**
 * Export grouped API
 */
const fetchLogicApp = {
  vsoPostStage,
  uidPostStage,
  getSpanUtilization,
};

export default fetchLogicApp;
