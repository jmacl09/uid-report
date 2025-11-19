// Small wrapper for posting to the project's Logic App webhook used across pages
export type LogicAppResponse = {
  Spans?: any[] | any;
  DataCenter?: string;
  [k: string]: any;
};

const LOGIC_APP_URL =
  "https://fibertools-dsavavdcfdgnh2cm.westeurope-01.azurewebsites.net:443/api/VSO/triggers/When_an_HTTP_request_is_received/invoke?api-version=2022-05-01&sp=%2Ftriggers%2FWhen_an_HTTP_request_is_received%2Frun&sv=1.0&sig=6ViXNM-TmW5F7Qd9_e4fz3IhRNqmNzKwovWvcmuNJto";

async function postToLogicApp(payload: object): Promise<LogicAppResponse> {
  const res = await fetch(LOGIC_APP_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  if (!res.ok) {
    const txt = await res.text().catch(() => "");
    throw new Error(`Logic App request failed: ${res.status} ${txt}`);
  }
  const data = await res.json();
  return data as LogicAppResponse;
}

export async function getSpanUtilization(span: string) {
  if (!span) throw new Error("Span is required");
  const payload = { Stage: "11", Span: span };
  return postToLogicApp(payload);
}

export async function postStagePayload(stage: string, payload: object) {
  const body = { Stage: stage, ...payload };
  return postToLogicApp(body);
}

const fetchLogicApp = {
  getSpanUtilization,
  postStagePayload,
};

export default fetchLogicApp;
