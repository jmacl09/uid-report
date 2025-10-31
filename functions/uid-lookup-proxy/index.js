const fetch = require("node-fetch");

// Expect LOGICAPP_URL to be the base URL without the UID parameter appended.
// Example:
// https://.../api/fiberflow/triggers/When_an_HTTP_request_is_received/invoke?api-version=...&sp=...&sv=1.0&sig=...

module.exports = async function (context, req) {
  const uid = (context.bindingData && context.bindingData.uid) || (req.query && req.query.uid) || "";
  const base = process.env.LOGICAPP_URL;

  const json = (status, body, headers = {}) => {
    context.res = { status, headers: { "Content-Type": "application/json", ...headers }, body };
  };

  if (!uid || !/^\d{11}$/.test(String(uid))) {
    return json(400, { error: "Invalid UID. It must contain exactly 11 numbers." });
  }
  if (!base) {
    return json(500, { error: "Server missing LOGICAPP_URL configuration." });
  }

  try {
    const url = `${base}&UID=${encodeURIComponent(uid)}`;
    const response = await fetch(url, { method: "GET", redirect: "follow" });

    if (!response.ok) {
      const text = await response.text().catch(() => "");
      return json(response.status, { error: `Upstream error ${response.status}`, details: text || undefined });
    }

    const data = await response.json();
    return json(200, data);
  } catch (err) {
    context.log.error("Proxy fetch failed:", err);
    return json(502, { error: "Failed to fetch upstream.", details: String(err && err.message || err) });
  }
};
