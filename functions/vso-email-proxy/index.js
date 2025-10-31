const fetch = require('node-fetch');

module.exports = async function (context, req) {
  try {
    const upstream = process.env.VSO_LOGICAPP_URL; // e.g. https://.../api/VSO/...&sig=...
    if (!upstream) {
      context.log.error('VSO_LOGICAPP_URL is not configured');
      context.res = {
        status: 500,
        body: { error: 'Server configuration missing (VSO_LOGICAPP_URL)' }
      };
      return;
    }

    // Pass-through payload. Expecting JSON with Stage: "1" or "2" and other fields
    const payload = (req && req.body) || {};

    const resp = await fetch(upstream, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });

    const contentType = resp.headers.get('content-type') || '';
    const isJson = /application\/json/i.test(contentType);

    if (!resp.ok) {
      const text = await resp.text().catch(() => '');
      context.res = {
        status: resp.status,
        body: isJson ? { error: text } : text
      };
      return;
    }

    const body = isJson ? await resp.json() : await resp.text();

    context.res = {
      status: 200,
      headers: { 'Content-Type': isJson ? 'application/json' : 'text/plain' },
      body
    };
  } catch (err) {
    context.log.error('vso-email-proxy error', err);
    context.res = {
      status: 500,
      body: { error: 'Unexpected error', details: (err && err.message) || String(err) }
    };
  }
};
