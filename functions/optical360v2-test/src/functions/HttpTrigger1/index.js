const { app } = require('@azure/functions');
const { TableClient, TableServiceClient } = require('@azure/data-tables');
let DefaultAzureCredential = null;
try { ({ DefaultAzureCredential } = require('@azure/identity')); } catch { /* optional */ }

function getTableClient(tableName) {
  const conn = process.env.TABLES_CONNECTION_STRING || process.env.AzureWebJobsStorage || '';
  const accountUrl = process.env.TABLES_ACCOUNT_URL || '';
  if (accountUrl && DefaultAzureCredential) {
    const cred = new DefaultAzureCredential();
    const client = new TableClient(accountUrl, tableName, cred);
    const svc = TableServiceClient.fromUrl(accountUrl, cred);
    return {
      client,
      ensureTable: async () => {
        try { await svc.createTable(tableName); } catch (e) { if (e.statusCode !== 409) throw e; }
      }
    };
  }
  if (!conn) throw new Error('Missing Azure Table configuration.');
  return {
    client: TableClient.fromConnectionString(conn, tableName),
    ensureTable: async () => {
      try {
        const svc = TableServiceClient.fromConnectionString(conn);
        await svc.createTable(tableName);
      } catch (e) {
        if (e.statusCode !== 409) throw e;
      }
    }
  };
}

app.http('HttpTrigger1', {
  methods: ['GET', 'POST', 'OPTIONS'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    const corsHeaders = {
      'Access-Control-Allow-Origin': 'https://optical360.net',
      'Access-Control-Allow-Methods': 'GET,POST,OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type,Authorization',
      'Access-Control-Allow-Credentials': 'true'
    };

    if (request.method === 'OPTIONS') return { status: 204, headers: corsHeaders };
    if (request.method === 'GET')
      return { status: 200, headers: corsHeaders, jsonBody: { ok: true, message: 'Ready.' } };

    let bodyText = await request.text();
    let payload = {};
    try {
      payload = JSON.parse(bodyText);
    } catch {
      return { status: 400, headers: corsHeaders, jsonBody: { ok: false, error: 'Invalid JSON', bodyText } };
    }

    const { category, uid, title, description = '', owner = 'Unknown', timestamp } = payload;
    if (!uid || !category || !title) {
      return { status: 400, headers: corsHeaders, jsonBody: { ok: false, error: 'Missing required fields', payload } };
    }

    try {
      // ✅ Corrected environment variable name here:
      const tableName = process.env.TABLE_NAME || 'Projects';
      const { client, ensureTable } = getTableClient(tableName);
      await ensureTable();

      const nowIso = timestamp ? new Date(timestamp).toISOString() : new Date().toISOString();
      const entity = {
        partitionKey: `UID_${uid}`,
        rowKey: nowIso,
        category,
        title,
        description,
        owner,
        savedAt: nowIso
      };

      await client.upsertEntity(entity, 'Merge');
      context.log(`Saved entity for UID ${uid} to table ${tableName}`);

      return { status: 200, headers: corsHeaders, jsonBody: { ok: true, entity } };
    } catch (err) {
      context.log('Error:', err);
      return { status: 500, headers: corsHeaders, jsonBody: { ok: false, error: String(err?.message || err) } };
    }
  }
});
