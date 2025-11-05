const { app } = require('@azure/functions');
const { TableClient, TableServiceClient } = require('@azure/data-tables');
let DefaultAzureCredential = null;
try { ({ DefaultAzureCredential } = require('@azure/identity')); } catch { /* optional */ }

//-------------------------------------------------------------
// Helper to get a TableClient instance
//-------------------------------------------------------------
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

//-------------------------------------------------------------
// HTTP Trigger
//-------------------------------------------------------------
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

    if (request.method === 'OPTIONS')
      return { status: 204, headers: corsHeaders };

    if (request.method === 'GET')
      return { status: 200, headers: corsHeaders, jsonBody: { ok: true, message: 'Ready.' } };

    //-----------------------------------------------------------
    // Parse request body
    //-----------------------------------------------------------
    let bodyText = await request.text();
    let payload = {};
    try {
      payload = JSON.parse(bodyText);
    } catch {
      context.log('DEBUG Invalid JSON body:', bodyText);
      return { status: 400, headers: corsHeaders, jsonBody: { ok: false, error: 'Invalid JSON', bodyText } };
    }

    context.log('DEBUG raw bodyText:', bodyText);
    context.log('DEBUG parsed payload:', JSON.stringify(payload));

    const { category, uid, title, description = '', owner = 'Unknown', timestamp } = payload;

    //-----------------------------------------------------------
    // Validate required fields
    //-----------------------------------------------------------
    if (!uid || !category || !title) {
      context.log('DEBUG missing fields:', { uid, category, title });
      return { status: 400, headers: corsHeaders, jsonBody: { ok: false, error: 'Missing required fields', payload } };
    }

    //-----------------------------------------------------------
    // Log environment variables so we can verify configuration
    //-----------------------------------------------------------
    context.log('DEBUG TABLES_ACCOUNT_URL:', process.env.TABLES_ACCOUNT_URL);
    context.log('DEBUG TABLE_NAME:', process.env.TABLE_NAME);
    context.log('DEBUG AzureWebJobsStorage present?:', !!process.env.AzureWebJobsStorage);

    //-----------------------------------------------------------
    // Insert entity into table
    //-----------------------------------------------------------
    try {
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
