// ================================================================
// Azure Function: Projects
// Handles POST requests from Optical360 frontend to save Notes,
// Comments, or Projects into Azure Table Storage
// ================================================================
const { app } = require('@azure/functions');
const { TableClient, TableServiceClient } = require('@azure/data-tables');
let DefaultAzureCredential = null;
try { ({ DefaultAzureCredential } = require('@azure/identity')); } catch { /* optional */ }

// ---------------------------------------------------------------
// Helper: Return a TableClient instance (Managed Identity or ConnStr)
// ---------------------------------------------------------------
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
        try { await svc.createTable(tableName); } 
        catch (e) { if (e.statusCode !== 409) throw e; } // ignore "already exists"
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

// ---------------------------------------------------------------
// HTTP Trigger: /api/projects
// ---------------------------------------------------------------
app.http('Projects', {
  methods: ['GET', 'POST', 'OPTIONS'],
  authLevel: 'anonymous',
  handler: async (request, context) => {

    // --- CORS headers ---
    const corsHeaders = {
      'Access-Control-Allow-Origin': 'https://optical360.net',
      'Access-Control-Allow-Methods': 'GET,POST,OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type,Authorization',
      'Access-Control-Allow-Credentials': 'true'
    };

    // --- Handle preflight (OPTIONS) ---
    if (request.method === 'OPTIONS')
      return { status: 204, headers: corsHeaders };

    // --- Health Check (GET) ---
    if (request.method === 'GET')
      return { status: 200, headers: corsHeaders, jsonBody: { ok: true, message: '✅ Projects API ready.' } };

    // -----------------------------------------------------------
    // ✅ Robust JSON parsing (for Node 18+ / Azure Functions v4)
    // -----------------------------------------------------------
    let payload = {};
    try {
      payload = await request.json();
    } catch (err) {
      context.log('❌ Invalid JSON body:', err);
      return { 
        status: 400, 
        headers: corsHeaders, 
        jsonBody: { ok: false, error: 'Invalid JSON', details: String(err) } 
      };
    }

    context.log('DEBUG parsed payload:', JSON.stringify(payload));

    const { category, uid, title, description = '', owner = 'Unknown', timestamp } = payload;

    // -----------------------------------------------------------
    // 🔸 Validate required fields
    // -----------------------------------------------------------
    if (!uid || !category || !title) {
      context.log('⚠️ Missing required fields:', { uid, category, title });
      return { 
        status: 400, 
        headers: corsHeaders, 
        jsonBody: { ok: false, error: 'Missing required fields', payload } 
      };
    }

    // -----------------------------------------------------------
    // 💾 Insert entity into Azure Table
    // -----------------------------------------------------------
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
      context.log(`✅ Saved entity for UID ${uid} to table ${tableName}`);

      return { status: 200, headers: corsHeaders, jsonBody: { ok: true, entity } };

    } catch (err) {
      context.log('❌ Table write error:', err);
      return { 
        status: 500, 
        headers: corsHeaders, 
        jsonBody: { ok: false, error: String(err?.message || err) } 
      };
    }
  }
});
