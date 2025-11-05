const { app } = require('@azure/functions');
const { TableClient, TableServiceClient } = require('@azure/data-tables');
let DefaultAzureCredential = null;
try { ({ DefaultAzureCredential } = require('@azure/identity')); } catch { /* optional */ }

function getTableClient(tableName) {
  const conn = process.env.TABLES_CONNECTION_STRING || process.env.AzureWebJobsStorage || '';
  const accountUrl = process.env.TABLES_ACCOUNT_URL || '';

  if (accountUrl && DefaultAzureCredential) {
    const cred = new DefaultAzureCredential();
    return new TableClient(accountUrl, tableName, cred);
  }
  return TableClient.fromConnectionString(conn, tableName);
}

app.http('DeleteItem', {
  methods: ['POST', 'OPTIONS'],
  authLevel: 'anonymous',
  handler: async (req, context) => {
    const corsHeaders = {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'POST,OPTIONS',
      'Access-Control-Allow-Headers': '*',
    };

    // Preflight (OPTIONS)
    if (req.method === 'OPTIONS') {
      return { status: 204, headers: corsHeaders };
    }

    let body;
    try {
      body = await req.json();
    } catch {
      return { status: 400, headers: corsHeaders, jsonBody: { ok: false, error: 'Invalid JSON body' } };
    }

    const { category, partitionKey, rowKey } = body;
    if (!partitionKey || !rowKey || !category) {
      return { status: 400, headers: corsHeaders, jsonBody: { ok: false, error: 'Missing partitionKey, rowKey, or category' } };
    }

    try {
      // Determine correct table
      const tableName =
        category === 'Comments' || category === 'Notes' ? 'Notes' :
        category === 'Projects' ? 'Projects' :
        process.env.TABLES_TABLE_NAME || 'Projects';

      const client = getTableClient(tableName);

      await client.deleteEntity(partitionKey, rowKey);

      return {
        status: 200,
        headers: corsHeaders,
        jsonBody: { ok: true, message: `Deleted ${category} from ${tableName}` },
      };
    } catch (err) {
      console.error('Delete failed:', err);
      return { status: 500, headers: corsHeaders, jsonBody: { ok: false, error: String(err?.message || err) } };
    }
  },
});
