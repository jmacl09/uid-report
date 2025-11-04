const { app } = require('@azure/functions');
const { TableClient, TableServiceClient } = require('@azure/data-tables');
let DefaultAzureCredential = null;
try { ({ DefaultAzureCredential } = require('@azure/identity')); } catch { /* optional */ }

// Helper to create a TableClient using the local Azurite or Azure connection string
function getTableClient(tableName) {
    const conn = process.env.TABLES_CONNECTION_STRING || process.env.AzureWebJobsStorage || '';
    const accountUrl = process.env.TABLES_ACCOUNT_URL || '';
    // 1) If account URL + DefaultAzureCredential available, use it
    if (accountUrl && DefaultAzureCredential) {
        const cred = new DefaultAzureCredential();
        const client = new TableClient(accountUrl, tableName, cred);
        const svc = TableServiceClient.fromUrl(accountUrl, cred);
        return {
            client,
            ensureTable: async () => {
                try { await svc.createTable(tableName); } catch (e) { if (!(e && e.statusCode === 409)) throw e; }
            }
        };
    }
    // 2) Else use connection string
    if (!conn) throw new Error('Missing Azure Table configuration. Provide TABLES_CONNECTION_STRING or TABLES_ACCOUNT_URL, or AzureWebJobsStorage for local Azurite.');
    return {
        client: TableClient.fromConnectionString(conn, tableName),
        ensureTable: async () => {
            try {
                const svc = TableServiceClient.fromConnectionString(conn);
                await svc.createTable(tableName);
            } catch (e) {
                if (!(e && e.statusCode === 409)) throw e;
            }
        }
    };
}

app.http('HttpTrigger1', {
    methods: ['GET', 'POST', 'OPTIONS'],
    authLevel: 'anonymous',
    handler: async (request, context) => {
        context.log(`Http function processed request for url "${request.url}"`);

        // Basic CORS handling
        const corsHeaders = {
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET,POST,OPTIONS',
            'Access-Control-Allow-Headers': '*',
        };

        if (request.method === 'OPTIONS') {
            // Preflight
            return { status: 204, headers: corsHeaders };
        }

        if (request.method === 'GET') {
            return {
                status: 200,
                headers: corsHeaders,
                jsonBody: { ok: true, message: 'UID storage endpoint is alive. Use POST with JSON to save.' },
            };
        }

        // POST: expect JSON body with { category, uid, title, description, owner, timestamp? }
        let payload;
        try {
            payload = await request.json();
        } catch (e) {
            const txt = await request.text().catch(() => '');
                    return { status: 400, headers: corsHeaders, jsonBody: { ok: false, error: 'Invalid JSON body', received: txt?.slice(0, 200) } };
        }

        const { category, uid, title, description, owner, timestamp } = payload || {};
            if (!uid || !category || !title) {
                return { status: 400, headers: corsHeaders, jsonBody: { ok: false, error: 'Missing required fields: uid, category, title' } };
        }

        try {
            const tableName = process.env.TABLES_TABLE_NAME || 'Projects';
            const { client, ensureTable } = getTableClient(tableName);
            await ensureTable();

            const nowIso = (() => {
                try { return timestamp ? (new Date(timestamp)).toISOString() : new Date().toISOString(); } catch { return new Date().toISOString(); }
            })();
            const entity = {
                partitionKey: `UID_${uid}`,
                rowKey: nowIso, // time-ordered; allows multiple comments
                category,
                title,
                description: description || '',
                owner: owner || 'Unknown',
                savedAt: nowIso,
            };

            await client.upsertEntity(entity, 'Merge');

                    return {
                        status: 200,
                        headers: corsHeaders,
                        jsonBody: { ok: true, message: `Saved ${category} for UID ${uid}`, entity },
                    };
            } catch (err) {
                // Use console.error because context.log may not expose .error in some hosts
                console.error('Save failed:', err);
                return { status: 500, headers: corsHeaders, jsonBody: { ok: false, error: String(err?.message || err) } };
        }
    },
});
