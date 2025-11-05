const { app } = require('@azure/functions');
const { TableClient, TableServiceClient } = require('@azure/data-tables');
let DefaultAzureCredential = null;
try { ({ DefaultAzureCredential } = require('@azure/identity')); } catch { /* optional */ }

// Helper to create a TableClient using connection string or managed identity
function getTableClient(tableName) {
    const conn = process.env.TABLES_CONNECTION_STRING || process.env.AzureWebJobsStorage || '';
    const accountUrl = process.env.TABLES_ACCOUNT_URL || '';

    // 1) Managed Identity (preferred)
    if (accountUrl && DefaultAzureCredential) {
        const cred = new DefaultAzureCredential();
        const client = new TableClient(accountUrl, tableName, cred);
        const svc = TableServiceClient.fromUrl(accountUrl, cred);
        return {
            client,
            ensureTable: async () => {
                try { await svc.createTable(tableName); } 
                catch (e) { if (!(e && e.statusCode === 409)) throw e; }
            }
        };
    }

    // 2) Connection string fallback
    if (!conn) throw new Error('Missing Azure Table configuration.');
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
        context.log(`HttpTrigger1 received ${request.method} for ${request.url}`);

        // --- Proper CORS setup ---
        const corsHeaders = {
            'Access-Control-Allow-Origin': 'https://optical360.net',
            'Access-Control-Allow-Methods': 'GET,POST,OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type,Authorization',
            'Access-Control-Allow-Credentials': 'true'
        };

        // Handle preflight
        if (request.method === 'OPTIONS') {
            return { status: 204, headers: corsHeaders };
        }

        // Simple health check
        if (request.method === 'GET') {
            return {
                status: 200,
                headers: corsHeaders,
                jsonBody: { ok: true, message: 'Storage endpoint active. Use POST with JSON body.' }
            };
        }

        // --- POST handling ---
        let payload;
        try {
            payload = await request.json();
            context.log('Received payload:', payload);
        } catch (e) {
            const txt = await request.text().catch(() => '');
            context.log('Invalid JSON body:', txt);
            return { 
                status: 400, 
                headers: corsHeaders, 
                jsonBody: { ok: false, error: 'Invalid JSON body', received: txt?.slice(0, 200) } 
            };
        }

        // Destructure fields with sane defaults
        const { category, uid, title, description = '', owner = 'Unknown', timestamp } = payload || {};

        // Validate required
        if (!uid || !category || !title) {
            return { 
                status: 400, 
                headers: corsHeaders, 
                jsonBody: { ok: false, error: 'Missing required fields: uid, category, title', received: payload } 
            };
        }

        try {
            const tableName = process.env.TABLES_TABLE_NAME || 'Projects';
            const { client, ensureTable } = getTableClient(tableName);
            await ensureTable();

            const nowIso = (() => {
                try { return timestamp ? new Date(timestamp).toISOString() : new Date().toISOString(); }
                catch { return new Date().toISOString(); }
            })();

            const entity = {
                partitionKey: `UID_${uid}`,
                rowKey: nowIso, // time-ordered for multiple entries
                category,
                title,
                description,
                owner,
                savedAt: nowIso
            };

            await client.upsertEntity(entity, 'Merge');

            context.log(`Saved ${category} for UID ${uid} at ${nowIso}`);

            return {
                status: 200,
                headers: corsHeaders,
                jsonBody: { ok: true, message: `Saved ${category} for UID ${uid}`, entity }
            };

        } catch (err) {
            console.error('Save failed:', err);
            return { 
                status: 500, 
                headers: corsHeaders, 
                jsonBody: { ok: false, error: String(err?.message || err) } 
            };
        }
    }
});
