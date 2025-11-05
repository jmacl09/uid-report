const { app } = require('@azure/functions');
const { TableClient, TableServiceClient } = require('@azure/data-tables');
let DefaultAzureCredential = null;
try { ({ DefaultAzureCredential } = require('@azure/identity')); } catch { /* optional */ }

// Helper: returns a client and ensures table exists
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
                try { await svc.createTable(tableName); } catch (e) { if (!(e && e.statusCode === 409)) throw e; }
            }
        };
    }

    if (!conn) throw new Error('Missing Azure Table configuration. Provide TABLES_CONNECTION_STRING or TABLES_ACCOUNT_URL.');
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
        context.log(`HTTP function processed request for URL "${request.url}"`);

        // --- CORS setup ---
        const corsHeaders = {
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET,POST,OPTIONS',
            'Access-Control-Allow-Headers': '*',
        };

        // --- OPTIONS preflight ---
        if (request.method === 'OPTIONS') {
            return { status: 204, headers: corsHeaders };
        }

        // --- GET health check ---
        if (request.method === 'GET') {
            return {
                status: 200,
                headers: corsHeaders,
                jsonBody: { ok: true, message: 'Storage endpoint is alive. Use POST with JSON body to save.' },
            };
        }

        // --- Parse POST body ---
        let payload;
        try {
            payload = await request.json();
        } catch {
            const txt = await request.text().catch(() => '');
            return { status: 400, headers: corsHeaders, jsonBody: { ok: false, error: 'Invalid JSON body', received: txt?.slice(0, 200) } };
        }

        const { category, uid, title, description, owner, timestamp } = payload || {};
        if (!uid || !category || !title) {
            return { status: 400, headers: corsHeaders, jsonBody: { ok: false, error: 'Missing required fields: uid, category, title' } };
        }

        try {
            // --- Determine correct table ---
            let tableName;
            if (category === 'Comments' || category === 'Notes') {
                tableName = 'Notes';
            } else if (category === 'Projects') {
                tableName = 'Projects';
            } else {
                tableName = process.env.TABLES_TABLE_NAME || 'Projects';
            }

            const { client, ensureTable } = getTableClient(tableName);
            await ensureTable();

            const nowIso = (() => {
                try { return timestamp ? (new Date(timestamp)).toISOString() : new Date().toISOString(); } catch { return new Date().toISOString(); }
            })();

            // --- Build entities differently per table ---
            let entity;
            if (tableName === 'Notes') {
                entity = {
                    partitionKey: `UID_${uid}`,
                    rowKey: nowIso,
                    UID: uid,
                    Category: category,
                    Comment: description || '',
                    User: owner || 'Unknown',
                    Title: title,
                    CreatedAt: nowIso,
                };
            } else if (tableName === 'Projects') {
                entity = {
                    partitionKey: `UID_${uid}`,
                    rowKey: nowIso,
                    UID: uid,
                    Category: category,
                    Title: title,
                    Description: description || '',
                    Owner: owner || 'Unknown',
                    CreatedAt: nowIso,
                    Status: 'In Progress',
                    LastUpdated: nowIso,
                };
            }

            await client.upsertEntity(entity, 'Merge');

            return {
                status: 200,
                headers: corsHeaders,
                jsonBody: {
                    ok: true,
                    message: `Saved ${category} for UID ${uid} → ${tableName} table`,
                    entity,
                },
            };
        } catch (err) {
            console.error('Save failed:', err);
            return { status: 500, headers: corsHeaders, jsonBody: { ok: false, error: String(err?.message || err) } };
        }
    },
});
