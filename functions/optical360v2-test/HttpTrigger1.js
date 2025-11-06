const { app } = require('@azure/functions');
const { TableClient } = require('@azure/data-tables');
let DefaultAzureCredential = null;
try { ({ DefaultAzureCredential } = require('@azure/identity')); } catch { /* optional */ }

// ✅ Helper to create a TableClient using Managed Identity or Connection String
function getTableClient(tableName) {
    const conn = process.env.TABLES_CONNECTION_STRING || process.env.AzureWebJobsStorage || '';
    const accountUrl = process.env.TABLES_ACCOUNT_URL || '';

    // 1️⃣ Prefer Managed Identity + account URL
    if (accountUrl && DefaultAzureCredential) {
        const cred = new DefaultAzureCredential();
        const client = new TableClient(accountUrl, tableName, cred);
        return {
            client,
            ensureTable: async () => {
                try {
                    await client.createTable(); // create if not exists
                } catch (e) {
                    if (!(e && e.statusCode === 409)) throw e;
                }
            }
        };
    }

    // 2️⃣ Fallback to connection string (local dev or Azurite)
    if (!conn) throw new Error('Missing Azure Table configuration. Provide TABLES_CONNECTION_STRING or TABLES_ACCOUNT_URL.');

    const client = TableClient.fromConnectionString(conn, tableName);
    return {
        client,
        ensureTable: async () => {
            try {
                await client.createTable();
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

        // ✅ Basic CORS handling
        const corsHeaders = {
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET,POST,OPTIONS',
            'Access-Control-Allow-Headers': '*',
        };

        if (request.method === 'OPTIONS') {
            return { status: 204, headers: corsHeaders };
        }

        if (request.method === 'GET') {
            try {
                const url = new URL(request.url);
                const uid = url.searchParams.get('uid');
                const category = url.searchParams.get('category');

                if (!uid) {
                    return {
                        status: 200,
                        headers: corsHeaders,
                        jsonBody: { ok: true, message: 'Supply ?uid={UID}[&category=Comments] to list saved entries.' },
                    };
                }

                const tableName = process.env.TABLES_TABLE_NAME || 'Projects';
                const { client, ensureTable } = getTableClient(tableName);
                await ensureTable();

                const filters = [`PartitionKey eq 'UID_${uid}'`];
                if (category) {
                    // Prefer lower-cased 'category' property by writer; fallback to 'Category' if present
                    filters.push(`(category eq '${category}' or Category eq '${category}')`);
                }
                const filterStr = filters.join(' and ');

                const items = [];
                for await (const entity of client.listEntities({ queryOptions: { filter: filterStr } })) {
                    items.push(entity);
                }
                // Sort newest first by rowKey (ISO timestamp) or savedAt
                items.sort((a, b) => {
                    const ak = a.rowKey || a.savedAt || '';
                    const bk = b.rowKey || b.savedAt || '';
                    return ak < bk ? 1 : ak > bk ? -1 : 0;
                });

                return {
                    status: 200,
                    headers: corsHeaders,
                    jsonBody: { ok: true, uid, category: category || null, count: items.length, items },
                };
            } catch (e) {
                return { status: 500, headers: corsHeaders, jsonBody: { ok: false, error: String(e?.message || e) } };
            }
        }

        // ✅ Parse JSON input
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
                try { return timestamp ? (new Date(timestamp)).toISOString() : new Date().toISOString(); }
                catch { return new Date().toISOString(); }
            })();

            const entity = {
                partitionKey: `UID_${uid}`,
                rowKey: nowIso,
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
            console.error('Save failed:', err);
            return { status: 500, headers: corsHeaders, jsonBody: { ok: false, error: String(err?.message || err) } };
        }
    },
});
