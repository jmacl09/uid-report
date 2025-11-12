const { app } = require('@azure/functions');
const { TableClient } = require('@azure/data-tables');
let DefaultAzureCredential = null;
try { ({ DefaultAzureCredential } = require('@azure/identity')); } catch { /* optional */ }

// ✅ Helper to create a TableClient using Managed Identity or Connection String
function getTableClient(tableName) {
    const conn = process.env.TABLES_CONNECTION_STRING || process.env.AzureWebJobsStorage || '';
    const accountUrl = process.env.TABLES_ACCOUNT_URL || '';
    // Prefer Managed Identity + account URL only for HTTPS endpoints.
    // Azurite/local endpoints use http:// and will reject bearer tokens; ensure we only
    // attempt DefaultAzureCredential when the accountUrl uses TLS (https://).
    const usesHttpsAccountUrl = typeof accountUrl === 'string' && accountUrl.toLowerCase().startsWith('https://');

    if (usesHttpsAccountUrl && DefaultAzureCredential) {
        const cred = new DefaultAzureCredential();
        const client = new TableClient(accountUrl, tableName, cred);
        return {
            client,
            auth: 'ManagedIdentity',
            ensureTable: async () => {
                try {
                    await client.createTable(); // create if not exists
                } catch (e) {
                    if (!(e && e.statusCode === 409)) throw e;
                }
            }
        };
    }

    // Fallback to connection string (local dev or Azurite) when Managed Identity is not suitable.
    if (!conn) throw new Error('Missing Azure Table configuration. Provide TABLES_CONNECTION_STRING or TABLES_ACCOUNT_URL (https) or AzureWebJobsStorage.');

    const client = TableClient.fromConnectionString(conn, tableName);
    return {
        client,
        auth: 'ConnectionString',
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
    methods: ['GET', 'POST', 'DELETE', 'OPTIONS'],
    authLevel: 'anonymous',
    handler: async (request, context) => {
        context.log(`Http function processed request for url "${request.url}"`);

        // ✅ Basic CORS handling
        const corsHeaders = {
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET,POST,DELETE,OPTIONS',
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

                // Determine target table. Prefer an explicit table name provided
                // by the caller (tableName/TableName/targetTable). Next, map known
                // categories (Calendar -> VsoCalendar, Troubleshooting -> Troubleshooting table env var).
                const chooseTable = (opts) => {
                    try {
                        if (opts && typeof opts === 'object') {
                            const t = opts.tableName || opts.TableName || opts.targetTable;
                            if (t) return String(t);
                            const cat = opts.category || opts.Category;
                            if (cat && String(cat).toLowerCase() === 'calendar') return process.env.TABLES_TABLE_NAME_VSO || 'VsoCalendar';
                            if (cat && String(cat).toLowerCase() === 'troubleshooting') return process.env.TABLES_TABLE_NAME_TROUBLESHOOTING || 'Troubleshooting';
                        }
                        // Fallback: if a plain category string was provided
                        if (typeof opts === 'string') {
                            const cat = opts;
                            if (cat && String(cat).toLowerCase() === 'calendar') return process.env.TABLES_TABLE_NAME_VSO || 'VsoCalendar';
                            if (cat && String(cat).toLowerCase() === 'troubleshooting') return process.env.TABLES_TABLE_NAME_TROUBLESHOOTING || 'Troubleshooting';
                        }
                    } catch (e) {
                        // ignore and fallback below
                    }
                    return process.env.TABLES_TABLE_NAME || 'Projects';
                };

                const tableName = chooseTable({ category });
                const { client, ensureTable, auth } = getTableClient(tableName);
                // Log which auth path and table is being used for easier debugging
                context.log && context.log(`[Table] GET -> table=${tableName} auth=${auth} accountUrl=${process.env.TABLES_ACCOUNT_URL ? process.env.TABLES_ACCOUNT_URL : '(using connection string)'} `);
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

        if (request.method === 'DELETE') {
            try {
                const url = new URL(request.url);
                let payload = {};
                try {
                    const txt = await request.text();
                    payload = txt ? JSON.parse(txt) : {};
                } catch { payload = {}; }

                const uid = payload.uid || payload.UID || url.searchParams.get('uid');
                const partitionKey = payload.partitionKey || payload.PartitionKey || (uid ? `UID_${uid}` : null);
                const rowKey = payload.rowKey || payload.RowKey || url.searchParams.get('rowKey');

                if (!partitionKey || !rowKey) {
                    return {
                        status: 400,
                        headers: corsHeaders,
                        jsonBody: { ok: false, error: 'Missing partitionKey and rowKey for delete.' },
                    };
                }

                // DELETE: use special VSO table for Calendar deletes when requested
                const chooseTable = (opts) => {
                    try {
                        if (opts && typeof opts === 'object') {
                            const t = opts.tableName || opts.TableName || opts.targetTable;
                            if (t) return String(t);
                            const cat = opts.category || opts.Category;
                            if (cat && String(cat).toLowerCase() === 'calendar') return process.env.TABLES_TABLE_NAME_VSO || 'VsoCalendar';
                            if (cat && String(cat).toLowerCase() === 'troubleshooting') return process.env.TABLES_TABLE_NAME_TROUBLESHOOTING || 'Troubleshooting';
                        }
                        if (typeof opts === 'string') {
                            const cat = opts;
                            if (cat && String(cat).toLowerCase() === 'calendar') return process.env.TABLES_TABLE_NAME_VSO || 'VsoCalendar';
                            if (cat && String(cat).toLowerCase() === 'troubleshooting') return process.env.TABLES_TABLE_NAME_TROUBLESHOOTING || 'Troubleshooting';
                        }
                    } catch (e) {}
                    return process.env.TABLES_TABLE_NAME || 'Projects';
                };

                const tableName = chooseTable(payload || payload.category || payload.Category || url.searchParams.get('category'));
                const { client, ensureTable, auth } = getTableClient(tableName);
                context.log && context.log(`[Table] DELETE -> table=${tableName} auth=${auth} accountUrl=${process.env.TABLES_ACCOUNT_URL ? process.env.TABLES_ACCOUNT_URL : '(using connection string)'} `);
                await ensureTable();

                try {
                    await client.deleteEntity(partitionKey, rowKey);
                } catch (err) {
                    if (err && err.statusCode === 404) {
                        return {
                            status: 404,
                            headers: corsHeaders,
                            jsonBody: { ok: false, error: 'Entity not found', partitionKey, rowKey },
                        };
                    }
                    throw err;
                }

                return {
                    status: 200,
                    headers: corsHeaders,
                    jsonBody: { ok: true, message: `Deleted ${rowKey}`, partitionKey, rowKey },
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

        const { category, uid, title, description, owner, timestamp, rowKey } = payload || {};
        if (!uid || !category || !title) {
            return { status: 400, headers: corsHeaders, jsonBody: { ok: false, error: 'Missing required fields: uid, category, title' } };
        }

        try {
            // POST (save): route Calendar category to the VSO-specific table name
            const chooseTable = (opts) => {
                try {
                    if (opts && typeof opts === 'object') {
                        const t = opts.tableName || opts.TableName || opts.targetTable;
                        if (t) return String(t);
                        const cat = opts.category || opts.Category;
                        if (cat && String(cat).toLowerCase() === 'calendar') return process.env.TABLES_TABLE_NAME_VSO || 'VsoCalendar';
                        if (cat && String(cat).toLowerCase() === 'troubleshooting') return process.env.TABLES_TABLE_NAME_TROUBLESHOOTING || 'Troubleshooting';
                    }
                    if (typeof opts === 'string') {
                        const cat = opts;
                        if (cat && String(cat).toLowerCase() === 'calendar') return process.env.TABLES_TABLE_NAME_VSO || 'VsoCalendar';
                        if (cat && String(cat).toLowerCase() === 'troubleshooting') return process.env.TABLES_TABLE_NAME_TROUBLESHOOTING || 'Troubleshooting';
                    }
                } catch (e) {}
                return process.env.TABLES_TABLE_NAME || 'Projects';
            };

            const tableName = chooseTable(payload || category);
            const { client, ensureTable, auth } = getTableClient(tableName);
            context.log && context.log(`[Table] POST -> table=${tableName} auth=${auth} accountUrl=${process.env.TABLES_ACCOUNT_URL ? process.env.TABLES_ACCOUNT_URL : '(using connection string)'} `);
            await ensureTable();

            const nowIso = (() => {
                try { return timestamp ? (new Date(timestamp)).toISOString() : new Date().toISOString(); }
                catch { return new Date().toISOString(); }
            })();
            const resolvedRowKey = (() => {
                if (rowKey && typeof rowKey === 'string' && rowKey.trim()) return rowKey.trim();
                return nowIso;
            })();

            // Build entity with canonical fields and also copy any additional payload keys
            // so callers can persist extra metadata (e.g., Status, dcCode, spans, etc.).
            const entity = {
                partitionKey: `UID_${uid}`,
                rowKey: resolvedRowKey,
                category,
                title,
                description: description || '',
                owner: owner || 'Unknown',
                savedAt: nowIso,
            };

            // Copy any other payload properties to the entity (excluding core fields)
            const coreKeys = new Set(['uid','UID','category','Category','title','Title','description','Description','owner','Owner','timestamp','Timestamp','rowKey','RowKey']);
            for (const k of Object.keys(payload || {})) {
                if (coreKeys.has(k)) continue;
                try {
                    // sanitize key names: ensure they are strings and not prototypes
                    if (typeof k === 'string' && k.trim()) entity[k] = payload[k];
                } catch (e) {}
            }

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
