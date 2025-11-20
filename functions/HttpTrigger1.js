const { app } = require('@azure/functions');
const { TableClient } = require('@azure/data-tables');
let DefaultAzureCredential = null;
try { ({ DefaultAzureCredential } = require('@azure/identity')); } catch { /* optional */ }

/* -------------------------------------------------------------------------
   GLOBAL: chooseTable()
   Shared by GET, POST, DELETE (fixes GET 500 error)
   ------------------------------------------------------------------------- */
function chooseTable(opts) {
    try {
        if (opts && typeof opts === 'object') {
            const t = opts.tableName || opts.TableName || opts.targetTable;
            if (t) return String(t);

            const cat = opts.category || opts.Category;
            if (cat && cat.toLowerCase() === 'calendar')        return process.env.TABLES_TABLE_NAME_VSO || 'VsoCalendar';
            if (cat && cat.toLowerCase() === 'troubleshooting') return process.env.TABLES_TABLE_NAME_TROUBLESHOOTING || 'Troubleshooting';
            if (cat && cat.toLowerCase() === 'suggestions')     return process.env.TABLES_TABLE_NAME_SUGGESTIONS || 'Suggestions';
            if (cat && cat.toLowerCase() === 'status')          return process.env.TABLES_TABLE_NAME_STATUS || 'UIDStatus';
            if (cat && cat.toLowerCase() === 'projects')        return process.env.TABLES_TABLE || process.env.TABLES_TABLE_NAME || 'Projects';
        }

        if (typeof opts === 'string') {
            const cat = opts.toLowerCase();
            if (cat === 'calendar')        return process.env.TABLES_TABLE_NAME_VSO || 'VsoCalendar';
            if (cat === 'troubleshooting') return process.env.TABLES_TABLE_NAME_TROUBLESHOOTING || 'Troubleshooting';
            if (cat === 'suggestions')     return process.env.TABLES_TABLE_NAME_SUGGESTIONS || 'Suggestions';
            if (cat === 'projects')        return process.env.TABLES_TABLE || process.env.TABLES_TABLE_NAME || 'Projects';
            if (cat === 'status')          return process.env.TABLES_TABLE_NAME_STATUS || 'UIDStatus';
        }
    } catch { }

    return process.env.TABLES_TABLE || process.env.TABLES_TABLE_NAME || 'Projects';
}

/* -------------------------------------------------------------------------
   GLOBAL: getTableClient()
   ------------------------------------------------------------------------- */
function getTableClient(tableName) {
    const accountUrl = process.env.TABLES_ACCOUNT_URL || '';
    const allowConnString = (process.env.TABLES_ALLOW_CONNECTION_STRING || '').toString() === '1';

    if (accountUrl.toLowerCase().startsWith('https://')) {
        if (!DefaultAzureCredential) {
            throw new Error('Managed Identity not available. Install @azure/identity.');
        }

        const cred = new DefaultAzureCredential();
        const client = new TableClient(accountUrl, tableName, cred);

        return {
            client,
            auth: 'ManagedIdentity',
            ensureTable: async () => {
                try { await client.createTable(); }
                catch (e) { if (!(e && e.statusCode === 409)) throw e; }
            }
        };
    }

    const conn = process.env.TABLES_CONNECTION_STRING || process.env.AzureWebJobsStorage || '';
    if (!allowConnString) {
        throw new Error('Managed Identity not configured AND connection strings disabled.');
    }

    if (!conn) throw new Error('No connection string available.');

    const client = TableClient.fromConnectionString(conn, tableName);

    return {
        client,
        auth: 'ConnectionString',
        ensureTable: async () => {
            try { await client.createTable(); }
            catch (e) { if (!(e && e.statusCode === 409)) throw e; }
        }
    };
}

/* -------------------------------------------------------------------------
   MAIN HTTP TRIGGER
   ------------------------------------------------------------------------- */
app.http('HttpTrigger1', {
    methods: ['GET', 'POST', 'DELETE', 'OPTIONS'],
    authLevel: 'anonymous',

    handler: async (request, context) => {

        // Basic CORS
        const corsHeaders = {
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET,POST,DELETE,OPTIONS',
            'Access-Control-Allow-Headers': '*',
        };

        if (request.method === 'OPTIONS') {
            return { status: 204, headers: corsHeaders };
        }

        /* -----------------------------------------------------------------
           GET
           ----------------------------------------------------------------- */
        if (request.method === 'GET') {
            try {
                const url = new URL(request.url);
                const category = url.searchParams.get('category');
                const qTable = url.searchParams.get('tableName');
                const uid = url.searchParams.get('uid');

                // Always use Suggestions table for suggestions (ignore UID)
                if ((category || '').toLowerCase() === 'suggestions' || qTable === 'Suggestions') {
                    const tableName = 'Suggestions';
                    const { client, ensureTable, auth } = getTableClient(tableName);
                    context.log(`[Table] GET all (Suggestions) table=${tableName} auth=${auth}`);
                    await ensureTable();
                    const items = [];
                    for await (const rawEntity of client.listEntities()) {
                        const entity = {};
                        try {
                            entity.partitionKey = rawEntity.partitionKey || rawEntity.PartitionKey || rawEntity.PK || rawEntity.Partition || '';
                            entity.rowKey = rawEntity.rowKey || rawEntity.RowKey || rawEntity.RK || rawEntity.row || '';
                            entity.title = rawEntity.title || rawEntity.Title || rawEntity.projectName || rawEntity.ProjectName || '';
                            entity.description = rawEntity.description || rawEntity.Description || '';
                            entity.owner = rawEntity.owner || rawEntity.Owner || '';
                            entity.savedAt = rawEntity.savedAt || rawEntity.savedAt || rawEntity.Timestamp || rawEntity.timestamp || entity.rowKey || new Date().toISOString();
                            for (const k of Object.keys(rawEntity || {})) {
                                if (!Object.prototype.hasOwnProperty.call(entity, k)) {
                                    try { entity[k] = rawEntity[k]; } catch { }
                                }
                            }
                        } catch (e) {
                            try { Object.assign(entity, rawEntity); } catch { }
                        }
                        items.push(entity);
                    }
                    items.sort((a, b) => (a.rowKey < b.rowKey ? 1 : -1));
                    return {
                        status: 200,
                        headers: corsHeaders,
                        jsonBody: { ok: true, category: 'suggestions', count: items.length, items }
                    };
                }

                // ...existing code for other categories...
                // If listing category without UID
                if (!uid) {
                    const catLower = (category || '').toLowerCase();
                    if (catLower === 'projects') {
                        const tableName = chooseTable({ tableName: qTable, category });
                        const { client, ensureTable, auth } = getTableClient(tableName);
                        context.log(`[Table] GET all (${category}) table=${tableName} auth=${auth}`);
                        await ensureTable();
                        const items = [];
                        for await (const rawEntity of client.listEntities()) {
                            const entity = {};
                            try {
                                entity.partitionKey = rawEntity.partitionKey || rawEntity.PartitionKey || rawEntity.PK || rawEntity.Partition || '';
                                entity.rowKey = rawEntity.rowKey || rawEntity.RowKey || rawEntity.RK || rawEntity.row || '';
                                entity.category = rawEntity.category || rawEntity.Category || '';
                                entity.title = rawEntity.title || rawEntity.Title || rawEntity.projectName || rawEntity.ProjectName || '';
                                entity.description = rawEntity.description || rawEntity.Description || '';
                                entity.owner = rawEntity.owner || rawEntity.Owner || '';
                                entity.savedAt = rawEntity.savedAt || rawEntity.savedAt || rawEntity.Timestamp || rawEntity.timestamp || entity.rowKey || new Date().toISOString();
                                for (const k of Object.keys(rawEntity || {})) {
                                    if (!Object.prototype.hasOwnProperty.call(entity, k)) {
                                        try { entity[k] = rawEntity[k]; } catch { }
                                    }
                                }
                            } catch (e) {
                                try { Object.assign(entity, rawEntity); } catch { }
                            }
                            items.push(entity);
                        }
                        items.sort((a, b) => (a.rowKey < b.rowKey ? 1 : -1));
                        return {
                            status: 200,
                            headers: corsHeaders,
                            jsonBody: { ok: true, category, count: items.length, items }
                        };
                    }
                    return {
                        status: 200,
                        headers: corsHeaders,
                        jsonBody: { ok: true, message: 'Supply ?uid=UID to filter entries.' }
                    };
                }

                // Normal filtered GET
                const tableName = chooseTable({ tableName: qTable, category });
                const { client, ensureTable, auth } = getTableClient(tableName);
                context.log(`[Table] GET -> table=${tableName} auth=${auth}`);
                await ensureTable();
                const filter = [`PartitionKey eq 'UID_${uid}'`];
                if (category) filter.push(`(category eq '${category}' or Category eq '${category}')`);
                const items = [];
                for await (const rawEntity of client.listEntities({ queryOptions: { filter: filter.join(' and ') } })) {
                    const entity = {};
                    try {
                        entity.partitionKey = rawEntity.partitionKey || rawEntity.PartitionKey || rawEntity.PK || rawEntity.Partition || '';
                        entity.rowKey = rawEntity.rowKey || rawEntity.RowKey || rawEntity.RK || rawEntity.row || '';
                        entity.category = rawEntity.category || rawEntity.Category || '';
                        entity.title = rawEntity.title || rawEntity.Title || '';
                        entity.description = rawEntity.description || rawEntity.Description || '';
                        entity.owner = rawEntity.owner || rawEntity.Owner || '';
                        entity.savedAt = rawEntity.savedAt || rawEntity.savedAt || rawEntity.Timestamp || rawEntity.timestamp || entity.rowKey || new Date().toISOString();
                        for (const k of Object.keys(rawEntity || {})) {
                            if (!Object.prototype.hasOwnProperty.call(entity, k)) {
                                try { entity[k] = rawEntity[k]; } catch { }
                            }
                        }
                    } catch (e) {
                        try { Object.assign(entity, rawEntity); } catch { }
                    }
                    items.push(entity);
                }
                items.sort((a, b) => (a.rowKey < b.rowKey ? 1 : -1));
                return {
                    status: 200,
                    headers: corsHeaders,
                    jsonBody: { ok: true, uid, category, count: items.length, items }
                };
            } catch (err) {
                context.log('GET ERROR:', err);
                return { status: 500, headers: corsHeaders, jsonBody: { ok: false, error: String(err) } };
            }
        }

        /* -----------------------------------------------------------------
           DELETE
           ----------------------------------------------------------------- */
        if (request.method === 'DELETE') {
            try {
                const url = new URL(request.url);
                let payload = {};

                try { payload = JSON.parse(await request.text()); } catch { }

                const uid = payload.uid || url.searchParams.get('uid');
                const partitionKey = payload.partitionKey || (uid ? `UID_${uid}` : null);
                const rowKey = payload.rowKey || url.searchParams.get('rowKey');
                const category = payload.category || url.searchParams.get('category');
                const tableName = chooseTable({ category });

                if (!partitionKey || !rowKey) {
                    return { status: 400, headers: corsHeaders, jsonBody: { ok: false, error: 'Missing partitionKey or rowKey.' } };
                }

                const { client, ensureTable, auth } = getTableClient(tableName);
                context.log(`[Table] DELETE -> table=${tableName} auth=${auth}`);
                await ensureTable();

                await client.deleteEntity(partitionKey, rowKey);

                return {
                    status: 200,
                    headers: corsHeaders,
                    jsonBody: { ok: true, message: `Deleted ${rowKey}` }
                };

            } catch (err) {
                return { status: 500, headers: corsHeaders, jsonBody: { ok: false, error: String(err) } };
            }
        }

        /* -----------------------------------------------------------------
           POST (save)
           ----------------------------------------------------------------- */
        let payload;
        try { payload = await request.json(); }
        catch (e) {
            return { status: 400, headers: corsHeaders, jsonBody: { ok: false, error: 'Invalid JSON' } };
        }

        const { category, uid, title, description, owner, timestamp, rowKey } = payload;
        const catLower = (category || '').toLowerCase();

        // Always use Suggestions table for suggestions, and do not require UID or category
        if (catLower === 'suggestions' || !category) {
            if (!title) {
                return { status: 400, headers: corsHeaders, jsonBody: { ok: false, error: 'Missing title.' } };
            }
            try {
                const tableName = 'Suggestions';
                const { client, ensureTable, auth } = getTableClient(tableName);
                context.log(`[Table] POST -> table=${tableName} auth=${auth}`);
                await ensureTable();
                const nowIso = timestamp ? new Date(timestamp).toISOString() : new Date().toISOString();
                const resolvedRowKey = (rowKey && rowKey.trim()) || nowIso;
                const entity = {
                    partitionKey: 'Suggestions',
                    rowKey: resolvedRowKey,
                    title,
                    description: description || '',
                    owner: owner || 'Unknown',
                    savedAt: nowIso,
                };
                await client.upsertEntity(entity, 'Merge');
                return {
                    status: 200,
                    headers: corsHeaders,
                    jsonBody: { ok: true, message: `Saved suggestion`, entity }
                };
            } catch (err) {
                context.log('POST ERROR:', err);
                return { status: 500, headers: corsHeaders, jsonBody: { ok: false, error: String(err) } };
            }
        }

        // ...existing code for other categories...
        if (!category || !title) {
            return { status: 400, headers: corsHeaders, jsonBody: { ok: false, error: 'Missing category or title.' } };
        }
        if (!uid) {
            return { status: 400, headers: corsHeaders, jsonBody: { ok: false, error: 'Missing UID.' } };
        }
        try {
            const tableName = chooseTable({ category });
            const { client, ensureTable, auth } = getTableClient(tableName);
            context.log(`[Table] POST -> table=${tableName} auth=${auth}`);
            await ensureTable();
            const nowIso = timestamp ? new Date(timestamp).toISOString() : new Date().toISOString();
            const resolvedRowKey = (rowKey && rowKey.trim()) || nowIso;
            const entity = {
                partitionKey: `UID_${uid}`,
                rowKey: resolvedRowKey,
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
                jsonBody: { ok: true, message: `Saved ${category}`, entity }
            };
        } catch (err) {
            context.log('POST ERROR:', err);
            return { status: 500, headers: corsHeaders, jsonBody: { ok: false, error: String(err) } };
        }
    }
});
