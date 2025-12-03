const { TableClient } = require("@azure/data-tables");
const crypto = require("crypto");

let DefaultAzureCredential = null;
try { ({ DefaultAzureCredential } = require("@azure/identity")); } catch { }

/* -------------------------------------------------------------------------
   EXPORT SIGNATURE
   ------------------------------------------------------------------------- */
module.exports = async function (context, req) {
    return await handleRequest(req, context);
};

/* -------------------------------------------------------------------------
   HELPERS
   ------------------------------------------------------------------------- */
function chooseTable(opts) {
    try {
        if (opts && typeof opts === "object") {
            const t = opts.tableName || opts.TableName || opts.targetTable;
            if (t) return String(t);

            const cat = opts.category || opts.Category;
            if (cat) {
                const c = cat.toLowerCase();
                if (c === "calendar") return process.env.TABLES_TABLE_NAME_VSO || "VsoCalendar";
                if (c === "troubleshooting") return process.env.TABLES_TABLE_NAME_TROUBLESHOOTING || "Troubleshooting";
                if (c === "suggestions") return process.env.TABLES_TABLE_NAME_SUGGESTIONS || "Suggestions";
                if (c === "projects") return process.env.TABLES_TABLE || process.env.TABLES_TABLE_NAME || "Projects";
                if (c === "status") return process.env.TABLES_TABLE_NAME_STATUS || "UIDStatus";
            }
        }

        if (typeof opts === "string") {
            const c = opts.toLowerCase();
            if (c === "calendar") return process.env.TABLES_TABLE_NAME_VSO || "VsoCalendar";
            if (c === "troubleshooting") return process.env.TABLES_TABLE_NAME_TROUBLESHOOTING || "Troubleshooting";
            if (c === "suggestions") return process.env.TABLES_TABLE_NAME_SUGGESTIONS || "Suggestions";
            if (c === "projects") return process.env.TABLES_TABLE || process.env.TABLES_TABLE_NAME || "Projects";
            if (c === "status") return process.env.TABLES_TABLE_NAME_STATUS || "UIDStatus";
        }
    } catch {}

    return process.env.TABLES_TABLE || process.env.TABLES_TABLE_NAME || "Projects";
}

/* ----------------------- FIXED VERSION (NO createTable) --------------------- */
function getTableClient(tableName) {
    const accountUrl = process.env.TABLES_ACCOUNT_URL || "";
    const allowConnString = (process.env.TABLES_ALLOW_CONNECTION_STRING || "").toString() === "1";

    // Managed Identity path
    if (accountUrl.toLowerCase().startsWith("https://")) {
        if (!DefaultAzureCredential) {
            throw new Error("Managed Identity not available — install @azure/identity.");
        }

        const cred = new DefaultAzureCredential();
        const client = new TableClient(accountUrl, tableName, cred);

        return {
            client,
            auth: "ManagedIdentity",
            ensureTable: async () => {} // DO NOTHING — MI cannot create tables
        };
    }

    // Local dev / Connection string fallback
    const conn = process.env.TABLES_CONNECTION_STRING || process.env.AzureWebJobsStorage || "";
    if (!allowConnString) throw new Error("Connection strings disabled & MI not configured.");
    if (!conn) throw new Error("No connection string available.");

    const client = TableClient.fromConnectionString(conn, tableName);

    return {
        client,
        auth: "ConnectionString",
        ensureTable: async () => {} // DO NOTHING
    };
}

/* -------------------------------------------------------------------------
   MAIN HANDLER
   ------------------------------------------------------------------------- */
async function handleRequest(request, context) {
    const corsHeaders = {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "GET,POST,DELETE,OPTIONS",
        "Access-Control-Allow-Headers": "*"
    };

    /* ------------------ OPTIONS ------------------ */
    if (request.method === "OPTIONS") {
        context.res = { status: 204, headers: corsHeaders };
        return;
    }

    /* ------------------ GET ------------------ */
    if (request.method === "GET") {
        try {
            const url = new URL(request.url);
            const category = url.searchParams.get("category");
            const qTable = url.searchParams.get("tableName");
            const uid = url.searchParams.get("uid");

            /* SUGGESTIONS: GET ALL */
            if ((category || "").toLowerCase() === "suggestions" || qTable === "Suggestions") {
                const tableName = "Suggestions";
                const { client, auth } = getTableClient(tableName);

                context.log(`[Table] GET ALL -> ${tableName} auth=${auth}`);

                const items = [];
                for await (const raw of client.listEntities()) {
                    items.push(mapEntity(raw));
                }

                // Sort by savedAt (newest first) if available
                items.sort((a, b) => (a.savedAt && b.savedAt ? (a.savedAt < b.savedAt ? 1 : -1) : (a.rowKey < b.rowKey ? 1 : -1)));

                context.res = { status: 200, headers: corsHeaders, body: { ok: true, items } };
                return;
            }

            /* GET ALL for Category (Projects only) */
            if (!uid) {
                if ((category || "").toLowerCase() === "projects") {
                    const tableName = chooseTable({ tableName: qTable, category });
                    const { client, auth } = getTableClient(tableName);

                    context.log(`[Table] GET ALL -> ${tableName} auth=${auth}`);

                    const items = [];
                    for await (const raw of client.listEntities()) {
                        items.push(mapEntity(raw));
                    }
                    items.sort((a, b) => (a.rowKey < b.rowKey ? 1 : -1));

                    context.res = { status: 200, headers: corsHeaders, body: { ok: true, items } };
                    return;
                }

                context.res = { status: 200, headers: corsHeaders, body: { ok: true, message: "Supply ?uid=UID" } };
                return;
            }

            /* NORMAL FILTERED GET */
            const tableName = chooseTable({ tableName: qTable, category });
            const { client, auth } = getTableClient(tableName);

            context.log(`[Table] GET -> table=${tableName} auth=${auth}`);

            const filter = [`PartitionKey eq 'UID_${uid}'`];
            if (category) filter.push(`(category eq '${category}' or Category eq '${category}')`);

            const items = [];
            for await (const raw of client.listEntities({
                queryOptions: { filter: filter.join(" and ") }
            })) {
                items.push(mapEntity(raw));
            }

            items.sort((a, b) => (a.rowKey < b.rowKey ? 1 : -1));

            context.res = {
                status: 200,
                headers: corsHeaders,
                body: { ok: true, uid, category, count: items.length, items }
            };
            return;

        } catch (err) {
            context.log("GET ERROR:", err);
            context.res = { status: 500, headers: corsHeaders, body: { ok: false, error: String(err) } };
            return;
        }
    }

    /* ------------------ DELETE ------------------ */
    if (request.method === "DELETE") {
        try {
            const url = new URL(request.url);
            let payload = {};
            try { payload = JSON.parse(await request.text()); } catch {}

            const uid = payload.uid || url.searchParams.get("uid");
            const partitionKey = payload.partitionKey || (uid ? `UID_${uid}` : null);
            const rowKey = payload.rowKey || url.searchParams.get("rowKey");
            const category = payload.category || url.searchParams.get("category");

            if (!partitionKey || !rowKey) {
                context.res = { status: 400, headers: corsHeaders, body: { ok: false, error: "Missing partitionKey or rowKey" } };
                return;
            }

            const tableName = chooseTable({ category });
            const { client, auth } = getTableClient(tableName);

            context.log(`[Table] DELETE -> ${tableName} auth=${auth}`);

            await client.deleteEntity(partitionKey, rowKey);

            context.res = { status: 200, headers: corsHeaders, body: { ok: true, deleted: rowKey } };
            return;

        } catch (err) {
            context.res = { status: 500, headers: corsHeaders, body: { ok: false, error: String(err) } };
            return;
        }
    }

    /* ------------------ POST ------------------ */
    let payload;
    try { payload = request.body || {}; }
    catch { 
        context.res = { status: 400, headers: corsHeaders, body: { ok: false, error: "Invalid JSON" } };
        return;
    }

    const { category, uid, title, description, owner, timestamp, rowKey } = payload;
    const catLower = (category || "").toLowerCase();

    /* ------------------ TROUBLESHOOTING POST ------------------ */
    if (catLower === "troubleshooting") {
        if (!uid || !description) {
            context.res = { status: 400, headers: corsHeaders, body: { ok: false, error: "Missing uid or description" } };
            return;
        }

        try {
            const tableName = "Troubleshooting";
            const { client, auth } = getTableClient(tableName);

            context.log(`[Table] POST -> ${tableName} auth=${auth}`);

            const nowIso = new Date().toISOString();
            const resolvedRowKey = rowKey?.trim() || nowIso;

            const entity = {
                partitionKey: `UID_${uid}`,
                rowKey: resolvedRowKey,
                category: "Troubleshooting",
                title: title || "Troubleshooting Entry",
                description,
                owner: owner || "Unknown",
                savedAt: nowIso
            };

            await client.upsertEntity(entity, "Merge");

            context.res = { status: 200, headers: corsHeaders, body: { ok: true, entity } };
            return;

        } catch (err) {
            context.log("POST ERROR:", err);
            context.res = { status: 500, headers: corsHeaders, body: { ok: false, error: String(err) } };
            return;
        }
    }

    /* ------------------ SUGGESTIONS ------------------ */
    if (catLower === "suggestions" || !category) {
        if (!title) {
            context.res = { status: 400, headers: corsHeaders, body: { ok: false, error: "Missing title" } };
            return;
        }

        try {
            const tableName = "Suggestions";
            const { client, auth } = getTableClient(tableName);

            context.log(`[Table] POST -> ${tableName} auth=${auth}`);

            const nowIso = timestamp ? new Date(timestamp).toISOString() : new Date().toISOString();

            // Try to extract submitter email from x-ms-client-principal header (App Service / Easy Auth)
            let submitterEmail = null;
            try {
                const principalHeader = request.headers && (request.headers['x-ms-client-principal'] || request.headers['X-MS-CLIENT-PRINCIPAL']);
                if (principalHeader) {
                    const buff = Buffer.from(principalHeader, 'base64');
                    const parsed = JSON.parse(buff.toString('utf8'));
                    submitterEmail = parsed && (parsed.userDetails || parsed.userId || parsed.identity || null);
                    if (!submitterEmail && parsed && Array.isArray(parsed.identities) && parsed.identities.length) {
                        const id = parsed.identities[0];
                        if (id && Array.isArray(id.claims)) {
                            const emailClaim = id.claims.find(c => (c.typ || c.type || '').toLowerCase().includes('email') || (c.typ || c.type || '').toLowerCase().includes('emails'));
                            if (emailClaim) submitterEmail = emailClaim.val || emailClaim.value || submitterEmail;
                        }
                    }
                }
            } catch { }

            const uniqueId = (crypto && crypto.randomUUID) ? crypto.randomUUID() : `${Date.now()}-${Math.random().toString(36).slice(2,8)}`;
            const resolvedRowKey = rowKey?.trim() || uniqueId;

            // Accept additional suggestion fields: type, name, summary, comment
            const entity = {
                partitionKey: "Suggestions",
                rowKey: resolvedRowKey,
                id: resolvedRowKey,
                title,
                type: payload.type || payload.suggestionType || "",
                name: payload.name || "",
                summary: payload.summary || "",
                comment: payload.comment || description || "",
                description: description || "",
                submitterEmail: submitterEmail || owner || "Unknown",
                owner: submitterEmail || owner || "Unknown",
                savedAt: nowIso
            };

            await client.upsertEntity(entity, "Merge");

            context.res = { status: 200, headers: corsHeaders, body: { ok: true, entity } };
            return;

        } catch (err) {
            context.log("POST ERROR:", err);
            context.res = { status: 500, headers: corsHeaders, body: { ok: false, error: String(err) } };
            return;
        }
    }

    /* ------------------ OTHER CATEGORY SAVES ------------------ */
    if (!category || !title) {
        context.res = { status: 400, headers: corsHeaders, body: { ok: false, error: "Missing category or title" } };
        return;
    }
    if (!uid) {
        context.res = { status: 400, headers: corsHeaders, body: { ok: false, error: "Missing UID" } };
        return;
    }

    try {
        const tableName = chooseTable({ category });
        const { client, auth } = getTableClient(tableName);

        context.log(`[Table] POST -> ${tableName} auth=${auth}`);

        const nowIso = timestamp ? new Date(timestamp).toISOString() : new Date().toISOString();
        const resolvedRowKey = rowKey?.trim() || nowIso;

        const entity = {
            partitionKey: `UID_${uid}`,
            rowKey: resolvedRowKey,
            category,
            title,
            description: description || "",
            owner: owner || "Unknown",
            savedAt: nowIso
        };

        await client.upsertEntity(entity, "Merge");

        context.res = { status: 200, headers: corsHeaders, body: { ok: true, entity } };
        return;

    } catch (err) {
        context.log("POST ERROR:", err);
        context.res = { status: 500, headers: corsHeaders, body: { ok: false, error: String(err) } };
        return;
    }
}

/* -------------------------------------------------------------------------
   MAP ENTITY
   ------------------------------------------------------------------------- */
function mapEntity(raw) {
    const out = {};

    out.partitionKey = raw.partitionKey || raw.PartitionKey || "";
    out.rowKey = raw.rowKey || raw.RowKey || "";
    out.category = raw.category || raw.Category || "";
    out.title = raw.title || raw.Title || "";
    out.description = raw.description || raw.Description || "";
    out.owner = raw.owner || raw.Owner || "";
    out.savedAt = raw.savedAt || raw.Timestamp || raw.timestamp || new Date().toISOString();

    for (const k of Object.keys(raw)) {
        if (!Object.prototype.hasOwnProperty.call(out, k)) {
            out[k] = raw[k];
        }
    }

    return out;
}
