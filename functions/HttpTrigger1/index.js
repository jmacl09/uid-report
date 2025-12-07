const { TableClient } = require("@azure/data-tables");
const crypto = require("crypto");

let DefaultAzureCredential = null;
try {
    ({ DefaultAzureCredential } = require("@azure/identity"));
} catch {}

/* =========================================================================
   HELPER — CATEGORY → TABLE NAME
   ========================================================================= */
function chooseTable(category) {
    if (!category) return "Projects";

    const c = category.toLowerCase();

    switch (c) {
        case "projects":
            return process.env.TABLES_TABLE_PROJECTS || "Projects";
        case "suggestions":
            return process.env.TABLES_TABLE_SUGGESTIONS || "Suggestions";
        case "troubleshooting":
            return process.env.TABLES_TABLE_TROUBLESHOOTING || "Troubleshooting";
        case "calendar":
            return process.env.TABLES_TABLE_CALENDAR || "VsoCalendar";
        case "status":
            return process.env.TABLES_TABLE_STATUS || "UIDStatus";
        case "notes":
            return process.env.TABLES_TABLE_NOTES || "Notes";
        case "comments":
            return process.env.TABLES_TABLE_COMMENTS || "Comments";
        case "activitylog":
            return process.env.TABLE_NAME_LOG || "ActivityLog";
        default:
            return process.env.TABLES_TABLE_DEFAULT || "Projects";
    }
}

/* =========================================================================
   TABLE CLIENT
   ========================================================================= */
function getTableClient(tableName) {
    const accountUrl = process.env.TABLES_ACCOUNT_URL || "";
    const allowConnString = (process.env.TABLES_ALLOW_CONNECTION_STRING || "") === "1";

    if (accountUrl.startsWith("https://")) {
        if (!DefaultAzureCredential) throw new Error("Missing @azure/identity");
        const cred = new DefaultAzureCredential();
        const client = new TableClient(accountUrl, tableName, cred);
        return { client, ensureTable: async () => {} };
    }

    const conn = process.env.TABLES_CONNECTION_STRING || process.env.AzureWebJobsStorage;
    if (!allowConnString) throw new Error("Connection strings disabled");
    if (!conn) throw new Error("Missing connection string");

    const client = TableClient.fromConnectionString(conn, tableName);
    return { client, ensureTable: async () => {} };
}

function getLogTableClient() {
    const tableName = process.env.TABLE_NAME_LOG || "ActivityLog";
    return getTableClient(tableName);
}

/* =========================================================================
   MAP ENTITY
   ========================================================================= */
function mapEntity(raw) {
    return {
        partitionKey: raw.partitionKey,
        rowKey: raw.rowKey,
        category: raw.category || "",
        title: raw.title || "",
        description: raw.description || "",
        owner: raw.owner || "",
        savedAt: raw.savedAt || raw.Timestamp || new Date().toISOString(),
        ...raw
    };
}

/* =========================================================================
   MAIN FUNCTION HANDLER
   ========================================================================= */
module.exports = async function (context, req) {
    const cors = {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "GET,POST,DELETE,OPTIONS",
        "Access-Control-Allow-Headers": "*"
    };

    if (req.method === "OPTIONS") {
        context.res = { status: 204, headers: cors };
        return;
    }

    /* =========================================================================
       LOG HANDLER
       ========================================================================= */
    const urlForLog = new URL(req.url, "http://localhost");
    const pathname = urlForLog.pathname.toLowerCase();

    const pathFromReq = (req.originalUrl || req.url || "").toString().toLowerCase();
    const headerOriginal = ((req.headers && (req.headers["x-ms-original-url"] ||
                                             req.headers["x-original-url"] ||
                                             req.headers["x-forwarded-path"])) || "").toString().toLowerCase();

    const isLogRequest =
        pathname.endsWith("/log") ||
        pathFromReq.includes("/api/log") ||
        headerOriginal.includes("/api/log");

    if ((req.method === "GET" || req.method === "POST") && isLogRequest) {
        try {
            const { client } = getLogTableClient();

            if (req.method === "POST") {
                const body = req.body || {};
                const email = (body.email || "").trim();
                const action = (body.action || "").trim();
                const metadata = body.metadata ?? null;

                if (!email || !action) {
                    context.res = { status: 400, headers: cors, body: { ok: false, error: "Missing email or action" } };
                    return;
                }

                const now = new Date().toISOString();
                const rowKey = crypto.randomUUID
                    ? crypto.randomUUID()
                    : `${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;

                const entity = {
                    partitionKey: "UID_undefined",
                    rowKey,
                    email,
                    action,
                    timestamp: now,
                    metadata: metadata ? JSON.stringify(metadata) : ""
                };

                await client.createEntity(entity);
                context.res = { status: 200, headers: cors, body: { ok: true } };
                return;
            }

            if (req.method === "GET") {
                const rawLimit = urlForLog.searchParams.get("limit");
                let limit = Number.parseInt(rawLimit || "", 10);
                if (!Number.isFinite(limit) || limit <= 0) limit = 500;

                const items = [];
                for await (const e of client.listEntities()) items.push(mapEntity(e));

                items.sort((a, b) => (a.savedAt > b.savedAt ? -1 : 1));
                context.res = {
                    status: 200,
                    headers: { ...cors, "Content-Type": "application/json" },
                    body: { ok: true, items: items.slice(0, limit) }
                };
                return;
            }
        } catch (err) {
            context.res = { status: 500, headers: cors, body: { ok: false, error: err.message } };
            return;
        }
    }

    /* =========================================================================
       GET HANDLER — INCLUDING PROJECTS
       ========================================================================= */
    if (req.method === "GET") {
        try {
            const url = new URL(req.url);
            const uid = url.searchParams.get("uid");
            const category = url.searchParams.get("category") || null;
            const tableName = chooseTable(category);
            const { client } = getTableClient(tableName);

            // ---------------------------------------------------------------
            // PROJECTS GET: only return projects where user is an owner
            // ---------------------------------------------------------------
            if (category && category.toLowerCase() === "projects") {
                const email = req.headers["x-ms-client-principal-name"] || "";
                const rows = [];

                for await (const e of client.listEntities()) {
                    const owners = e.owners ? JSON.parse(e.owners) : [];
                    if (owners.includes(email)) rows.push(mapEntity(e));
                }

                rows.sort((a, b) => (a.savedAt > b.savedAt ? -1 : 1));

                context.res = { status: 200, headers: cors, body: { ok: true, items: rows } };
                return;
            }

            // Normal GET (notes/suggestions/status/etc)
            if (!uid) {
                const items = [];
                for await (const e of client.listEntities()) items.push(mapEntity(e));

                context.res = { status: 200, headers: cors, body: { ok: true, items } };
                return;
            }

            const filter = [`PartitionKey eq 'UID_${uid}'`];
            if (category) filter.push(`category eq '${category}'`);

            const items = [];
            for await (const e of client.listEntities({ queryOptions: { filter: filter.join(" and ") } })) {
                items.push(mapEntity(e));
            }

            context.res = { status: 200, headers: cors, body: { ok: true, items } };
            return;
        } catch (err) {
            context.res = { status: 500, headers: cors, body: { ok: false, error: err.message } };
            return;
        }
    }

    /* =========================================================================
       DELETE — INCLUDING PROJECTS
       ========================================================================= */
    if (req.method === "DELETE") {
        try {
            const { category, rowKey } = req.body || {};
            const tableName = chooseTable(category);
            const { client } = getTableClient(tableName);

            // --------------------------
            // PROJECT DELETE RULES
            // --------------------------
            if (category.toLowerCase() === "projects") {
                await client.deleteEntity("Projects", rowKey);
                context.res = { status: 200, headers: cors, body: { ok: true, deleted: rowKey } };
                return;
            }

            // Default delete
            const { partitionKey } = req.body;
            if (!partitionKey || !rowKey) {
                context.res = { status: 400, headers: cors, body: { ok: false, error: "Missing partitionKey or rowKey" } };
                return;
            }

            await client.deleteEntity(partitionKey, rowKey);
            context.res = { status: 200, headers: cors, body: { ok: true, deleted: rowKey } };
            return;
        } catch (err) {
            context.res = {
                status: 500,
                headers: cors,
                body: { ok: false, error: err.message }
            };
            return;
        }
    }

    /* =========================================================================
       POST — PROJECTS + EXISTING LOGIC
       ========================================================================= */

    const body = req.body || {};
    const { category, title, description } = body;
    const cat = (category || "").toLowerCase();
    const tableName = chooseTable(cat);
    const { client } = getTableClient(tableName);

    // ---------------------------------------------------------------
    //  PROJECT CREATE / UPDATE
    // ---------------------------------------------------------------
    if (cat === "projects") {
        try {
            const projectId = body.rowKey || crypto.randomUUID();
            const owners = body.owners ? body.owners : [body.owner];
            const uids = body.uids || [];

            const entity = {
                partitionKey: "Projects",
                rowKey: projectId,
                category: "Projects",
                title,
                description: description || "",
                owners: JSON.stringify(owners),
                uids: JSON.stringify(uids),
                savedAt: new Date().toISOString()
            };

            await client.upsertEntity(entity, "Merge");

            context.res = {
                status: 200,
                headers: cors,
                body: { ok: true, entity }
            };
            return;
        } catch (err){
            context.res = { status: 500, headers: cors, body: { ok: false, error: err.message } };
            return;
        }
    }

    /* =========================================================================
       ORIGINAL CREATE LOGIC (NOT USED FOR PROJECTS)
       ========================================================================= */
    if (!category) {
        context.res = { status: 400, headers: cors, body: { ok: false, error: "Missing category" } };
        return;
    }

    if (!title) {
        context.res = { status: 400, headers: cors, body: { ok: false, error: "Missing title" } };
        return;
    }

    const uid = body.uid;
    if (["notes", "comments", "status", "troubleshooting"].includes(cat) && !uid) {
        context.res = { status: 400, headers: cors, body: { ok: false, error: "Missing UID" } };
        return;
    }

    const now = new Date().toISOString();
    const rowKey = body.rowKey || now;

    const entity = {
        partitionKey: cat === "suggestions" ? "Suggestions" : `UID_${uid}`,
        rowKey,
        category,
        title,
        description: description || "",
        owner: body.owner || "Unknown",
        savedAt: now
    };

    await client.upsertEntity(entity, "Merge");

    context.res = {
        status: 200,
        headers: cors,
        body: { ok: true, entity }
    };
};
