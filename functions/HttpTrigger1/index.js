const { TableClient } = require("@azure/data-tables");
const crypto = require("crypto");

let DefaultAzureCredential = null;
try {
    ({ DefaultAzureCredential } = require("@azure/identity"));
} catch { }

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

        /* ⭐ NEW: ActivityLog support ⭐ */
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

    // Managed Identity
    if (accountUrl.startsWith("https://")) {
        if (!DefaultAzureCredential) throw new Error("Missing @azure/identity");
        const cred = new DefaultAzureCredential();
        const client = new TableClient(accountUrl, tableName, cred);
        return { client, ensureTable: async () => {} };
    }

    // Connection string fallback
    const conn = process.env.TABLES_CONNECTION_STRING || process.env.AzureWebJobsStorage;
    if (!allowConnString) throw new Error("Connection strings disabled");
    if (!conn) throw new Error("Missing connection string");

    const client = TableClient.fromConnectionString(conn, tableName);
    return { client, ensureTable: async () => {} };
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

    /* OPTIONS */
    if (req.method === "OPTIONS") {
        context.res = { status: 204, headers: cors };
        return;
    }

    /* =============== GET =================== */
    if (req.method === "GET") {
        try {
            const url = new URL(req.url);
            const uid = url.searchParams.get("uid");
            const category = url.searchParams.get("category") || null;

            const tableName = chooseTable(category);
            const { client } = getTableClient(tableName);

            // GET all if no UID is supplied
            if (!uid) {
                const items = [];
                for await (const e of client.listEntities()) {
                    items.push(mapEntity(e));
                }

                context.res = { status: 200, headers: cors, body: { ok: true, items } };
                return;
            }

            // GET filtered by UID
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

    /* =============== DELETE =================== */
    if (req.method === "DELETE") {
        try {
            const body = req.body || {};
            const { partitionKey, rowKey, category } = body;

            if (!partitionKey || !rowKey) {
                context.res = { status: 400, headers: cors, body: { ok: false, error: "Missing partitionKey or rowKey" } };
                return;
            }

            const tableName = chooseTable(category);
            const { client } = getTableClient(tableName);

            await client.deleteEntity(partitionKey, rowKey);
            context.res = { status: 200, headers: cors, body: { ok: true, deleted: rowKey } };
            return;

        } catch (err) {
            context.res = { status: 500, headers: cors, body: { ok: false, error: err.message } };
            return;
        }
    }

    /* =============== POST =================== */
    const body = req.body || {};
    const { uid, category, title, description, owner } = body;

    if (!category) {
        context.res = { status: 400, headers: cors, body: { ok: false, error: "Missing category" } };
        return;
    }

    const cat = category.toLowerCase();
    const tableName = chooseTable(cat);
    const { client } = getTableClient(tableName);

    if (!title) {
        context.res = { status: 400, headers: cors, body: { ok: false, error: "Missing title" } };
        return;
    }

    // UID-required categories
    if (["notes", "comments", "projects", "status", "troubleshooting"].includes(cat) && !uid) {
        context.res = { status: 400, headers: cors, body: { ok: false, error: "Missing UID" } };
        return;
    }

    /* =========================================================================
       SPECIAL HANDLING: ActivityLog DOES NOT USE UID
       ========================================================================= */
    let partitionKey = cat === "activitylog" 
        ? "UserActivity"      // cleaner PK for logging entries
        : (cat === "suggestions" ? "Suggestions" : `UID_${uid}`);

    /* Build entity */
    const now = new Date().toISOString();
    const rowKey = body.rowKey || now;

    const entity = {
        partitionKey,
        rowKey,
        category,
        title,
        description: description || "",
        owner: owner || "Unknown",
        savedAt: now
    };

    await client.upsertEntity(entity, "Merge");

    context.res = { status: 200, headers: cors, body: { ok: true, entity } };
};
