const { TableClient } = require("@azure/data-tables");
const crypto = require("crypto");

let DefaultAzureCredential = null;
try {
    ({ DefaultAzureCredential } = require("@azure/identity"));
} catch {}

/* =========================================================================
   AUTHENTICATION — REQUIRED FOR ALL CALLS
   ========================================================================= */
function getUser(req) {
    const header = req.headers["x-ms-client-principal"];
    if (!header) return null;

    try {
        const decoded = Buffer.from(header, "base64").toString("utf8");
        const principal = JSON.parse(decoded);

        if (!principal || !principal.userDetails) return null;

        const email = principal.userDetails.toLowerCase();

        if (!email.endsWith("@microsoft.com")) return null;

        return {
            email,
            roles: principal.userRoles || []
        };
    } catch {
        return null;
    }
}

/* =========================================================================
   HELPERS — TABLE NAME
   ========================================================================= */
function chooseTable(category) {
    if (!category) return "Projects";

    const c = category.toLowerCase();

    switch (c) {
        case "projects": return process.env.TABLES_TABLE_PROJECTS || "Projects";
        case "suggestions": return process.env.TABLES_TABLE_SUGGESTIONS || "Suggestions";
        case "troubleshooting": return process.env.TABLES_TABLE_TROUBLESHOOTING || "Troubleshooting";
        case "calendar": return process.env.TABLES_TABLE_CALENDAR || "VsoCalendar";
        case "status": return process.env.TABLES_TABLE_STATUS || "UIDStatus";
        case "notes": return process.env.TABLES_TABLE_NOTES || "Notes";
        case "comments": return process.env.TABLES_TABLE_COMMENTS || "Comments";
        case "activitylog": return process.env.TABLE_NAME_LOG || "ActivityLog";
        default: return process.env.TABLES_TABLE_DEFAULT || "Projects";
    }
}

/* =========================================================================
   TABLE CLIENT FACTORY
   ========================================================================= */
function getTableClient(tableName) {
    const accountUrl = process.env.TABLES_ACCOUNT_URL || "";
    const allowConnString = (process.env.TABLES_ALLOW_CONNECTION_STRING || "") === "1";

    if (accountUrl.startsWith("https://")) {
        if (!DefaultAzureCredential) throw new Error("Missing @azure/identity");
        const cred = new DefaultAzureCredential();
        return { client: new TableClient(accountUrl, tableName, cred) };
    }

    const conn = process.env.TABLES_CONNECTION_STRING || process.env.AzureWebJobsStorage;
    if (!allowConnString) throw new Error("Connection strings disabled");
    const client = TableClient.fromConnectionString(conn, tableName);
    return { client };
}

function getLogTableClient() {
    const tableName = process.env.TABLE_NAME_LOG || "ActivityLog";
    return getTableClient(tableName);
}

/* =========================================================================
   MAP ENTITY FROM AZURE TABLE STORAGE
   ========================================================================= */
function mapEntity(raw) {
    return {
        partitionKey: raw.partitionKey,
        rowKey: raw.rowKey,
        category: raw.category || "",
        title: raw.title || "",
        description: raw.description || "",
        owner: raw.owner || "",
        owners: raw.owners,
        savedAt: raw.savedAt || raw.Timestamp || new Date().toISOString(),
        status: raw.status,
        ...raw
    };
}

/* =========================================================================
   MAIN HTTP HANDLER
   ========================================================================= */
module.exports = async function (context, req) {
    const cors = {
        "Access-Control-Allow-Origin": "https://optical360.net",
        "Access-Control-Allow-Methods": "GET,POST,DELETE,OPTIONS",
        "Access-Control-Allow-Headers": "Content-Type,Authorization,X-MS-CLIENT-PRINCIPAL",
        "Access-Control-Allow-Credentials": "true"
    };

    if (req.method === "OPTIONS") {
        context.res = { status: 204, headers: cors };
        return;
    }

    /* ------------------ AUTH REQUIRED ------------------ */
    const user = getUser(req);
    if (!user) {
        context.res = { status: 401, headers: cors, body: { ok: false, error: "Unauthorized" } };
        return;
    }
    const isAdmin = user.email === "joshmaclean@microsoft.com";

    /* =========================================================================
       LOG HANDLER — GET & POST
       ========================================================================= */
    const urlForLog = new URL(req.url, "http://localhost");
    const pathname = urlForLog.pathname.toLowerCase();
    const isLogRequest = pathname.endsWith("/log");

    if (isLogRequest && (req.method === "GET" || req.method === "POST")) {
        try {
            const { client } = getLogTableClient();

            if (req.method === "POST") {
                const now = new Date().toISOString();
                const rowKey = crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}`;

                const entity = {
                    partitionKey: "Log",
                    rowKey,
                    email: user.email,
                    action: req.body?.action || "",
                    timestamp: now,
                    metadata: req.body?.metadata ? JSON.stringify(req.body.metadata) : ""
                };

                await client.createEntity(entity);
                context.res = { status: 200, headers: cors, body: { ok: true } };
                return;
            }

            if (req.method === "GET") {
                const items = [];
                for await (const e of client.listEntities()) items.push(mapEntity(e));
                items.sort((a, b) => b.savedAt.localeCompare(a.savedAt));

                context.res = { status: 200, headers: cors, body: { ok: true, items } };
                return;
            }

        } catch (err) {
            context.res = { status: 500, headers: cors, body: { ok: false, error: err.message } };
            return;
        }
    }

    /* =========================================================================
       GET (ALL TABLES, INCLUDING PROJECT PERMISSIONS)
       ========================================================================= */
    if (req.method === "GET") {
        try {
            const url = new URL(req.url);
            const uid = url.searchParams.get("uid");
            const category = url.searchParams.get("category") || null;

            const tableName = chooseTable(category);
            const { client } = getTableClient(tableName);

            const items = [];

            for await (const e of client.listEntities()) {
                const mapped = mapEntity(e);

                if (category === "projects") {
                    const owners = JSON.parse(mapped.owners || "[]");
                    if (!owners.includes(user.email) && !isAdmin) continue;
                }

                if (uid && mapped.partitionKey !== `UID_${uid}`) continue;

                items.push(mapped);
            }

            context.res = { status: 200, headers: cors, body: { ok: true, items } };
            return;

        } catch (err) {
            context.res = { status: 500, headers: cors, body: { ok: false, error: err.message } };
            return;
        }
    }

    /* =========================================================================
       DELETE — ADMIN OR OWNER/CO-OWNER
       ========================================================================= */
    if (req.method === "DELETE") {
        try {
            const { partitionKey, rowKey, category } = req.body || {};
            if (!partitionKey || !rowKey) {
                context.res = { status: 400, headers: cors, body: { ok: false, error: "Missing keys" } };
                return;
            }

            const tableName = chooseTable(category);
            const { client } = getTableClient(tableName);

            let existing;
            try {
                existing = await client.getEntity(partitionKey, rowKey);
            } catch {
                context.res = { status: 404, headers: cors, body: { ok: false, error: "Not found" } };
                return;
            }

            let owners = [];
            if (existing.owners) owners = JSON.parse(existing.owners);

            const isOwner = owners.includes(user.email);

            if (!isAdmin && !isOwner) {
                context.res = { status: 403, headers: cors, body: { ok: false, error: "Forbidden" } };
                return;
            }

            await client.deleteEntity(partitionKey, rowKey);
            context.res = { status: 200, headers: cors, body: { ok: true, deleted: rowKey } };
            return;

        } catch (err) {
            context.res = { status: 500, headers: cors, body: { ok: false, error: err.message } };
            return;
        }
    }

    /* =========================================================================
       POST — SUGGESTIONS UPDATE/DELETE (OWNER OR ADMIN)
       ========================================================================= */
    const body = req.body || {};
    const cat = (body.category || "").toLowerCase();
    const tableName = chooseTable(cat);
    const { client } = getTableClient(tableName);

    /* -------- Suggestions: Update Status (ANY USER) -------- */
    if (cat === "suggestions" && body.operation === "update") {
        try {
            const existing = await client.getEntity("Suggestions", body.rowKey);
            existing.status = body.status || existing.status || "New";
            await client.updateEntity(existing, "Merge");

            context.res = { status: 200, headers: cors, body: { ok: true } };
            return;

        } catch (err) {
            context.res = { status: 500, headers: cors, body: { ok: false, error: err.message } };
            return;
        }
    }

    /* -------- Suggestions: Delete (OWNER OR ADMIN) -------- */
    if (cat === "suggestions" && body.operation === "delete") {
        try {
            const existing = await client.getEntity("Suggestions", body.rowKey);

            const owner = (existing.owner || "").toLowerCase();
            if (!isAdmin && owner !== user.email) {
                context.res = { status: 403, headers: cors, body: { ok: false, error: "Forbidden" } };
                return;
            }

            await client.deleteEntity("Suggestions", body.rowKey);

            context.res = { status: 200, headers: cors, body: { ok: true } };
            return;

        } catch (err) {
            context.res = { status: 500, headers: cors, body: { ok: false, error: err.message } };
            return;
        }
    }

    /* =========================================================================
       CREATE / UPDATE ENTITY (ALL CATEGORIES)
       ========================================================================= */

    const { uid, title, description, owner } = body;

    if (!body.category) {
        context.res = { status: 400, headers: cors, body: { ok: false, error: "Missing category" } };
        return;
    }

    if (!title) {
        context.res = { status: 400, headers: cors, body: { ok: false, error: "Missing title" } };
        return;
    }

    const now = new Date().toISOString();
    const rowKey = body.rowKey || now;

    /* -------- PROJECTS (SPECIAL LOGIC) -------- */
    if (cat === "projects") {
        const owners = Array.isArray(body.owners)
            ? body.owners
            : [user.email];

        const entity = {
            partitionKey: "Projects",
            rowKey,
            category: "projects",
            title,
            description: description || "",
            owners: JSON.stringify(owners),
            savedAt: now
        };

        await client.upsertEntity(entity, "Merge");

        context.res = { status: 200, headers: cors, body: { ok: true, entity } };
        return;
    }

    /* -------- TABLES THAT REQUIRE UID -------- */
    if (["notes", "comments", "status", "troubleshooting"].includes(cat) && !uid) {
        context.res = { status: 400, headers: cors, body: { ok: false, error: "Missing UID" } };
        return;
    }

    /* -------- GENERIC ENTITY CREATION -------- */
    const safeOwner = owner?.toLowerCase() === user.email ? owner : user.email;

    const entity = {
        partitionKey: cat === "suggestions" ? "Suggestions" : `UID_${uid}`,
        rowKey,
        category: body.category,
        title,
        description: description || "",
        owner: safeOwner,
        savedAt: now
    };

    await client.upsertEntity(entity, "Merge");

    context.res = {
        status: 200,
        headers: cors,
        body: { ok: true, entity }
    };
};
