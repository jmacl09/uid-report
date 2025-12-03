const { TableClient } = require("@azure/data-tables");
const { DefaultAzureCredential } = require("@azure/identity");
const crypto = require("crypto");

// --------------------- HELPERS ---------------------
function getTableClient(tableName) {
    const allowConnString = process.env.TABLES_ALLOW_CONNECTION_STRING;
    const connStr = process.env.TABLE_CONNECTION;

    if (allowConnString === '1' && connStr) {
        return TableClient.fromConnectionString(connStr, tableName);
    }

    const url = process.env.TABLES_ACCOUNT_URL;
    if (!url) throw new Error("TABLES_ACCOUNT_URL missing");

    const credential = new DefaultAzureCredential();
    return new TableClient(url, tableName, credential);
}

function map(raw) {
    return {
        partitionKey: raw.partitionKey,
        rowKey: raw.rowKey,
        category: raw.category || "",
        title: raw.title || "",
        description: raw.description || "",
        owner: raw.owner || "",
        savedAt: raw.savedAt || raw.timestamp || new Date().toISOString(),
        ...raw
    };
}

// ---------------------------------------------------
// MAIN ENTRY
// ---------------------------------------------------
module.exports = async function (context, req) {
    const method = req.method.toUpperCase();

    const cors = {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "GET,POST,DELETE,OPTIONS",
        "Access-Control-Allow-Headers": "*"
    };

    if (method === "OPTIONS") {
        context.res = { status: 204, headers: cors };
        return;
    }

    const url = new URL(req.url);
    const category = (url.searchParams.get("category") || req.body?.category || "").toLowerCase();
    const uid = url.searchParams.get("uid") || req.body?.uid || null;

    // ---------------------------------------------------
    // TABLE NAMES FROM AZURE APP SETTINGS
    // ---------------------------------------------------
    const SUGGESTIONS_TABLE = process.env.TABLES_TABLE_NAME_SUGGESTIONS || "Suggestions";
    const TROUBLE_TABLE = process.env.TABLES_TABLE_NAME_TROUBLESHOOTING || "Troubleshooting";

    // ---------------------------------------------------
    // SUGGESTIONS — GET ALL
    // ---------------------------------------------------
    if (method === "GET" && category === "suggestions") {
        const client = getTableClient(SUGGESTIONS_TABLE);
        const items = [];

        for await (const entity of client.listEntities()) {
            items.push(map(entity));
        }

        items.sort((a, b) => (a.savedAt < b.savedAt ? 1 : -1));

        context.res = { status: 200, headers: cors, body: items };
        return;
    }

    // ---------------------------------------------------
    // TROUBLESHOOTING — GET BY UID
    // ---------------------------------------------------
    if (method === "GET" && category === "troubleshooting") {
        if (!uid) {
            context.res = { status: 400, headers: cors, body: "Missing UID" };
            return;
        }

        const client = getTableClient(TROUBLE_TABLE);
        const filter = `PartitionKey eq 'UID_${uid}'`;

        const items = [];
        for await (const entity of client.listEntities({ queryOptions: { filter } })) {
            items.push(map(entity));
        }

        items.sort((a, b) => (a.rowKey < b.rowKey ? 1 : -1));

        context.res = { status: 200, headers: cors, body: items };
        return;
    }

    // ---------------------------------------------------
    // SUGGESTIONS — POST NEW
    // ---------------------------------------------------
    if (method === "POST" && category === "suggestions") {
        const { title, description, owner } = req.body;

        if (!title || !description) {
            context.res = { status: 400, headers: cors, body: "Missing title or description" };
            return;
        }

        const client = getTableClient(SUGGESTIONS_TABLE);

        const entity = {
            partitionKey: "Suggestions",
            rowKey: crypto.randomUUID(),
            category: "Suggestions",
            title,
            description,
            owner: owner || "Anonymous",
            savedAt: new Date().toISOString()
        };

        await client.upsertEntity(entity);

        context.res = { status: 200, headers: cors, body: { ok: true, entity } };
        return;
    }

    // ---------------------------------------------------
    // TROUBLESHOOTING — POST NEW
    // ---------------------------------------------------
    if (method === "POST" && category === "troubleshooting") {
        const { title, description, owner } = req.body;

        if (!uid) {
            context.res = { status: 400, headers: cors, body: "Missing UID" };
            return;
        }
        if (!description) {
            context.res = { status: 400, headers: cors, body: "Missing description" };
            return;
        }

        const client = getTableClient(TROUBLE_TABLE);

        const rowKey = new Date().toISOString();

        const entity = {
            partitionKey: `UID_${uid}`,
            rowKey,
            category: "Troubleshooting",
            title: title || "Troubleshooting Entry",
            description,
            owner: owner || "Unknown",
            savedAt: rowKey
        };

        await client.upsertEntity(entity);

        context.res = { status: 200, headers: cors, body: { ok: true, entity } };
        return;
    }

    // ---------------------------------------------------
    // DELETE — ANY CATEGORY
    // ---------------------------------------------------
    if (method === "DELETE") {
        const rowKey = url.searchParams.get("rowKey") || req.body?.rowKey;
        const partitionKey = url.searchParams.get("partitionKey") || req.body?.partitionKey;

        if (!partitionKey || !rowKey) {
            context.res = { status: 400, headers: cors, body: "Missing PK or RK" };
            return;
        }

        const tableName =
            category === "suggestions"
                ? SUGGESTIONS_TABLE
                : TROUBLE_TABLE;

        const client = getTableClient(tableName);

        await client.deleteEntity(partitionKey, rowKey);

        context.res = { status: 200, headers: cors, body: { ok: true } };
        return;
    }

    context.res = { status: 400, headers: cors, body: "Invalid request" };
};
