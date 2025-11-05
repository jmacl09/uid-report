const { app } = require("@azure/functions");
const { TableClient, TableServiceClient } = require("@azure/data-tables");
let DefaultAzureCredential = null;
try {
  ({ DefaultAzureCredential } = require("@azure/identity"));
} catch {
  /* optional */
}

// === Helper to build a TableClient ===
function getTableClient(tableName) {
  const conn = process.env.TABLES_CONNECTION_STRING || process.env.AzureWebJobsStorage || "";
  const accountUrl = process.env.TABLES_ACCOUNT_URL || "";

  if (accountUrl && DefaultAzureCredential) {
    const cred = new DefaultAzureCredential();
    const client = new TableClient(accountUrl, tableName, cred);
    const svc = TableServiceClient.fromUrl(accountUrl, cred);
    return {
      client,
      ensureTable: async () => {
        try {
          await svc.createTable(tableName);
        } catch (e) {
          if (!(e && e.statusCode === 409)) throw e;
        }
      },
    };
  }

  if (!conn)
    throw new Error(
      "Missing Azure Table configuration. Provide TABLES_CONNECTION_STRING, TABLES_ACCOUNT_URL, or AzureWebJobsStorage."
    );

  return {
    client: TableClient.fromConnectionString(conn, tableName),
    ensureTable: async () => {
      try {
        const svc = TableServiceClient.fromConnectionString(conn);
        await svc.createTable(tableName);
      } catch (e) {
        if (!(e && e.statusCode === 409)) throw e;
      }
    },
  };
}

// === MAIN FUNCTION ===
app.http("HttpTrigger1", {
  methods: ["GET", "POST", "OPTIONS"],
  authLevel: "anonymous",

  handler: async (request, context) => {
    context.log(`HttpTrigger1 called: ${request.method} ${request.url}`);

    const corsHeaders = {
      "Access-Control-Allow-Origin": "https://optical360.net",
      "Access-Control-Allow-Methods": "GET,POST,OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type, Authorization",
      "Access-Control-Allow-Credentials": "true",
    };

    // Handle preflight
    if (request.method === "OPTIONS") {
      return { status: 204, headers: corsHeaders };
    }

    if (request.method === "GET") {
      return {
        status: 200,
        headers: corsHeaders,
        jsonBody: { ok: true, message: "Function is alive. Use POST to save data." },
      };
    }

    // === Handle POST ===
    let payload;
    try {
      payload = await request.json();
    } catch {
      const txt = await request.text().catch(() => "");
      return {
        status: 400,
        headers: corsHeaders,
        jsonBody: { ok: false, error: "Invalid JSON body", received: txt?.slice(0, 200) },
      };
    }

    const { category, uid, title, description, owner, timestamp } = payload || {};
    if (!uid || !category || !title) {
      return {
        status: 400,
        headers: corsHeaders,
        jsonBody: { ok: false, error: "Missing required fields: uid, category, title" },
      };
    }

    // === Choose table based on category ===
    let tableName;
    switch (category) {
      case "Notes":
      case "Comments":
        tableName = "Notes";
        break;
      default:
        tableName = "Projects";
    }

    context.log(`Saving ${category} for UID ${uid} into table ${tableName}`);

    try {
      const { client, ensureTable } = getTableClient(tableName);
      await ensureTable();

      const nowIso =
        timestamp && !isNaN(new Date(timestamp))
          ? new Date(timestamp).toISOString()
          : new Date().toISOString();

      const entity = {
        partitionKey: `UID_${uid}`,
        rowKey: nowIso, // allows multiple notes per UID
        category,
        title,
        description: description || "",
        owner: owner || "Unknown",
        savedAt: nowIso,
      };

      await client.upsertEntity(entity, "Merge");

      return {
        status: 200,
        headers: corsHeaders,
        jsonBody: { ok: true, message: `Saved ${category} for UID ${uid}`, entity },
      };
    } catch (err) {
      console.error("Save failed:", err);
      return { status: 500, headers: corsHeaders, jsonBody: { ok: false, error: String(err?.message || err) } };
    }
  },
});
