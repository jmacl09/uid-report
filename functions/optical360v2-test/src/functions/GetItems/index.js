const { app } = require("@azure/functions");
const { TableClient } = require("@azure/data-tables");
let DefaultAzureCredential = null;
try {
  ({ DefaultAzureCredential } = require("@azure/identity"));
} catch {
  /* optional */
}

function getTableClient(tableName) {
  const conn = process.env.TABLES_CONNECTION_STRING || process.env.AzureWebJobsStorage || "";
  const accountUrl = process.env.TABLES_ACCOUNT_URL || "";
  if (accountUrl && DefaultAzureCredential) {
    const cred = new DefaultAzureCredential();
    return new TableClient(accountUrl, tableName, cred);
  }
  return TableClient.fromConnectionString(conn, tableName);
}

app.http("GetItems", {
  methods: ["GET"],
  authLevel: "anonymous",
  handler: async (req, context) => {
    const uid = req.query.get("uid");
    const category = req.query.get("category");

    if (!uid || !category) {
      return { status: 400, jsonBody: { ok: false, error: "Missing uid or category" } };
    }

    const tableName = category === "Notes" || category === "Comments" ? "Notes" : "Projects";
    const client = getTableClient(tableName);

    const entities = [];
    for await (const entity of client.listEntities({
      queryOptions: { filter: `PartitionKey eq 'UID_${uid}'` },
    })) {
      entities.push(entity);
    }

    return { status: 200, jsonBody: { ok: true, uid, category, items: entities } };
  },
});
