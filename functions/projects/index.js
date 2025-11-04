const { TableClient } = require("@azure/data-tables");
const { DefaultAzureCredential } = require("@azure/identity");

module.exports = async function (context, req) {
  const TABLES_ACCOUNT_URL = process.env.TABLES_ACCOUNT_URL || "https://optical360.table.core.windows.net";
  const TABLES_TABLE_NAME = process.env.TABLES_TABLE_NAME || "Projects";
  const credential = new DefaultAzureCredential();

  try {
    const client = new TableClient(TABLES_ACCOUNT_URL, TABLES_TABLE_NAME, credential);

    // Ensure the request includes required fields
    if (!req.body || !req.body.uid || !req.body.category || !req.body.title) {
      context.res = {
        status: 400,
        body: { ok: false, message: "Missing required fields: uid, category, or title" }
      };
      return;
    }

    const { uid, category, title, description, owner } = req.body;

    // Construct entity
    const entity = {
      partitionKey: `UID_${uid}`,
      rowKey: new Date().toISOString(),
      category,
      title,
      description: description || "",
      owner: owner || "Unknown",
      timestamp: new Date().toISOString()
    };

    await client.createEntity(entity);

    context.res = {
      status: 200,
      body: {
        ok: true,
        message: `✅ Added ${category} entry for UID ${uid} successfully.`,
        entity
      }
    };
  } catch (err) {
    context.log.error("Error writing to Table Storage:", err.message);
    context.res = {
      status: 500,
      body: { ok: false, error: err.message }
    };
  }
};
