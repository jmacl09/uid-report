const { TableClient, DefaultAzureCredential } = require("@azure/data-tables");

const TABLES_ACCOUNT_URL = process.env.TABLES_ACCOUNT_URL || "";
const TABLES_TABLE_NAME = process.env.TABLES_TABLE_NAME || "Projects";

let tableClient = null;

function getTableClient() {
  if (tableClient) return tableClient;

  // Prefer managed identity when deployed, else fall back to local dev storage
  if (TABLES_ACCOUNT_URL) {
    const credential = new DefaultAzureCredential();
    tableClient = new TableClient(TABLES_ACCOUNT_URL, TABLES_TABLE_NAME, credential);
  } else if (process.env.TABLES_CONNECTION_STRING) {
    tableClient = TableClient.fromConnectionString(
      process.env.TABLES_CONNECTION_STRING,
      TABLES_TABLE_NAME
    );
  } else {
    throw new Error("No valid table credentials found (TABLES_ACCOUNT_URL or TABLES_CONNECTION_STRING)");
  }

  return tableClient;
}

async function ensureTable(client) {
  try {
    await client.createTable();
  } catch (err) {
    if (!/TableAlreadyExists/i.test(err.message)) throw err;
  }
}

function getUserPrincipal(req) {
  try {
    const header = req.headers && (req.headers["x-ms-client-principal"] || req.headers["X-MS-CLIENT-PRINCIPAL"]);
    if (!header) return { email: "anonymous@example.com", alias: "anonymous" };
    const decoded = JSON.parse(Buffer.from(header, "base64").toString("utf8"));
    const email = String(decoded?.userDetails || "anonymous@example.com").toLowerCase();
    const alias = email.includes("@") ? email.split("@")[0] : email;
    return { email, alias };
  } catch {
    return { email: "anonymous@example.com", alias: "anonymous" };
  }
}

module.exports = async function (context, req) {
  const method = (req.method || "GET").toUpperCase();
  const { email: userEmail, alias: userAlias } = getUserPrincipal(req);

  const client = getTableClient();
  await ensureTable(client);

  const json = (status, body) => {
    context.res = { status, headers: { "Content-Type": "application/json" }, body };
  };

  try {
    if (method === "GET") {
      const rows = [];
      for await (const e of client.listEntities({
        queryOptions: { filter: `PartitionKey eq '${userEmail}'` },
      })) {
        rows.push(e);
      }
      return json(200, rows);
    }

    if (method === "POST") {
      const body = req.body || {};
      const id = body.id || `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
      const entity = {
        partitionKey: userEmail,
        rowKey: id,
        Name: body.name || "Untitled",
        CreatedAt: new Date().toISOString(),
        CreatedByEmail: userEmail,
        CreatedByAlias: userAlias,
        Data: JSON.stringify(body.data || {}),
        SourceUids: JSON.stringify(body.sourceUids || []),
      };
      await client.createEntity(entity);
      return json(201, entity);
    }

    if (method === "PUT") {
      const id =
        context.bindingData?.id || req.query?.id || req.body?.id;
      if (!id) return json(400, { error: "Missing id" });
      const existing = await client.getEntity(userEmail, id).catch(() => null);
      if (!existing) return json(404, { error: "Not found" });
      if (existing.CreatedByEmail && existing.CreatedByEmail !== userEmail)
        return json(403, { error: "Forbidden" });

      const body = req.body || {};
      const updated = {
        ...existing,
        Name: body.name ?? existing.Name,
        Data: body.data ? JSON.stringify(body.data) : existing.Data,
        SourceUids: body.sourceUids
          ? JSON.stringify(body.sourceUids)
          : existing.SourceUids,
      };
      await client.updateEntity(updated, "Replace");
      return json(200, updated);
    }

    if (method === "DELETE") {
      const id = context.bindingData?.id || req.query?.id;
      if (!id) return json(400, { error: "Missing id" });
      const existing = await client.getEntity(userEmail, id).catch(() => null);
      if (!existing) return json(404, { error: "Not found" });
      if (existing.CreatedByEmail && existing.CreatedByEmail !== userEmail)
        return json(403, { error: "Forbidden" });
      await client.deleteEntity(userEmail, id);
      return json(204, null);
    }

    return json(405, { error: "Unsupported method" });
  } catch (err) {
    context.log.error(err);
    return json(500, { error: String(err.message || err) });
  }
};
