const { TableClient, odata } = require("@azure/data-tables");

const tableName = process.env.TABLES_TABLE_NAME;
const connStr = process.env.TABLES_CONNECTION_STRING;
const tableClient = TableClient.fromConnectionString(connStr, tableName);

function getUserEmail(req) {
  try {
    const header = req.headers.get("x-ms-client-principal");
    if (!header) return "anonymous@example.com";
    const decoded = JSON.parse(Buffer.from(header, "base64").toString("utf8"));
    return decoded?.userDetails?.toLowerCase() || "anonymous@example.com";
  } catch {
    return "anonymous@example.com";
  }
}

async function handler(request, context) {
  const method = request.method.toUpperCase();
  const userEmail = getUserEmail(request);

  try {
    if (method === "GET") {
      const entities = [];
      for await (const e of tableClient.listEntities({
        queryOptions: { filter: odata`PartitionKey eq ${userEmail}` },
      })) entities.push(e);
      return { status: 200, jsonBody: entities };
    }

    if (method === "POST") {
      const body = await request.json();
      const id = body.id || Date.now().toString();
      const entity = {
        partitionKey: userEmail,
        rowKey: id,
        Name: body.name || "Untitled",
        CreatedAt: new Date().toISOString(),
        CreatedByEmail: userEmail,
        Data: JSON.stringify(body.data || {}),
      };
      await tableClient.createEntity(entity);
      return { status: 201, jsonBody: entity };
    }

    if (method === "PUT") {
      const body = await request.json();
      const id = body.id;
      if (!id) throw new Error("Missing id");
      const entity = await tableClient.getEntity(userEmail, id);
      const updated = { ...entity, ...body };
      await tableClient.updateEntity(updated, "Replace");
      return { status: 200, jsonBody: updated };
    }

    if (method === "DELETE") {
      const url = new URL(request.url);
      const id = url.searchParams.get("id");
      if (!id) throw new Error("Missing id");
      await tableClient.deleteEntity(userEmail, id);
      return { status: 204 };
    }

    return { status: 405, body: "Unsupported method" };
  } catch (err) {
    context.log.error(err);
    return { status: 500, body: `Error: ${err.message}` };
  }
}

module.exports = handler;  // Ensure this line is at the end of your file to export the handler
