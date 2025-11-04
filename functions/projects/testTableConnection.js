// testTableConnection.js
import { TableClient } from "@azure/data-tables";
import { DefaultAzureCredential } from "@azure/identity";

const accountUrl = "https://optical360.table.core.windows.net";
const tableName = "Projects";
const credential = new DefaultAzureCredential();

const client = new TableClient(accountUrl, tableName, credential);

(async () => {
  try {
    console.log("🔄 Testing connection to Azure Table Storage...");
    const entities = client.listEntities();
    let count = 0;
    for await (const e of entities) {
      console.log("Entity:", e);
      if (++count >= 3) break;
    }
    console.log(`✅ Connected successfully! Retrieved ${count} entities.`);
  } catch (err) {
    console.error("❌ Connection failed:", err.message);
  }
})();
