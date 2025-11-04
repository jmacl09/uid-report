const { TableClient } = require("@azure/data-tables");
const { ManagedIdentityCredential } = require("@azure/identity");

module.exports = async function (context, req) {
    context.log("🔍 Starting testConnection...");

    const tableAccountUrl =
        process.env.AZURE_TABLE_SERVICE_URI || process.env.TABLES_ACCOUNT_URL || "https://optical360.table.core.windows.net";
    const tableName = process.env.TABLES_TABLE_NAME || "Projects";

    try {
        // Initialize credential (Managed Identity)
        const credential = new ManagedIdentityCredential(process.env.AZURE_CLIENT_ID);

        // Test token acquisition
        const token = await credential.getToken("https://storage.azure.com/");
        context.log(`✅ MSI token retrieved: ${token ? "Yes" : "No"}`);

        // Connect to Table Storage
        const tableClient = new TableClient(tableAccountUrl, tableName, credential);
        const testEntity = {
            partitionKey: "test",
            rowKey: `${Date.now()}`,
            message: "MSI test worked!"
        };

        await tableClient.createEntity(testEntity);
        context.log("✅ Successfully inserted test entity into Table Storage");

        context.res = {
            status: 200,
            body: {
                ok: true,
                message: "✅ Connected via Managed Identity and inserted test record",
                account: tableAccountUrl,
                table: tableName
            }
        };
    } catch (err) {
        context.log.error("❌ ERROR:", err.message);
        context.res = {
            status: 500,
            body: {
                ok: false,
                message: err.message,
                stack: err.stack
            }
        };
    }
};
