# fibertools-api (Azure Functions)

This folder contains the Azure Functions API for Projects.

## Structure

- `projects/` — HTTP function with route `/api/projects/{id?}`
- `host.json` — Functions host config
- `package.json` — Node project (use Node 18+, Functions v4)
- `.funcignore` — files excluded from publish (local.settings.json, root artifacts)

## Local development

1. Install tools
   - Azure Functions Core Tools v4
   - Node 18+
2. Install deps
   ```sh
   npm install
   ```
3. Start
   ```sh
   npm start
   ```
   Host will expose GET/POST/PUT/DELETE at `/api/projects`.

## Required settings

These are app settings in Azure (Function App → Configuration → Application settings):

- `AzureWebJobsStorage` — Storage connection string (not `UseDevelopmentStorage=true`)
- `TABLES_CONNECTION_STRING` — Same storage account connection string (for Table Storage)
- `TABLES_TABLE_NAME` — `ProjectSnapshots` (or your choice)

> local.settings.json is for local use only and is not deployed.

## Deploy

Option A (CLI):

1. Create a Function App (Consumption Plan, Node 18, Functions v4) and a Storage account.
2. In the Function App → Configuration, set the settings above.
3. From this folder:
   ```sh
   func azure functionapp publish <YOUR_FUNCTION_APP_NAME>
   ```

Option B (VS Code):

- Use the Azure Functions extension → Deploy to Function App → select this folder.

## Auth (recommended)

Enable Azure AD authentication so ownership checks work:
- Function App → Authentication → Add identity provider (Microsoft)
- Action to take when request is not authenticated: `Log in with Azure Active Directory`
- Add your front-end origin(s) to CORS (Function App → CORS)

When protected by EasyAuth, the function reads the user from `X-MS-CLIENT-PRINCIPAL`.
