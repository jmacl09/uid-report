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
   - Python 3.10+
   - Node 18+ (only if using JS functions)
2. Install deps
   - Python (from repo root or functions folder)
     ```powershell
     # Windows PowerShell
     python -m venv .venv
     .venv\Scripts\Activate.ps1
     pip install -r functions/requirements.txt
     ```
   - Node (if you use JS functions)
     ```powershell
     npm install --prefix functions
     ```
3. Start
   ```powershell
   func start
   ```
   Host will expose:
   - POST `/api/vso2` — Kusto-backed endpoint for VSO2 page (Stage: VSO_Details | Email_Template)
   - existing functions as configured

## Required settings (Managed Identity only — no connection strings)

Use the Function App's System-assigned Managed Identity (MI). Grant it access to your storage account and configure only the endpoint URL:

- `TABLES_ACCOUNT_URL` — `https://<storage-account-name>.table.core.windows.net` (e.g., `https://optical360.table.core.windows.net`)
- `TABLES_TABLE_NAME` — optional, defaults to `projects`

Also ensure the Function App has:

- System-assigned identity: Enabled (Function App → Identity → System assigned → On)
- RBAC role on the storage account: Assign the MI the role `Storage Table Data Contributor` at the scope of the storage account.

Notes:

- Do not configure table connection strings. Authentication uses Azure AD via Managed Identity.
- For local dev, `DefaultAzureCredential` will use your Azure CLI or VS Code login. Run `az login` and keep `TABLES_ACCOUNT_URL` set locally in `local.settings.json`.

### Kusto (Azure Data Explorer) permissions

The Function App's System-assigned Managed Identity must be granted access to Kusto cluster `waneng.westus2.kusto.windows.net` database `waneng`.

- Object ID: `0cb74659-d1ff-452b-bf79-40ecfb321a67`
- Assign at least `Viewer`/`User` role and table query permissions for `DarkFiberTracker` and `LinkMetadata`.
- For local development, your signed-in user must also have query permissions.

## Deploy

Option A (CLI):

1. Create a Function App (Consumption Plan, Node 18, Functions v4) and a Storage account.
2. Enable the Function App's System-assigned Identity and assign `Storage Table Data Contributor` on your storage account.
3. In the Function App → Configuration, set `TABLES_ACCOUNT_URL` and (optionally) `TABLES_TABLE_NAME`.
4. From this folder:
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
