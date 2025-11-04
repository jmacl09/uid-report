import logging
import azure.functions as func
from azure.data.tables import TableClient
from azure.identity import DefaultAzureCredential
import json
import os

def main(req: func.HttpRequest) -> func.HttpResponse:
    logging.info("SaveProject function triggered.")

    # Load environment variables
    account_url = os.environ.get("TABLES_ACCOUNT_URL")
    table_name = os.environ.get("TABLES_TABLE_NAME")

    if not account_url or not table_name:
        logging.error("No environment configuration found.")
        return func.HttpResponse("Configuration missing.", status_code=500)

    # Authenticate with Managed Identity
    credential = DefaultAzureCredential()

    # ✅ Create a client the correct way
    table_client = TableClient(
        endpoint=f"{account_url}/{table_name}",
        table_name=table_name,
        credential=credential
    )

    try:
        req_body = req.get_json()
        user_id = req_body.get("userId")
        project_id = req_body.get("projectId")
        project_data = req_body.get("projectData", {})

        entity = {
            "PartitionKey": user_id,
            "RowKey": project_id,
            "ProjectData": json.dumps(project_data)
        }

        table_client.create_entity(entity=entity)

        return func.HttpResponse(
            json.dumps({"status": "success", "message": "Project saved successfully."}),
            mimetype="application/json"
        )

    except Exception as e:
        logging.exception("Error saving project.")
        return func.HttpResponse(
            json.dumps({"status": "error", "message": str(e)}),
            mimetype="application/json",
            status_code=500
        )
