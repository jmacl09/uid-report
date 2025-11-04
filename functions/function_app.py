import azure.functions as func
import json
import logging
from typing import Any, Dict, List

from azure.identity import (
	DefaultAzureCredential,
	AzureCliCredential,
	VisualStudioCodeCredential,
	InteractiveBrowserCredential,
)
from azure.kusto.data import KustoClient, KustoConnectionStringBuilder, ClientRequestProperties


app = func.FunctionApp()


KUSTO_CLUSTER = "https://waneng.westus2.kusto.windows.net"
KUSTO_DB = "waneng"


def _get_kusto_client() -> KustoClient:
	"""
	Build a Kusto client using the best available local auth method.
	Order of preference for local dev:
	- Azure CLI login (az login) via SDK helper when available
	- azure-identity token provider (CLI/VS Code/Browser) via token provider hook
	- Device code auth as a last resort

	This also works in production when DefaultAzureCredential can acquire MSI.
	"""

	# Helper: obtain an AAD token using developer-friendly credentials
	def _dev_token(scope: str) -> str:
		for cred in (
			# Try Azure CLI first so we use the signed-in Entra ID (az login)
			AzureCliCredential(),
			# VS Code signed-in identity if available
			VisualStudioCodeCredential(),
			# Interactive browser (last resort for local)
			InteractiveBrowserCredential(),
		):
			try:
				t = cred.get_token(scope)
				return t.token
			except Exception:  # pragma: no cover - best-effort chain
				continue
		raise Exception("No developer credential available. Run 'az login' or sign into VS Code.")

	# Scope for Azure Data Explorer (Kusto)
	scope = "https://kusto.kusto.windows.net/.default"

	# Try SDK helpers in descending order of compatibility
	try:
		if hasattr(KustoConnectionStringBuilder, "with_azure_cli_authentication"):
			# Uses Azure CLI cached token for the signed-in user
			kcsb = KustoConnectionStringBuilder.with_azure_cli_authentication(KUSTO_CLUSTER)
			logging.info("Kusto auth: with_azure_cli_authentication")
			return KustoClient(kcsb)
	except Exception as e:
		logging.warning("Azure CLI auth path failed: %s", e)

	# Fallback: provide a token provider hook (covers many SDK versions)
	def token_provider() -> str:
		# Prefer MSI in real hosting; locally prefer developer identity
		try:
			# Try MSI/Default chain first (works in cloud; may fail locally)
			mi = DefaultAzureCredential()
			return mi.get_token(scope).token
		except Exception:
			return _dev_token(scope)

	# Try both token provider wiring options based on SDK capabilities
	if hasattr(KustoConnectionStringBuilder, "with_aad_token_provider"):
		kcsb = KustoConnectionStringBuilder.with_aad_token_provider(KUSTO_CLUSTER, token_provider)
		logging.info("Kusto auth: with_aad_token_provider (token hook)")
		return KustoClient(kcsb)

	if hasattr(KustoConnectionStringBuilder, "__init__") and hasattr(KustoConnectionStringBuilder, "set_token_provider"):
		kcsb = KustoConnectionStringBuilder(KUSTO_CLUSTER)
		kcsb.set_token_provider(token_provider)  # type: ignore[attr-defined]
		logging.info("Kusto auth: set_token_provider (token hook)")
		return KustoClient(kcsb)

	# Last resort: device authentication (prints device code)
	if hasattr(KustoConnectionStringBuilder, "with_aad_device_authentication"):
		kcsb = KustoConnectionStringBuilder.with_aad_device_authentication(KUSTO_CLUSTER)
		logging.info("Kusto auth: with_aad_device_authentication (device code)")
		return KustoClient(kcsb)

	raise RuntimeError("No compatible Kusto auth method found in installed azure-kusto-data SDK.")


def _build_crp(params: Dict[str, Any]) -> ClientRequestProperties:
	crp = ClientRequestProperties()
	for k, v in params.items():
		crp.set_parameter(k, v if v is not None else "")
	return crp


# Shared query prelude with parameters
QUERY_HEADER = (
	"declare query_parameters("
	"FacilityCodeA:string, DiversityParam:string, SpliceRackParam:string);\n"
)


FACILITY_ONLY_QUERY = QUERY_HEADER + r'''
DarkFiberTracker
| extend Fc = trim(" ", FacilityCodeA)
| where Fc == FacilityCodeA
| extend
	SpanID       = SpanID,
	Diversity    = Diversity,
	RawIDF       = tostring(IDF_A),
	RawSplice    = SpliceRackA,
	RawRackUnit  = tostring(SpliceRackUnitA),
	WiringScope  = toupper(WiringScope)
| extend
	SpliceRackA = case(
		isempty(RawSplice) or isempty(RawRackUnit), "UNKNOWN",
		strcat(RawSplice, " U", RawRackUnit)
	)
| join kind=leftouter (
	LinkMetadata
	| extend SpanID = tostring(SolutionId)
	| summarize LinkLifecycleState = any(LinkLifecycleState) by SpanID
) on SpanID
| where isnotempty(LinkLifecycleState)
| extend
	Status = LinkLifecycleState,
	IDF_A  = replace("-", "", case(
		RawIDF contains "IDF",                           
		RawIDF,
		RawIDF matches regex @"^[1-4]$",                 
		strcat("IDF", RawIDF),
		RawIDF == "" and WiringScope == "RNG-RNG",       
		"RNG",
		RawIDF == "" and WiringScope startswith "EDGE-", 
		"EDGE",
		RawIDF
	))
| extend
	Color = case(
		Status has "InProduction",  "Good",
		Status has "InMaintenance", "Warning",
		Status has "Final",         "Attention",
									 "Accent"
	),
	OpticalLink = strcat(
		"https://phynet.trafficmanager.net/Optical/OpticalLinkMonitor?",
		"spanId=", SpanID,
		"&timespan=7d&multiSpanMode=false"
	)
| extend
	FormattedSpans = tostring(
		strcat_array(
			toscalar(
				DarkFiberTracker
				| extend Fc2 = trim(" ", FacilityCodeA)
				| where Fc2 == FacilityCodeA
				| join kind=leftouter (
					LinkMetadata
					| extend SpanID = tostring(SolutionId)
					| summarize LinkLifecycleState = any(LinkLifecycleState) by SpanID
				  ) on SpanID
				| where isnotempty(LinkLifecycleState)
				| distinct SpanID
				| summarize FS = make_list(toupper(SpanID))
				| project FormattedSpans = FS
			),
			","
		)
	)
| project
	SpanID,
	Diversity,
	IDF_A,
	SpliceRackA,
	WiringScope,
	Status,
	Color,
	OpticalLink,
	FormattedSpans
| order by Diversity asc
| as SpanSummary
| extend
	prodSpans  = toscalar(SpanSummary | summarize countif(Status has "InProduction")),
	maintSpans = toscalar(SpanSummary | summarize countif(Status has "InMaintenance"))
| extend
	prodPct = iif(
		(prodSpans + maintSpans) == 0,
		0,
		toint(100.0 * prodSpans / (prodSpans + maintSpans))
	)
'''


DIVERSITY_SPLICE_QUERY = QUERY_HEADER + r'''
DarkFiberTracker
| extend Fc = trim(" ", FacilityCodeA)
| where Fc == FacilityCodeA
| extend
	SpanID       = SpanID,
	Diversity    = Diversity,
	RawIDF       = tostring(IDF_A),
	RawSplice    = SpliceRackA,
	RawRackUnit  = tostring(SpliceRackUnitA),
	WiringScope  = toupper(WiringScope),
	State        = State
| join kind=leftouter (
	LinkMetadata
	| where not(StartDevice contains "omt")
	| extend SpanID = tostring(SolutionId)
	| summarize LinkLifecycleState = any(LinkLifecycleState) by SpanID
) on SpanID
| where isnotempty(LinkLifecycleState) or State == "New"
| extend
	SpliceRackA = case(
	  isempty(RawSplice) or isempty(RawRackUnit), "UNKNOWN",
	  strcat(RawSplice, " U", RawRackUnit)
	),
	IDF_A = replace("-", "", case(
	  RawIDF contains "IDF",                           RawIDF,
	  RawIDF matches regex @"^[1-4]$",                 strcat("IDF", RawIDF),
	  RawIDF == "" and WiringScope == "RNG-RNG",       "RNG",
	  RawIDF == "" and WiringScope startswith "EDGE-", "EDGE",
	  RawIDF
	)),
	Status = iif(State == "New", "New", LinkLifecycleState)
| where 
	tolower(Diversity) contains tolower(DiversityParam)
  and tolower(SpliceRackA) contains tolower(SpliceRackParam)
| extend
	Color = case(
	  Status == "InProduction",  "Good",
	  Status == "InMaintenance", "Warning",
	  Status == "New",           "Accent",
								  "Attention"
	),
	OpticalLink = strcat(
	  "https://phynet.trafficmanager.net/Optical/OpticalLinkMonitor/LoadSpanDetails?",
	  "spanId=", SpanID,
	  "&timespan=7d&multiSpanMode=false"
	),
	FormattedSpans = toscalar(
	  DarkFiberTracker
	  | extend Fc2 = trim(" ", FacilityCodeA)
	  | where Fc2 == FacilityCodeA
	  | join kind=leftouter (
		  LinkMetadata
		  | where not(StartDevice contains "omt")
		  | extend SpanID = tostring(SolutionId)
		  | summarize LinkLifecycleState = any(LinkLifecycleState) by SpanID
		) on SpanID
	  | extend
		  State2      = State,
		  RawSplice2  = SpliceRackA,
		  RawRackUnit2= tostring(SpliceRackUnitA)
	  | where isnotempty(LinkLifecycleState) or State2 == "New"
	  | extend
		  SpliceRackB = case(
			isempty(RawSplice2) or isempty(RawRackUnit2), "UNKNOWN",
			strcat(RawSplice2, " U", RawRackUnit2)
		  ),
		  Diversity2 = Diversity
	  | where 
		  tolower(Diversity2) contains tolower(DiversityParam)
		and tolower(SpliceRackB) contains tolower(SpliceRackParam)
	  | distinct SpanID
	  | extend Quoted = strcat('"', toupper(SpanID), '"')
	  | summarize FS = strcat_array(make_list(Quoted), ", ")
	)
| project
	SpanID,
	Diversity,
	IDF_A,
	SpliceRackA,
	WiringScope,
	Status,
	Color,
	OpticalLink,
	FormattedSpans
| order by Diversity asc
| as SpanSummary
| extend
	prodSpans  = toscalar(SpanSummary | summarize countif(Status == "InProduction")),
	maintSpans = toscalar(SpanSummary | summarize countif(Status == "InMaintenance"))
| extend
	prodPct = toint(100.0 * prodSpans / (prodSpans + maintSpans))
'''


DC_AND_DIVERSITY_QUERY = QUERY_HEADER + r'''
DarkFiberTracker
| extend Fc = trim(" ", FacilityCodeA)
| where Fc == FacilityCodeA
| extend
	SpanID       = SpanID,
	Diversity    = Diversity,
	RawIDF       = tostring(IDF_A),
	RawSplice    = SpliceRackA,
	RawRackUnit  = tostring(SpliceRackUnitA),
	WiringScope  = toupper(WiringScope),
	State        = State
| join kind=leftouter (
	LinkMetadata
	| where not(StartDevice contains "omt")
	| extend SpanID = tostring(SolutionId)
	| summarize LinkLifecycleState = any(LinkLifecycleState) by SpanID
) on SpanID
| where isnotempty(LinkMetadata.LinkLifecycleState) or State == "New"
| extend
	IDF_A = replace("-", "", case(
	  RawIDF contains "IDF",                           RawIDF,
	  RawIDF matches regex @"^[1-4]$",                 strcat("IDF", RawIDF),
	  RawIDF == "" and WiringScope == "RNG-RNG",       "RNG",
	  RawIDF == "" and WiringScope startswith "EDGE-", "EDGE",
	  RawIDF
	)),
	SpliceRackA = case(
	  isempty(RawSplice) or isempty(RawRackUnit), "UNKNOWN",
	  strcat(RawSplice, " U", RawRackUnit)
	),
	Status = iif(State == "New", "New", LinkLifecycleState)
| where tolower(Diversity) contains tolower(DiversityParam)
| extend
	Color = case(
	  Status has "InProduction",  "Good",
	  Status has "InMaintenance", "Warning",
	  Status has "Final",         "Attention",
	  Status == "New",            "Accent",
								   "Accent"
	),
	OpticalLink = strcat(
	  "https://phynet.trafficmanager.net/Optical/OpticalLinkMonitor/LoadSpanDetails?",
	  "spanId=", SpanID,
	  "&timespan=7d&multiSpanMode=false"
	)
| extend
	FormattedSpans = toscalar(
	  DarkFiberTracker
	  | extend Fc2 = trim(" ", FacilityCodeA)
	  | where Fc2 == FacilityCodeA
	  | join kind=leftouter (
		  LinkMetadata
		  | where not(StartDevice contains "omt")
		  | extend SpanID = tostring(SolutionId)
		  | summarize LinkLifecycleState = any(LinkLifecycleState) by SpanID
		) on SpanID
	  | where isnotempty(LinkLifecycleState) or State == "New"
	  | distinct SpanID
	  | extend Quoted = strcat('"', toupper(SpanID), '"')
	  | summarize FS = strcat_array(make_list(Quoted), ", ")
	  | project FS
	)
| project
	SpanID,
	Diversity,
	IDF_A,
	SpliceRackA,
	WiringScope,
	Status,
	Color,
	OpticalLink,
	FormattedSpans
| order by Diversity asc
| as SpanSummary
| extend
	prodSpans  = toscalar(SpanSummary | summarize countif(Status == "InProduction")),
	maintSpans = toscalar(SpanSummary | summarize countif(Status == "InMaintenance"))
| extend
	prodPct = iif(
	  (prodSpans + maintSpans) == 0,
	  0,
	  toint(100.0 * prodSpans / (prodSpans + maintSpans))
	)
'''


FACILITY_SPLICE_QUERY = QUERY_HEADER + r'''
DarkFiberTracker
| extend Fc = trim(" ", FacilityCodeA)
| where Fc == FacilityCodeA
| extend
	SpanID       = SpanID,
	Diversity    = Diversity,
	RawIDF       = tostring(IDF_A),
	RawSplice    = SpliceRackA,
	RawRackUnit  = tostring(SpliceRackUnitA),
	WiringScope  = toupper(WiringScope),
	State        = State
| join kind=leftouter (
	LinkMetadata
	| where not(StartDevice contains "omt")
	| extend SpanID = tostring(SolutionId)
	| summarize LinkLifecycleState = any(LinkLifecycleState) by SpanID
) on SpanID
| where isnotempty(LinkLifecycleState) or State == "New"
| extend
	SpliceRackA = case(
	  isempty(RawSplice) or isempty(RawRackUnit), "UNKNOWN",
	  strcat(RawSplice, " U", RawRackUnit)
	),
	IDF_A = replace("-", "", case(
	  RawIDF contains "IDF",                           RawIDF,
	  RawIDF matches regex @"^[1-4]$",                 strcat("IDF", RawIDF),
	  RawIDF == "" and WiringScope == "RNG-RNG",       "RNG",
	  RawIDF == "" and WiringScope startswith "EDGE-", "EDGE",
	  RawIDF
	)),
	Status = iif(State == "New", "New", LinkLifecycleState)
| where tolower(SpliceRackA) contains tolower(SpliceRackParam)
| extend
	Color = case(
	  Status has "InProduction",  "Good",
	  Status has "InMaintenance", "Warning",
	  Status has "Final",         "Attention",
	  Status == "New",            "Accent",
								   "Accent"
	),
	OpticalLink = strcat(
	  "https://phynet.trafficmanager.net/Optical/OpticalLinkMonitor/LoadSpanDetails?",
	  "spanId=", SpanID,
	  "&timespan=7d&multiSpanMode=false"
	),
	FormattedSpans = toscalar(
	  DarkFiberTracker
	  | extend Fc2 = trim(" ", FacilityCodeA)
	  | where Fc2 == FacilityCodeA
	  | join kind=leftouter (
		  LinkMetadata
		  | where not(StartDevice contains "omt")
		  | extend SpanID = tostring(SolutionId)
		  | summarize LinkLifecycleState = any(LinkLifecycleState) by SpanID
		) on SpanID
	  | extend
		  RawSplice2   = SpliceRackA,
		  RawRackUnit2 = tostring(SpliceRackUnitA),
		  State2       = State
	  | where isnotempty(LinkLifecycleState) or State2 == "New"
	  | extend
		  SpliceRackB = case(
			isempty(RawSplice2) or isempty(RawRackUnit2), "UNKNOWN",
			strcat(RawSplice2, " U", RawRackUnit2)
		  )
	  | where tolower(SpliceRackB) contains tolower(SpliceRackParam)
	  | distinct SpanID
	  | extend Quoted = strcat('"', toupper(SpanID), '"')
	  | summarize FS = strcat_array(make_list(Quoted), ", ")
	),
	Scope = case(
	  tolower(WiringScope) == "rngdc", "RNG-DC",
	  WiringScope
	)
| project
	SpanID,
	Diversity,
	IDF_A,
	SpliceRackA,
	WiringScope,
	Status,
	Color,
	OpticalLink,
	FormattedSpans
| order by Diversity asc
| as SpanSummary
| extend
	prodSpans  = toscalar(SpanSummary | summarize countif(Status == "InProduction")),
	maintSpans = toscalar(SpanSummary | summarize countif(Status == "InMaintenance"))
| extend
	prodPct = iif(
	  (prodSpans + maintSpans) == 0,
	  0,
	  toint(100.0 * prodSpans / (prodSpans + maintSpans))
	)
'''


def _execute_query(client: KustoClient, query: str, params: Dict[str, Any]) -> List[Dict[str, Any]]:
	crp = _build_crp(params)
	result = client.execute(KUSTO_DB, query, properties=crp)
	tables = result.primary_results
	if not tables:
		return []
	table = tables[0]
	columns = [c.column_name for c in table.columns]
	rows: List[Dict[str, Any]] = []
	for r in table.rows:
		# Some SDK versions return Row objects; handle both list/tuple and Row
		try:
			values = list(r)
		except TypeError:
			values = [r.get(c) for c in columns]
		row = {columns[i]: values[i] for i in range(len(columns))}
		rows.append(row)
	return rows


@app.route(route="vso2", methods=["GET", "POST", "OPTIONS"], auth_level=func.AuthLevel.ANONYMOUS)
def vso2_handler(req: func.HttpRequest) -> func.HttpResponse:
	# Simple GET for health/info so clicking the console link works
	if req.method == "GET":
		body = {
			"message": "VSO2 API is running.",
			"usage": {
				"POST /api/vso2": {
					"required": ["Stage"],
					"optional": ["FacilityCodeA", "Diversity", "SpliceRackA"],
					"notes": "Use Stage='VSO_Details' and provide FacilityCodeA. Use 'N' for Diversity/SpliceRackA to mean no filter.",
				}
			},
		}
		return func.HttpResponse(
			json.dumps(body),
			status_code=200,
			mimetype="application/json",
			headers={
				"Access-Control-Allow-Origin": "*",
				"Access-Control-Allow-Methods": "GET, POST, OPTIONS",
				"Access-Control-Allow-Headers": "Content-Type",
			},
		)

	# Handle CORS preflight quickly
	if req.method == "OPTIONS":
		return func.HttpResponse(
			status_code=200,
			headers={
				"Access-Control-Allow-Origin": "*",
				"Access-Control-Allow-Methods": "GET, POST, OPTIONS",
				"Access-Control-Allow-Headers": "Content-Type",
			},
		)

	try:
		body = req.get_json()
	except ValueError:
		return func.HttpResponse(
			json.dumps({"error": "Invalid JSON"}),
			status_code=400,
			mimetype="application/json",
		)

	stage = (body or {}).get("Stage")
	if not stage:
		return func.HttpResponse(
			json.dumps({"error": "Missing required field: Stage"}),
			status_code=400,
			mimetype="application/json",
		)

	# Normalize inputs
	facility = (body.get("FacilityCodeA") or "").strip()
	diversity = body.get("Diversity") or ""
	splice_rack = body.get("SpliceRackA") or ""

	if stage == "VSO_Details":
		# Logic App used sentinel 'N' to mean no filter
		diversity_is_none = diversity == "N"
		splice_is_none = splice_rack == "N"

		client = _get_kusto_client()

		params = {
			"FacilityCodeA": facility,
			"DiversityParam": diversity,
			"SpliceRackParam": splice_rack,
		}

		if diversity_is_none and splice_is_none:
			query = FACILITY_ONLY_QUERY
		elif (not diversity_is_none) and (not splice_is_none):
			query = DIVERSITY_SPLICE_QUERY
		elif (not diversity_is_none) and splice_is_none:
			query = DC_AND_DIVERSITY_QUERY
		else:  # diversity_is_none and (not splice_is_none)
			query = FACILITY_SPLICE_QUERY

		try:
			rows = _execute_query(client, query, params)
		except Exception as e:
			logging.exception("Kusto query failed")
			return func.HttpResponse(
				json.dumps({"error": "Kusto query failed", "details": str(e)}),
				status_code=500,
				mimetype="application/json",
				headers={
					"Access-Control-Allow-Origin": "*",
					"Access-Control-Allow-Methods": "GET, POST, OPTIONS",
					"Access-Control-Allow-Headers": "Content-Type",
				},
			)

		payload = {"Spans": rows, "DataCenter": facility}
		return func.HttpResponse(
			json.dumps(payload),
			status_code=200,
			mimetype="application/json",
			headers={
				"Access-Control-Allow-Origin": "*",
				"Access-Control-Allow-Methods": "GET, POST, OPTIONS",
				"Access-Control-Allow-Headers": "Content-Type",
			},
		)

	elif stage == "Email_Template":
		# This action in Logic Apps uses Office 365 connector to send an email.
		# In code, this requires Microsoft Graph application permissions and configuration
		# for the managed identity to send mail (Mail.Send) – which cannot be done purely in code here.
		# We return a 501 to indicate configuration is required.
		return func.HttpResponse(
			json.dumps({
				"message": "Email sending requires Microsoft Graph app permissions (Mail.Send) "+
						   "granted to the Function App's managed identity, or keep this step in Logic Apps.",
			}),
			status_code=501,
			mimetype="application/json",
			headers={
				"Access-Control-Allow-Origin": "*",
				"Access-Control-Allow-Methods": "GET, POST, OPTIONS",
				"Access-Control-Allow-Headers": "Content-Type",
			},
		)

	else:
		return func.HttpResponse(
			json.dumps({"error": f"Unknown Stage '{stage}'"}),
			status_code=400,
			mimetype="application/json",
		)