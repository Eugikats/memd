# Azure Excel Intelligence Engine — infrastructure

Provisions Blob Storage, Service Bus, Azure AI Search, Key Vault references, and a Container Apps Job that runs `python -m excel_engine --from-servicebus`.

Pulse App Service remains the conversational front door. Set these app settings after deploy:

| Setting | Purpose |
|---------|---------|
| `PULSE_EXCEL_STORAGE_ACCOUNT` or `PULSE_EXCEL_STORAGE_CONNECTION` | Workbook Blob store |
| `PULSE_EXCEL_BLOB_CONTAINER` | Default `excel-workbooks` |
| `PULSE_EXCEL_SERVICEBUS` | Service Bus connection string |
| `PULSE_EXCEL_SERVICEBUS_QUEUE` | Default `excel-jobs` |
| `PULSE_EXCEL_SEARCH_ENDPOINT` / `PULSE_EXCEL_SEARCH_KEY` | Workbook intelligence index |
| `PULSE_KEY_VAULT_URI` | Optional secret hydration |
| `PULSE_EXCEL_LOGIC_APP_URL` | Optional ready-job webhook |
| `PULSE_GRAPH_EXCEL_DRIVE_ID` | Optional Graph Excel recalc |
| `PULSE_TENANT_ID` | Blob/search path isolation prefix |

Without these, Pulse falls back to local disk under `PULSE_DATA_DIR/excel` and an in-process worker thread.

## Deploy

```bash
# From repo root
./infra/excel_engine/deploy.sh <resource-group> <location> <prefix>
```

Or apply Bicep:

```bash
az deployment group create -g <rg> -f infra/excel_engine/main.bicep \
  -p prefix=<prefix> location=<location>
```

## Worker image

Build and push the Container Apps Job image:

```bash
az acr build -r <acr> -t pulse-excel-engine:latest -f infra/excel_engine/Dockerfile .
```

## Logic Apps / Functions delivery

Point a Logic App HTTP trigger at `PULSE_EXCEL_LOGIC_APP_URL`. On `job.ready`, Pulse POSTs `{job_id, history_id, download_path, user, tenant, …}`. Alternatively call `POST /api/excel/jobs/<id>/deliver` with `{recipients:[…]}` to email via Graph.
