# Pulse – Supply Chain Reporting POC (Mulwana Group / Nice House of Plastics)

Working proof of concept for the **Supply Chain Reporting Dashboards** and **Supply Chain Command Centre**
described in the Otic technical proposal (12 Oct 2026), connected read-only to Business Central `NHOP-TEST`.

## Run
```
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
copy .env.example .env          # fill in BC credentials (+ optional ANTHROPIC_API_KEY)
.venv\Scripts\python -m pulse.extract --days 90   # pull from Business Central (~2 min)
.venv\Scripts\python -m pulse.model               # build the model + reconciliation
.venv\Scripts\python app.py                       # http://localhost:8050
```
The **Refresh from Business Central** button in the app re-runs extract + model on demand (logged to `data/refresh_log.jsonl`).

## Architecture (matches proposal ch. 7.1)
| Step | Code | Notes |
|---|---|---|
| Authenticate | `pulse/bc_client.py` | Entra ID client-credentials; token cached/refreshed |
| Extract | `pulse/extract.py` | Paged (20k/page), retry on 429/5xx, read-only GETs only |
| Stage | `data/raw/*.json.gz` + `run_log.json` | Records kept as received for traceability |
| Model + Calculate | `pulse/model.py`, `pulse/config.py` | One governed model; all measures |
| Present | `pulse/views.py`, `static/index.html` | 11 dashboards, filters, drill-down, CSV export |
| Command centre | `pulse/command.py` | Claude (`claude-opus-5-5`) with read-only tools over the same views; audit log `data/audit.jsonl` |

### Business Central sources
| Data | Endpoint |
|---|---|
| Items, vendors | API v2.0 `items`, `vendors` |
| Item Ledger Entries (open → stock; 90-day movements → usage) | OData `ItemLedgerEntries` (has Location Code, Remaining Qty) |
| Purchase Lines (outstanding orders) | OData `Purchase_Lines_Excel` |
| Production BOM header/lines | OData `Production_BOM_Excel`, `Production_BOM_ExcelProdBOMLine` |
| Locations | OData `Location_Card_Excel` |

The standard v2.0 `itemLedgerEntries` API has no location code, so the published OData pages are used for the
POC. For production, the proposal's small read-only API extension should replace them.

## Configuration to review with the Supply Chain team
`pulse/config.py`: item class by inventory posting group, entry types counted as consumption/demand,
cover thresholds (7 / 14 / 30 / 180 days), site grouping of locations. `pulse/model.py`: stale PO threshold (90 days)
and overdue-arrival assumption (7 days). All measure definitions are shown in the app under *Definitions & data notes*.

## Findings from the first load (16 Oct 2026)
- Stock reconciles to BC item inventory for 2,735 of 2,744 items (99.7%); differences are mainly NI- items with negative inventory.
- All 359 outstanding PO lines are past their expected receipt date; 255 are >90 days overdue (stale).
- 3,755 of 7,584 certified BOM item lines have Quantity per = 0 and are excluded from WIP/readiness calculations.
- The company ID supplied (`e813eea0-…`) does not exist in NHOP-TEST; the POC uses **Nice House of Plastics** (`fe9d99b4-…`).

## Agent (Command Center)
The Command Center is agentic: it plans multi-step work, gathers evidence from the governed views, then creates
**proposed actions** in Pulse (`pulse/agent.py`). It never writes to Business Central.

| Action | What it produces | After approval |
|---|---|---|
| Chase overdue POs | One drafted expedite message per vendor (most overdue value first) | Copy / download and send from your mail |
| Plan stock transfers | Inter-site transfer plan from location shortages vs. surplus | Export CSV to raise transfer orders in BC |
| Alerts | Rule on a model metric (stockouts ≤7d, overdue lines, high exceptions, constrained products, item cover) | Active; re-checked after every refresh |
| Watchlist | Items pinned to the Overview | — |
| Management briefing | Markdown briefing with ranked tables and recommendations | Download |

With `ANTHROPIC_API_KEY` set, Claude decides which tools to call for any request. Without it, requests are
matched to the five built-in playbooks (buttons under the ask box). Proposals live in `data/actions.json`; every
run is logged in `data/audit.jsonl`. Review everything on the **Agent actions** page.

## Azure deployment (16 Oct 2026)
| Resource | Name | Notes |
|---|---|---|
| Subscription | Microsoft Azure Sponsorship (`5f5765ae-…`), tenant Otic Group | |
| Resource group | `rg-pulse-mulwana-poc` | Everything for this POC lives here |
| Web app | `pulse-mulwana-mooy0` → https://pulse-mulwana-mooy0.azurewebsites.net | Linux, Python 3.14, plan `asp-pulse-mulwana-wus2` (B2, West US 2) |
| Sign-in | App Service Authentication, Entra app `Pulse Mulwana Supply Chain POC` (`2c8983af-…`) | Otic Group tenant accounts only; unauthenticated requests are refused |
| Azure OpenAI | `oai-pulse-mulwana-mooy0` (East US), deployment `gpt-5.4` (DataZoneStandard, 100K TPM) | Web app uses its managed identity (role *Cognitive Services OpenAI User*); no key stored |
| Data | `/home/data` (persistent) | Seeded from the bundled snapshot on first start; **Refresh from Business Central** reloads live |

App settings hold the BC connection (`BC_*`), `AZURE_OPENAI_ENDPOINT`, `AZURE_OPENAI_DEPLOYMENT`, `PULSE_DATA_DIR`.
Redeploy: zip `app.py requirements.txt README.md pulse/ static/ data/raw/` and run
`az webapp deploy -g rg-pulse-mulwana-poc -n pulse-mulwana-mooy0 --src-path pulse_deploy.zip --type zip`.
To let Mulwana staff sign in, invite them as guests in the Otic Group tenant (or move the app registration to Mulwana's tenant).

## Reporting & governance (added 16 Oct 2026)
| Page | What it does | Code |
|---|---|---|
| Report library | 16 standard reports (stock status, valuation, cover, open/overdue POs, vendor performance, stockout exposure, risk register, rebalancing, production readiness, exceptions, KPI summary, weekly movement trends, data quality & reconciliation). Preview, filter, export **Excel / CSV / Print-PDF**, save to history, schedule | `pulse/reports.py` |
| Report builder | Custom reports on 7 governed datasets: pick columns or group & summarise (sum/avg/min/max/count), conditions, sort, limit; export or save under *My reports* (then schedulable) | `run_custom()` |
| Schedules & history | Daily / weekdays / weekly / monthly runs (East Africa Time); every run saves an Excel file to history. Email delivery is a next step (needs an approved mail sender) | `pulse/scheduler.py` |
| Trends | Weekly flow (consumption, output, sales, receipts) and a KPI snapshot per data load that builds up through the proof period | `kpi_history.jsonl` |
| Data refresh | Scheduled BC refresh (default 06:00 EAT, Mon–Sat), run log with record counts and reconciliation %, refresh now | `/api/refresh`, `/api/schedules/refresh` |
| Audit log | Questions, agent runs, approvals, report exports/generation and refreshes, with user and time | `/api/audit` |

The agent has a `generate_report` tool: "give me an Excel of vendor delivery performance" produces the file and an action card.
