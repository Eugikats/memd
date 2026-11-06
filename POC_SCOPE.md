# MEMD POC scope and presentation guide

## Meeting narrative

**Problem to test:** energy, petroleum, minerals, and infrastructure evidence can be spread across periodic reports, spreadsheets, and operator systems. The Ministry may have difficulty getting a current, comparable view of performance, exceptions, and follow-up needs. This is a stakeholder hypothesis to validate with MEMD, not a measured finding about Ministry operations.

**What this POC shows:** a single place to inspect sector indicators, ask a question against the records available to the POC, inspect a conceptual pipeline asset view, and reproduce a report workbook from stored values. The dashboard values, asset readings, and location geometry are synthetic examples stored in a local SQLite database.

Suggested walk-through:

1. Start on the Command Centre and show four sector KPI values and period changes drawn from stored indicator-period records. The wave graphic is decorative and does not plot the observations.
2. Open Data catalogue, change a latest-period indicator value, return to Command Centre, and point out that the corresponding KPI reflects the changed database record.
3. Open Pipeline digital twin, select an asset, orbit/zoom the 3D scene, inspect recent telemetry, then change a reading and show that it is persisted and plotted.
4. Ask the assistant a sector question and identify that the figures are synthetic. If Azure OpenAI is configured, it explains the records supplied to it; otherwise the local evidence response remains available. The current main chat does not expose its structured evidence payload as clickable citations; closing this gap is in the BRD build plan.
5. Upload an `.xlsx` template with an exact indicator label in one cell and a blank adjacent value cell, or a `{{MINERAL_PRODUCTION_VALUE}}` placeholder. Review the formula/style scan and download the generated workbook.
6. Close with the boundary: this is a workflow prototype. A Ministry or operator pilot needs agreed data owners, access, identifiers, reporting definitions, security review, and technical validation.

## Functional scope today

| Workflow | POC behavior | Stored state |
|---|---|---|
| Sector dashboard | Four database-backed KPIs with period changes and a static decorative wave | `indicators` table |
| Pipeline twin | WebGL three-dimensional pipeline scene, asset selection, latest measurements, telemetry history | `assets` and `telemetry` tables |
| Telemetry adjustment | Validate and append a reading; flag high vibration or low pressure for review | `telemetry`; high-severity assets update in `assets` |
| Sector assistant | Offline evidence answer; optional Azure OpenAI summary constrained to database facts | Database is read-only to the answer path |
| Workbook report agent | Inspect `.xlsx` sheet geometry, formulas, labels, and style samples; populate exact matches/placeholders; append data extract | `report_templates`, generated workbook, `audit` |
| Data catalogue | View measures and edit stored indicator values | `indicators`; `audit` |

The routes are implemented in `app.py`; SQLite and report processing are in `pulse/ministry.py`; Ministry pages are rendered in `static/index.html`.

## Synthetic seed profile

- 4 sectors: minerals, petroleum, energy, and strategic projects.
- 16 measures over four reporting periods (64 sector-indicator observations).
- 5 conceptual pipeline assets, with 30 daily sample telemetry points each.
- Example high-vibration asset to exercise exception handling.
- Local storage path: `data/memd.sqlite3`; uploaded template outputs: `data/uploaded_templates/`.

Changing indicator values changes the overview API and relevant KPI values after reloading; the homepage wave image stays unchanged. Changing telemetry appends a timestamped database row and refreshes the selected asset chart. Seed data is created on initialisation when the indicator table is empty; to reset the local seed database, stop the server and remove `data/memd.sqlite3` before the next start.

See the [BRD gap assessment and build plan](docs/BRD_BUILD_PLAN.md) for outstanding requirements and implementation acceptance gates.

## Report reproduction boundary

The report agent accepts `.xlsx` up to 20 MB. It uses exact indicator-name matching, explicit `{{INDICATOR_NAME}}` tokens, and a conservative adjacent-cell rule. It does not infer a complex report's full business semantics or formula dependencies automatically. It preserves formulas and cell formatting in tested workbooks, but it does not recalculate formulas. Macros, pivots, charts, embedded objects, links, print layouts, and advanced workbook features may not survive conversion. Use an approved, non-sensitive workbook for POC testing, review all output in Excel, and agree the output mapping with report owners before production use.

## Decision and safety boundary

Pipeline geometry and readings are fictional. No live API, sensor, SCADA, GIS, operator, or Ministry connection exists. Thresholds are illustrative and are not operational alarms. The twin is not an engineering digital twin and must not be used for operations, safety decisions, maintenance scheduling, or investment approval. Human officials and the relevant infrastructure operators retain decision authority.

The optional Azure model provides language synthesis only. The POC does not permit the model to change database values, issue commands to field equipment, or approve a report or decision.

## Path to an operator-connected feasibility study

1. Confirm the Ministry's specific decision to support and whether MEMD, UNOC, EACOP, or another party owns the required information.
2. Agree a read-only, non-sensitive sample: asset register, approved locations, measurement names/units, status history, and reporting frequency.
3. Ask the data owner to confirm interfaces, access rights, retention, security requirements, and what can be shown to Ministry users.
4. Validate the asset identity/coordinate model and measurement thresholds with the operator's engineering team.
5. Define one measurable pilot outcome (for example, how quickly an authorised user identifies and routes a declared telemetry exception) and an evaluation baseline.
6. Estimate connectivity, data engineering, 3D asset modelling, hosting, licensing, cyber review, and support before offering a full digital-twin delivery commitment.

## Suggested meeting language

“We have prepared a concept that brings sector evidence, questions, reports, and an illustrative pipeline asset view into one workspace. The values and telemetry you see today are synthetic and the twin is not live. We would like to learn whether there is a priority Ministry decision to support, who owns the data required for it, and whether a small read-only feasibility study would be useful. If there is, we will define scope and cost only after confirming the data access and operator validation requirements.”
