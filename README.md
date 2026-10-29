# MEMD Decision Intelligence POC

A Ministry of Energy and Mineral Development decision-intelligence proof of concept, built on the Mulwana POC's web application shell. The Ministry version keeps the source app's compact navigation and card-based visual style while replacing its visible workflows with sector indicators, a pipeline-asset twin, evidence questions, a data catalogue, and report-template processing.

## Run locally

From PowerShell in this folder:

```powershell
\.venv\Scripts\python.exe app.py
```

Open [http://127.0.0.1:8054](http://127.0.0.1:8054). If the port is occupied, set another port first:

```powershell
$env:PORT = "8055"
\.venv\Scripts\python.exe app.py
```

If `.venv` has not been created yet:

```powershell
python -m venv .venv
\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## What works in this POC

- **Command Centre:** sector indicator trends for minerals, petroleum, energy, and strategic projects, drawn from the local database.
- **Pipeline digital twin:** interactive 3D station and pipeline scene, asset selector, telemetry history, and editable readings. Changes are written to SQLite and immediately affect the asset view.
- **Ask the assistant:** Azure OpenAI analyses the records supplied from the SQLite workspace, cites periods and measures, and states evidence limits. If Azure is unavailable, a deterministic evidence response remains available.
- **Pipeline monitoring:** request an Azure analysis of an asset's recent telemetry. It identifies measured changes and alert readings, while leaving safety conclusions and operating decisions to the operator.
- **Reports:** upload an Excel `.xlsx` template (up to 20 MB). The processor inspects sheets, formula cells, labels, and cell formatting; Azure can map clear report labels to known indicators, while exact labels and `{{INDICATOR_NAME}}` tokens are handled locally. Existing formulas and styles are retained, and a sector-data sheet is appended for download.
- **Data catalogue:** inspect and edit persisted indicator values, and open pipeline assets in the twin.
- **MEMD mark:** a local SVG lettermark replaces the source product logo.

## Data and limitations

The first start seeds `data/memd.sqlite3` with synthetic POC records. They are not an official Ministry, UNOC, EACOP, TotalEnergies, SCADA, or operator feed. The indicators and asset telemetry exercise the workflows; they must not be used as sector facts, safety information, or operating instructions.

The twin is a conceptual visualisation, not a surveyed route, validated engineering model, safety system, or live digital replica. This workspace is connected to the Azure OpenAI deployment `gpt-5.4-mini` in resource group `rg-memd-dev`. Local authentication uses the signed-in Azure CLI identity (`az login`) with the `Cognitive Services OpenAI User` role; a hosted deployment should use a managed identity with that role. Copy `.env.example` to `.env` and set the Azure endpoint and deployment name for another environment. Azure is used for evidence synthesis, report-label mapping, and asset telemetry analysis; calculations and rule checks remain deterministic.

Excel processing uses `openpyxl`. It preserves workbook formulas and cell styles, but does not calculate formulas or guarantee retention of macros, pivot caches, embedded objects, external connections, or every advanced Excel feature. Azure mappings are checked against known indicator names, periods, worksheet labels, and a confidence threshold before a value is populated. Review the generated workbook in Excel before relying on it.

All decisions remain with authorised Ministry officials and pipeline operators. This POC does not write to external operational systems.

## Local architecture

| Area | Implementation |
|---|---|
| Web app and APIs | Flask (`app.py`) |
| Sector indicators and asset telemetry | SQLite (`data/memd.sqlite3`) via `pulse/ministry.py` |
| Original source analytics shell | `pulse/` and `static/index.html` |
| 3D pipeline scene | Locally bundled Three.js WebGL scene (`static/vendor/`); works without a CDN connection |
| Excel inspection and output | `openpyxl` in `pulse/ministry.py` |
| Generated reports and audit | `data/uploaded_templates/` and SQLite audit table |

The bundled Three.js runtime is distributed under its MIT license in [static/vendor/THREE-LICENSE.txt](static/vendor/THREE-LICENSE.txt).

See [POC_SCOPE.md](POC_SCOPE.md) for the meeting narrative, supported interactions, and follow-on steps toward an operator-connected pilot.
