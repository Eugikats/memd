# MEMD BRD gap assessment and build plan

Assessment date: 7 November 2026

Status: proposed implementation backlog; features below are not completed by this document.

Code baseline: `4e4c5cab7ebf6a087759f64e4f15673382435f56`

Delivery repository: https://github.com/Eugikats/memd (branch `main`).

## Outcome to build

An official can select a sector and period, investigate a performance change, inspect its supporting records, ask a grounded question, generate an editable cited briefing, and assign a follow-up whose outcome is recorded. Access and significant actions are controlled and auditable.

Keep the Mulwana visual shell, spacing, navigation behaviour and reusable components. Add the missing ministry workflows within that shell. Do not redesign the homepage as part of closing the BRD gaps.

## Assessment basis

- Source: `MEMD_Decision_Intelligence_BRD.docx`, version 0.1, proposed baseline dated 21 October 2026, supplied by the project owner. Its local source was `C:\Users\Administrator\Downloads\MEMD_Decision_Intelligence_BRD.docx`; this is a source reference, not a dependency required to run the application.
- Code reviewed: [Flask routes](../app.py), [ministry data and AI functions](../pulse/ministry.py), [frontend](../static/index.html), [document ingestion](../pulse/documents.py), and existing tests.
- The 2 November deployment check found the live homepage matched this code baseline, the overview API was anonymously accessible, and App Service authentication was disabled. Deployment state must be rechecked before the security milestone; this document does not claim a fresh cloud security review on 7 November.
- Existing Mulwana modules are not counted as implemented MEMD capabilities unless connected to a ministry workflow and verified. No new functional or cloud tests were performed for this documentation-only change.

## What exists now

- Flask application with a SQLite ministry store, 16 seeded measures across four periods (64 observations), five conceptual pipeline assets and telemetry history.
- Original Mulwana command-centre composition with database-backed KPI values and a decorative static wave image. The homepage wave does not plot the values. Twin telemetry charts do.
- Azure-backed sector explanation, asset analysis and optional workbook-label mapping, with local fallbacks. These are separate functions, not a complete coordinated agent workflow.
- Workbook upload, background processing, template history and download. Recognised fields can be filled; formulas/styles are preserved within the processor's limits, not recalculated.
- Editable indicator values, browser-local profile/preferences and manually executed browser-local threshold rules.
- Partial audit entries for data edits and report processing.

## Requirement-by-requirement gaps

“Partial” means a useful implementation exists but does not satisfy the complete requirement. No BRD requirement is signed off by this assessment.

| ID / BRD priority | Current status and evidence | Remaining work | Build milestone |
|---|---|---|---|
| BR-01 Must: integrated overview and sector detail | Partial. `renderCommand`, `ministryCockpitFeed`, `renderMemdCatalogue` show cross-sector indicators; no focused sector investigation flow. | Minerals, petroleum and energy detail views with period/record filters, comparable trends and links from KPIs. Projects remains an additional supported domain. | M2 |
| BR-02 Must: definition, unit, period, source, freshness | Partial. `indicators` stores unit/period/source/update time, but `overview()` substitutes a generic source and omits freshness/definitions. | Indicator catalogue with definitions, aggregation, owner, evidence source/version, reporting interval and freshness policy; show these with each measure and in exports. | M1–M2 |
| BR-03 Must: supporting approved records | Missing for sector findings. Aggregate indicators cannot be traced to production returns, deliveries or project milestones. Twin readings are only a limited example. | Record-level evidence IDs, provenance and aggregate-to-record links; protected evidence drawer with approved record/document references. | M1–M2 |
| BR-04 Must: grounded natural-language answers | Partial. `answer()` sends selected records/excerpts to Azure; `renderAgentResult()` does not render its structured `evidence` payload. | Typed read-only retrieval tools, explicit period/filter handling, usable citations, missing/conflicting evidence handling and an auditable run trace. | M3 |
| BR-05 Must: selected measures/periods/findings to cited draft | Partial. `renderMemdReports` reproduces workbooks and exports data; it is not the briefing workflow. | Report selection, editable narrative draft, citations, versioning, export and review state; correct period-aware workbook mapping. | M4 |
| BR-06 Must: missing/stale/synthetic/unavailable status | Gap and scope conflict. Generic workspace labels replaced mode labels following later owner requests. | Record truthful data mode/quality/freshness; agree unobtrusive presentation and export treatment or explicitly amend the BRD. No blanket warning-banner redesign. | M0–M2 |
| BR-07 Must: approved threshold exceptions | Partial. `renderMemdAgents` stores rules in localStorage and runs them manually. | Shared versioned rules, approval/effective dates, server evaluation, evidence-linked exceptions and duplicate prevention. | M5 |
| BR-08 Should: accountable follow-up | Missing from the active MEMD flow. Inherited supply-chain action code is not a ministry register. | Owner, due date, evidence, status history and resolution; create from an answer, finding or exception. | M5 |
| BR-09 Must: role/access limits | Missing end-to-end. `current_user()` reads an identity header but MEMD routes do not enforce roles; profile editing is browser-local. | Trusted hosted authentication, server-enforced role checks on all data/file/action routes and classification/department filters where agreed. | M1 |
| BR-10 Must: traceable user and AI actions | Partial. SQLite audit captures a subset of edits; it lacks complete actor/model/evidence/action lineage. | Common audit events with actor, outcome, run/correlation IDs, evidence versions, model/prompt versions and human dispositions. | M1, M3–M5 |
| BR-11 Should: documented validated adapters | Partial foundation only. Upload extraction and manual edits are not a governed source adapter. | Validated import contract, preview/quarantine, idempotent ingestion, source metadata and a documented example adapter. | M1, M6 |
| BR-12 Must: advisory recommendations and recorded disposition | Partial. Prompts preserve human decision authority; no integrated disposition register. | Authorised accept/reject/request-review flow with reason, actor, timestamp and evidence snapshot. No operational control or enforcement. | M5 |

## Scope decisions to record before implementation

1. **Application baseline:** later owner instructions chose Mulwana instead of the BRD's CIP-derived application and separate HTML prototype. Keep Mulwana as the active baseline; retain CIP references for future reuse. Record the change without claiming the original two-prototype deliverable is complete.
2. **Presentation:** later instructions removed synthetic/demo banners and retained the original homepage wave image. Preserve that design. Propose small metadata details and export source notes for data status; obtaining agreement or amending BR-06 is a release dependency, not permission to silently restore banners.
3. **Additional features:** keep the conceptual 3D twin and workbook reproduction, but do not let them replace sector evidence, briefing and follow-up requirements. A live engineering twin is outside this build.
4. **Data access:** start with approved dummy datasets. Live confidential feeds and production-grade SharePoint/OneDrive integrations are outside the original POC; describe them as future adapters unless explicitly added to scope.
5. **Sign-off owners:** assign a product owner, ministry data focal person, report reviewer and ICT/security approver. Confirm roles, hosting/retention, indicator meanings and the test questions. Names and dates remain unassigned.

## Delivery sequence and acceptance gates

Build thin vertical slices in the existing Flask application. Extract reusable frontend components as necessary; avoid a framework rewrite. Milestones are dependency order, not a delivery-date commitment.

### M0 — Record the agreed baseline

Deliverables:
- Record the scope decisions above and reconcile the BRD's two-prototype requirement with the Mulwana implementation.
- Capture reference screenshots for command centre, navigation, chat and reports at desktop/mobile sizes.
- Define at least one complete question-to-decision scenario per sector with expected records, calculations and citations.

Gate: product owner approves the scope/change record and test scenarios. Original shell, logo and Charles greeting remain consistent with current presentation preferences.

### M1 — Identity, governed records and audit foundation

Depends on M0's data/role decisions.

Deliverables:
- Protect the hosted app with trusted Microsoft identity; implement server permissions for viewer, analyst, steward and administrator as a proposed role set for approval. Treat client profile values as preferences, never identity or authority.
- Apply permissions to ministry endpoints, uploads, generated files and inherited Mulwana endpoints; disable irrelevant inherited routes rather than leaving alternate access paths.
- Add migrations for dataset/source, indicator definition, observation, evidence record and import batch metadata. Include stable IDs, source version, unit, interval start/end, frequency, owner, sensitivity, data mode, quality status and last successful refresh.
- Replace hard-coded latest-period logic with explicit period ordering and comparability rules. Never compare annual totals with quarterly totals as equivalent intervals.
- Add actor-attributed audit events and migration/backup/restore procedures without deleting existing workspace data.

Gate: unauthenticated access is denied; a viewer cannot edit values or retrieve a restricted document by URL; approved imports retain source identity; migration preserves existing observations; audit shows actor and outcome. Unit tests use isolated stores and mocked identity/providers, not paid cloud calls.

### M2 — Sector investigation and source transparency

Depends on M1.

Deliverables:
- Sector selection within the existing visual shell; period filters and real charts in sector detail views. Keep the home wave decorative unless the owner separately approves replacing it.
- KPI → measure detail → calculation → supporting records/document location.
- Metadata presentation for source, definition, period, freshness and limitations, using the agreed M0 treatment.
- Fixtures that cover minerals returns/licence references, petroleum product/location/date records and electricity access/reliability records. Project milestones are an additional scenario. Label examples in their metadata and avoid treating sample thresholds as regulatory rules.

Gate: changing a source record changes the expected aggregation and data-driven chart; every measure has its required metadata; stale/missing/conflicting inputs are visible; unsupported annual/quarterly comparisons are prevented; navigation works on narrow screens without changing the original design language.

### M3 — Grounded and auditable agent workflow

Depends on M1–M2.

Deliverables:
- Introduce a provider boundary for Azure and a bounded workflow: resolve question → retrieve authorised evidence → calculate deterministically → synthesise → validate citations → render evidence.
- Typed tools such as `query_indicators`, `get_supporting_records`, `search_approved_documents` and `evaluate_approved_rule`; validate tool inputs, role, sector, period and result limits. No arbitrary model-authored SQL or write access.
- Structured answers with evidence IDs and source locations, selected filters, limitations and run ID. Persist model/deployment, prompt version, retrieval/tool events and failures with appropriate retention; avoid unnecessary raw sensitive content in logs.
- Explicit missing/conflicting-evidence responses, prompt-injection handling for uploaded content and truthful fallback/provider status.

Gate: every factual claim in the agreed test set has a usable supporting reference or is explicitly unsupported (the BRD's 100% criterion); forbidden records are never retrieved; fabricated citations fail validation; unavailable Azure produces a clear limited response. Perform one separately authorised/configured Azure integration smoke test alongside deterministic local tests.

### M4 — Briefings and reliable workbook generation

Depends on M1 and M3.

Deliverables:
- Select measures, compatible reporting periods and findings; create a saved draft with editable sections, evidence references, author and revision history.
- Export an editable DOCX and readable PDF with draft/version/period/source information. Retain Excel reproduction as a separate workflow, with a mapping preview and explicit unmatched-field results.
- Fix period selection in both exact-label and placeholder mappings. Currently the placeholder dictionary is overwritten while traversing every period, so a token can select an older value while an adjacent label selects `2026 Q1`.
- Add restart-safe report jobs, duplicate submission handling and explicit processing failures. Validate upload limits, file type and workbook resource limits; do not promise unlimited files or automatic recalculation.

Gate: selected periods populate correctly regardless of database row order; report evidence resolves to the saved versions; edits survive reload; exports preserve citations/draft status; representative formulas/styles survive and calculation limits are stated; interrupted jobs recover or clearly fail instead of staying queued forever.

### M5 — Exceptions, follow-ups and human decisions

Depends on M1–M3; attach M4 drafts where applicable.

Deliverables:
- Move rules from browser storage to a shared versioned register with approved thresholds, units, scope and effective dates; server-side evaluation with evidence snapshots and idempotent exception creation.
- Follow-up register with title, accountable owner, due date, status, linked evidence/exception/report and resolution history.
- Human disposition with accept/reject/request-review, reason and actor/timestamp; permissions and audit for reassignment, changes and closure.
- Keep actions advisory and within the application. No autonomous external notices, licensing decisions, dispatch or equipment commands.

Gate: an authorised official can trace an exception through assignment and recorded resolution; repeated evaluations do not duplicate open work; unauthorised approvals fail; historical decisions retain the evidence and rule version actually reviewed.

### M6 — Adapters, acceptance and release

Depends on M1–M5.

Deliverables:
- Document a CSV/Excel source adapter contract with required schema, owner/source identity, units, periods, validation errors and idempotency key. Add a second approved sample dataset through it without changing UI logic.
- Provide import preview, quarantine/rejection reasons, reconciliation counts and refresh health. Describe SharePoint/OneDrive as optional transport adapters feeding the same validation path, with access scopes and ownership to be agreed before connecting them.
- Run the three sector scenarios end-to-end, role denial tests, citation evaluation, report export checks and desktop/mobile visual regression against M0.
- Update README/runbook, data dictionary, API/adapter notes and BRD traceability with test evidence and unresolved decisions. Keep CIP as a future-reference option.
- Deploy a reviewed build to staging, verify authenticated live paths and recovery, and record release/rollback identifiers before updating the presentation portal.

Gate: all BRD Must requirements pass their acceptance checks or have explicit owner-approved scope amendments; Should requirements have either passing checks or recorded deferrals. No claim of full BRD completion based solely on visible buttons or passing unit tests.

## Engineering entry points

| Existing area | Planned change |
|---|---|
| `app.py` | Trusted identity and permission middleware, bounded endpoints, consistent error/audit handling, removal or protection of unused inherited routes. |
| `pulse/ministry.py` | Split storage/migrations, indicator queries, evidence, rules, report mapping and provider calls into focused modules as each slice is built. Preserve current API compatibility while migrating. |
| `pulse/documents.py` | Approved document metadata, stable evidence locations, extraction limits and access checks. |
| `static/index.html` | Reuse original shell/composer; add shared measure, evidence, report editor and follow-up components; render structured citations. |
| `tests/` | Isolated database/identity/provider tests, period correctness, citation fixtures, role checks, job recovery and end-to-end UI scenarios. |

## Priority and known deferred gaps

These gaps are deliberately not fixed by this planning commit. They remain open until a later implementation commit includes verification evidence.

| Priority | Gap | Resolution |
|---|---|---|
| P0 | Publicly accessible deployment at last check; no MEMD route permissions. | Recheck and close through M1 before authorised ministry data use. |
| P1 | Placeholder period overwrite; hard-coded latest period and mixed interval comparisons. | Period model in M1; mapping regression fix in M4. |
| P1 | Missing provenance, citation rendering, governed imports and claim validation. | M1–M3, M6. |
| P1 | Briefing drafts, owner-based follow-ups and disposition capture absent. | M4–M5. |
| P1 | BRD data-status requirements conflict with later presentation requests. | M0 decision; preserve source truth without unapproved visual changes. |
| P2 | Report jobs use daemon threads without durable restart recovery. | M4. |
| P2 | Browser-local rules/preferences and incomplete audit. | M1, M3–M5; retain local-only display preferences where appropriate. |
| P2 | Homepage graphic and old scope documentation imply more chart capability than exists. | README/scope wording corrected in this commit; real sector charts belong to M2. |

Live operational integrations, calibrated engineering simulation, enterprise SLAs, national forecasts and operational enforcement remain outside the original POC. They require separate discovery and authorisation, not an implied extension of this build plan.

## Repository and release discipline

- `origin` targets `Eugikats/memd`; push this platform's work to `main` there.
- Preserve the previous Otic repository as `otic-origin` for reference. The Mulwana remote remains a baseline source only; never push MEMD changes there.
- Retain repository history and CIP reference: https://github.com/Otic-Group/CIP_Web_App. Do not reinitialise the project or replace its README with an empty template.
- Do not commit `.env`, generated workbooks, workspace databases, uploaded user documents or credentials.
- This commit publishes code history and planning documents only. It does not deploy the app, alter Azure authentication, implement the milestones or imply the Azure deployment source has switched repositories.
