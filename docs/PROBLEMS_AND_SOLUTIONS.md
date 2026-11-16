# MEMD Problems & Solutions Decision Log

This document catalogues **all** identified business, technical, security, and testing problems in the current MEMD POC (`HEAD` `72fa78a`) and defines the concrete solution for each. It complements [BRD_BUILD_PLAN.md](BRD_BUILD_PLAN.md) with implementation-level decisions tied to the existing codebase ([app.py](../app.py), [pulse/ministry.py](../pulse/ministry.py), [pulse/documents.py](../pulse/documents.py), [static/index.html](../static/index.html), and `tests/`).

---

## Priority Summary

| Priority | Problems | Why First |
|---|---|---|
| **P0 (Blocker)** | 1. Unauthenticated endpoints & no RBAC (`BR-09`) | Cannot put real Ministry data on an open app. |
| **P1 (Core BRD)** | 2–8, 11–12: Metadata/provenance (`BR-02`, `BR-03`), period integrity, sector drill-downs (`BR-01`), data-status clarity (`BR-06`), grounded AI citations (`BR-04`), Excel period bug & briefing drafts (`BR-05`), follow-ups & dispositions (`BR-08`, `BR-12`), full audit (`BR-10`) | Required for trustworthy numbers, cited AI/briefings, and official accountability. |
| **P2 (Hardening)** | 9–10, 13–14: Durable report jobs, shared server-side rules (`BR-07`), validated data adapters (`BR-11`), expanded test suite | Required for reliability, scale, and operational handover. |

---

## Detailed Problems & Solutions

### 1. No Authentication, Role Enforcement, or Route Isolation (`BR-09` — P0)
- **Problem:** `current_user()` in `app.py` reads `X-MS-CLIENT-PRINCIPAL-NAME`, but `/api/memd/*`, `/api/ask`, file uploads/downloads, and inherited Mulwana routes (`/api/cockpit`, `/api/actions`, `/api/alerts`, `/api/reports/*`, `/api/excel/*`, `/api/schedules/*`) enforce no authentication or role checks. User profiles are stored in browser `localStorage`.
- **Solution:**
  1. Enable trusted hosted authentication (Microsoft Entra ID / App Service Auth) and reject unauthenticated requests on all `/api/*` routes in `app.py` (with a local dev/test auth stub controlled strictly by test fixtures, never active in production).
  2. Introduce server-side RBAC middleware with 4 roles: `viewer`, `analyst`, `steward`, `administrator`.
  3. Enforce role checks per route (e.g., `viewer` cannot call `/api/memd/indicator`, `/api/memd/twin/*/telemetry`, or `/api/documents` upload; restrict document downloads by role/classification).
  4. Disable or remove unused inherited Mulwana routes from `app.py` so they cannot act as unguarded backdoors.
  5. Keep `localStorage` strictly for UI display preferences, never identity or permissions.

---

### 2. Missing Indicator Definitions, Freshness & Record Provenance (`BR-02`, `BR-03` — P1)
- **Problem:** `indicators` in `pulse/ministry.py` stores `sector, metric, unit, period, value, source, updated_at`, but `overview()` substitutes a generic source and omits definitions, aggregation rules, owners, reporting intervals, and freshness policies. Aggregate numbers cannot be traced back to supporting approved records (production returns, licence references, deliveries, milestones).
- **Solution:**
  1. Add SQLite schema migrations for:
     - `indicator_definitions` (stable ID, sector, metric, definition, aggregation method, unit, reporting interval, owner, freshness SLA, sensitivity).
     - `datasets` / `import_batches` (source system, version, ingested_at, data_mode, quality_status).
     - `evidence_records` (record ID, indicator_id, period, source_ref, document_id/location, value, approved_by, approved_at).
  2. Update `overview()` and add `/api/memd/indicator/<id>/evidence` so every KPI and catalogue row returns its definition, owner, source version, freshness status, and linked supporting records.
  3. Add a protected Evidence Drawer in `static/index.html` to inspect supporting records and approved document references from any KPI.

---

### 3. Hard-Coded Latest Period & Mixed-Interval Comparisons (`BR-01`, `BR-02` — P1)
- **Problem:** Period handling relies on string sorting / hard-coded latest-period assumptions, with no guard against comparing incompatible intervals (e.g., annual totals vs. quarterly totals).
- **Solution:**
  1. Store explicit `interval_start`, `interval_end`, and `frequency` (`monthly`, `quarterly`, `annual`) on every observation and period definition.
  2. Replace string-based "latest period" logic with deterministic chronological ordering (`interval_end DESC`).
  3. Enforce interval-compatibility checks in trend queries, period comparisons, and briefing selections so annual and quarterly figures are never compared as equivalent.

---

### 4. No Focused Sector Investigation Views or Real Sector Charts (`BR-01` — P1/P2)
- **Problem:** `renderCommand` and `renderMemdCatalogue` in `static/index.html` show cross-sector summaries, while the homepage wave graphic is a static decorative image. Users cannot drill into Minerals, Petroleum, or Energy with period filters and real trend charts.
- **Solution:**
  1. Keep the Mulwana command-centre shell and decorative wave intact as agreed in `BRD_BUILD_PLAN.md`.
  2. Add dedicated sector investigation views (`Minerals`, `Petroleum`, `Energy`, plus `Projects`) accessible directly from homepage KPIs.
  3. Render real data-driven trend charts (using the same charting pattern as twin telemetry) backed by SQLite, with period/record filters and links into the Evidence Drawer.

---

### 5. Data Status / Synthetic-Mode Presentation Conflict (`BR-06` — P1)
- **Problem:** Generic workspace labels replaced mode labels per earlier owner requests, leaving users unable to tell from the screen or exports whether data is live, synthetic/sample, stale, or missing.
- **Solution:**
  1. Record truthful `data_mode` (`synthetic`, `staging`, `approved`), `quality_status`, and `last_refreshed` in the database per dataset/indicator.
  2. Surface status unobtrusively via small per-indicator metadata badges/tooltips in the catalogue/detail views and mandatory source-status notes in all exports (Word, PDF, Excel) — avoiding full-width warning banners on the homepage.
  3. Obtain formal M0 sign-off on this presentation approach.

---

### 6. Unrendered & Unvalidated AI Citations + Unbounded Retrieval (`BR-04` — P1)
- **Problem:** `ministry.answer()` in `pulse/ministry.py` sends selected records/excerpts to Azure OpenAI and returns an `evidence` payload, but `renderAgentResult()` in `static/index.html` does not render structured citations, there is no claim-to-citation validation, and retrieval/tool calls are not strictly typed or role-scoped.
- **Solution:**
  1. Refactor `ministry.answer()` into a bounded pipeline:
     `resolve question -> typed read-only retrieval tools -> deterministic calculations -> AI synthesis -> citation validation -> structured response`.
  2. Implement typed tools (`query_indicators`, `get_supporting_records`, `search_approved_documents`, `evaluate_approved_rule`) with strict parameter, role, sector, period, and row-limit validation (no model-authored SQL).
  3. Validate that every factual claim in the response maps to a retrieved `evidence_id`; reject/flag unsupported or fabricated citations and return explicit missing/conflicting-evidence messages when data is absent.
  4. Render clickable citations (`evidence_id`, period, source location) in `renderAgentResult()`, and add prompt-injection sanitization for uploaded document text.

---

### 7. Period-Overwrite Bug in Excel Template Mapping (`BR-05` — P1)
- **Problem:** In `pulse/ministry.py` (`_inspect_and_generate_path` / `_ai_report_matches`), traversing all periods overwrites the placeholder dictionary, causing `{{INDICATOR_NAME}}` tokens to populate an older period's value while adjacent labels populate `2026 Q1`.
- **Solution:**
  1. Require an explicit target `period` parameter (defaulting to the chronologically latest compatible period, never arbitrary row order) when building the indicator lookup map.
  2. Namespace period-specific tokens if multi-period templates are used (e.g., `{{METRIC::2026_Q1}}`) and ensure deterministic single-period resolution.
  3. Add a mapping-preview endpoint and return an explicit list of unmatched fields before/with generation.
  4. Add a regression test with out-of-order SQLite rows to guarantee the selected period is always populated.

---

### 8. Missing Briefing Workflow & DOCX/PDF Export (`BR-05` — P1)
- **Problem:** `renderMemdReports` only reproduces uploaded `.xlsx` templates and exports raw data. Officials cannot select measures/periods/findings, write/edit a cited narrative draft, track review state, or export DOCX/PDF briefings.
- **Solution:**
  1. Add `briefings` and `briefing_revisions` tables in SQLite storing selected measures, compatible periods, linked evidence/citations, editable narrative sections, author, version, and review state (`draft`, `in_review`, `approved`).
  2. Build a Briefing Editor UI in `static/index.html` where officials compose/edit drafts and insert verified citations from Q&A or sector views.
  3. Implement DOCX export using `python-docx` (already in `requirements.txt`) and PDF export (e.g., via headless LibreOffice or pure-Python PDF generation), stamping draft/version/period/source-status metadata on every export.
  4. Keep Excel template reproduction as a separate workflow alongside Briefings.

---

### 9. Non-Durable Background Report Jobs (`BR-05` — P2)
- **Problem:** `create_report_job()` / `_run_report_job()` in `pulse/ministry.py` spawns in-memory daemon threads. If the server restarts, queued/running jobs stay stuck in `queued`/`running` forever.
- **Solution:**
  1. Persist job state transitions (`queued`, `running`, `ready`, `failed`), attempt counts, and idempotency/upload hashes in `report_jobs`.
  2. On app startup, mark any orphaned `running` jobs as `failed` (with a clear restart error) or re-queue idempotently.
  3. Enforce upload size/type/sheet limits and handle duplicate submissions cleanly.

---

### 10. Browser-Local Rules & No Server-Side Exception Engine (`BR-07` — P1/P2)
- **Problem:** `renderMemdAgents` stores threshold rules in `localStorage` and evaluates them manually in the browser. There is no shared approved rule register, versioning, or duplicate-safe exception generation.
- **Solution:**
  1. Add `rules` (rule_id, version, sector, metric, operator, threshold, unit, effective_from, effective_to, approved_by, status) and `exceptions` (exception_id, rule_id, rule_version, period, evidence_snapshot_json, status, created_at) tables in SQLite.
  2. Evaluate approved rules server-side (on data refresh/edit or manual trigger) with idempotency keys (`rule_id + rule_version + indicator_id + period`) so repeated runs never create duplicate open exceptions.

---

### 11. No Accountable Follow-Up Register or Human Disposition Workflow (`BR-08`, `BR-12` — P1)
- **Problem:** Inherited `/api/actions` is supply-chain code, not a ministry follow-up register. There is no way to assign an officer and due date to a finding/exception/briefing or record an authorised `accept` / `reject` / `request-review` disposition with the exact evidence reviewed.
- **Solution:**
  1. Add `followups` (id, title, source_type, source_id, owner, due_date, status, resolution_notes, created_by, updated_at) and `dispositions` (id, target_type, target_id, decision [`accept`|`reject`|`request_review`], reason, actor, decided_at, evidence_snapshot_json, rule_version) tables.
  2. Expose CRUD/transition endpoints protected by RBAC (`analyst`+ to create/assign, `steward`/`administrator` to approve/close).
  3. Keep all actions strictly advisory inside the app (no autonomous external notices or operational commands).

---

### 12. Incomplete Audit Trail (`BR-10` — P1/P2)
- **Problem:** `audit(id, at, action, detail)` in `pulse/ministry.py` only logs a subset of edits and report runs, missing actor identity, outcome, correlation/run IDs, AI model/prompt versions, and evidence versions.
- **Solution:**
  1. Standardize a common `audit_events` schema: `id, at, actor, role, action, resource_type, resource_id, outcome, correlation_id, model_version, prompt_version, evidence_ids_json, detail_json`.
  2. Emit audit events centrally via a helper for all data edits, imports, AI queries/tool calls, briefing exports, rule evaluations, follow-ups, and dispositions.

---

### 13. No Validated Data Ingestion / Adapter Contract (`BR-11` — Should)
- **Problem:** Manual indicator edits and ad-hoc uploads lack schema validation, preview/quarantine for bad rows, reconciliation counts, and idempotency.
- **Solution:**
  1. Define a documented CSV/Excel adapter contract (required columns: `sector, metric, unit, period, interval_start, interval_end, frequency, value, source_ref, owner`).
  2. Build a two-step import flow: `Preview & Validate` (returns valid count, quarantined/rejected rows with reasons, and reconciliation totals) -> `Commit` (idempotent by batch hash/source key).
  3. Load a second approved sample dataset through this adapter to prove schema-driven ingestion without UI changes.

---

### 14. Insufficient Test Coverage (`M1`–`M6`)
- **Problem:** Current tests (11 Python unit tests + 1 JS syntax check) only cover POC happy paths and offline fallbacks. Security, period logic, citations, job recovery, and UI workflows are untested.
- **Solution:**
  Expand `tests/` alongside each milestone:
  - **M1:** Unauthenticated denial, RBAC role-matrix tests, restricted document URL denial, migration preservation, audit actor attribution.
  - **M2:** Mixed annual/quarterly rejection, aggregate-to-record recalculation, metadata completeness.
  - **M3:** 100% claim-to-citation validation, fabricated-citation rejection, prompt-injection tests, plus 1 gated Azure smoke test.
  - **M4:** Out-of-order SQLite row period-mapping regression test, DOCX/PDF export checks, interrupted job recovery.
  - **M5:** Idempotent exception creation, unauthorised disposition denial, historical snapshot preservation.
  - **M6:** End-to-end sector scenarios and desktop/mobile UI checks.

---

## Execution Order (Aligned to `M0`–`M6`)
1. **M0 (Decisions):** Assign sign-off owners, confirm `BR-06` metadata presentation, and define 1 test scenario per sector.
2. **M1 (Security & Data Foundation):** Auth/RBAC + disable unused Mulwana routes + schema migrations + period ordering + common audit.
3. **M2 (Sector Investigation):** Sector detail views, real trend charts, indicator metadata, and Evidence Drawer.
4. **M3 (Grounded AI):** Typed retrieval tools, deterministic calculations, citation validation, and UI citation rendering.
5. **M4 (Briefings & Excel Fix):** Fix Excel period-overwrite bug, add mapping preview, durable jobs, and DOCX/PDF briefing drafts.
6. **M5 (Accountability):** Server-side rules/exceptions, follow-up register, and human disposition workflow.
7. **M6 (Adapters & Acceptance):** CSV/Excel import adapter with quarantine, full test suite, staging verification.

