# MEMD Decision Intelligence — Meeting Brief

## The problem

Minerals, petroleum, electricity, and project information sits across reports, spreadsheets, and systems owned by different teams. Officials can spend time reconciling it before they can compare performance, spot exceptions, or prepare a decision. The POC tests whether one evidence workspace can shorten that path.

**Magnitude:** MEMD has not yet supplied a measured cost or delay baseline. Measure staff hours per reporting cycle, time to answer cross-sector questions, late or missing returns, and time from an exception to an assigned follow-up. Do not present a financial saving until those baselines are agreed.

## How the system helps

- **Unify:** load approved datasets, retain source and reporting period, map shared measures and identifiers, and flag gaps or inconsistent units.
- **Analyse:** AI answers questions from retrieved records, compares periods, highlights threshold exceptions, and shows the evidence behind each finding.
- **Produce:** a report agent reads an uploaded Excel template, matches agreed indicator names or placeholders, fills available values, and returns a workbook for review.
- **Support action:** route findings to an official or data owner for follow-up. The system advises; authorised people make decisions.

## Data sources

Proposed sources to confirm with MEMD data owners:

- **Minerals:** MEMD mineral licensing, production returns, inspections, and revenue records.
- **Petroleum:** import, demand, stock, and distribution records held by MEMD and the responsible supply-chain entities.
- **Energy:** generation, access, reliability, and distribution measures from MEMD and the relevant regulators, utilities, and agencies.
- **Projects and pipeline assets:** project schedules and milestones from project owners; asset registers and approved telemetry from UNOC, EACOP, and relevant operators where they authorise access.
- **Published evidence:** official Ministry and government reports for initial context and cross-checking.

These are candidate owners and datasets, not confirmed live connections. The current POC runs on a local database; the Ministry must confirm the authoritative source, owner, definitions, identifiers, update frequency, and access rights for each feed.

## Getting data into the platform

Start with approved Excel or CSV extracts. For recurring updates, connect a Ministry SharePoint or OneDrive library so an authorised process can ingest new files, preserve their source and version, validate columns, and refresh the workspace. APIs or a secure database connection can be added for systems that support them. Every feed needs an agreed owner, permission, schedule, and error path.

## Proposed first measure of value

Choose one recurring Ministry question or report. Record its current preparation time and data gaps, run the same task through the workspace, then compare time, completeness, and reviewer corrections with the data owner.
