"""End-to-end Excel job: extract → index → rebuild → validate → optional Graph recalc → deliver."""
from __future__ import annotations

import json

from excel_engine import delivery, extract, graph_excel, rebuild, search_index, validate
from excel_engine import keyvault as kv


def run_job(job_id: str, model=None) -> dict:
    from pulse import documents
    from pulse import excel_blob
    from pulse import excel_jobs

    kv.hydrate_env_from_vault()
    job = excel_jobs.get(job_id)
    if not job:
        raise ValueError(f"Unknown job {job_id}")

    def bump(status, progress, message, **extra):
        return excel_jobs.update(job_id, status=status, progress=progress, message=message, **extra) or {}

    doc = documents.get(job["document_id"])
    if not doc:
        return bump("failed", 100, "Upload expired", error="That upload is no longer available. Attach it again.")

    # Ensure bytes available (blob-backed)
    if "bytes" not in doc or not doc["bytes"]:
        key = doc.get("blob_key") or doc.get("uri")
        if not key:
            return bump("failed", 100, "Upload bytes missing", error="Upload storage missing.")
        try:
            doc["bytes"] = excel_blob.get_bytes(key)
            documents.cache_bytes(job["document_id"], doc["bytes"])
        except Exception as e:
            return bump("failed", 100, "Could not load workbook", error=str(e)[:300])

    if model is None:
        try:
            from pulse import model as M
            model = M.load()
        except Exception as e:
            return bump("failed", 100, "Model unavailable", error=f"Could not load supply chain model: {e}"[:300])

    name = doc.get("name") or "workbook.xlsx"
    if name.lower().endswith((".xlsx", ".xlsm")):
        return _run_xlsx(job_id, job, doc, model, bump)

    bump("extracting", 10, "Extracting worksheets, formulas, and layout (fast scan)")
    extracted = extract.extract_workbook(doc["bytes"], name, fast=True)
    extract_key = f"{job.get('tenant', 'default')}/extracts/{job_id}.json"
    excel_blob.put_bytes(extract_key, json.dumps(extracted).encode("utf-8"), "application/json")
    bump("extracting", 25, "Extraction complete", extract=extracted.get("summary"),
         unsupported=extracted.get("unsupported") or [])

    bump("indexing", 35, "Indexing workbook intelligence")
    indexed = search_index.index_extract(job_id, job["document_id"], job.get("tenant") or "default", extracted)
    bump("indexing", 45, f"Indexed {indexed.get('indexed', 0)} chunks ({indexed.get('backend')})")

    bump("rebuilding", 55, "Rebuilding template with system figures (deterministic)")

    def on_progress(i, n, sheet_name):
        pct = 55 + int(20 * (i / max(n, 1)))
        bump("rebuilding", min(pct, 74), f"Filling sheet {i}/{n}: {sheet_name}")

    filled = rebuild.rebuild_from_document(
        model, job["document_id"], job.get("view"), on_progress=on_progress,
        column_map=job.get("column_map"))
    if filled.get("error"):
        return bump("failed", 100, "Rebuild failed", error=filled["error"],
                    unmatched=filled.get("unmatched") or [], sheets=filled.get("sheets") or [])

    bump("validating", 75, "Validating structure and formulas")
    report = validate.validate_rebuild(extracted, filled["bytes"], filled)
    if report["status"] == "fail":
        return bump("failed", 100, "Validation failed", error="; ".join(report["errors"])[:400],
                    validation=report, unsupported=report.get("unsupported") or [])

    recalc = graph_excel.recalculate(filled["bytes"], filename=(doc.get("name") or "report.xlsx")[:120])
    filled["bytes"] = recalc.get("bytes") or filled["bytes"]
    unsupported = list(report.get("unsupported") or [])
    if recalc.get("flag"):
        unsupported.append(recalc["flag"])

    filled = rebuild.persist_history(filled, trigger="excel_engine")
    out_key = f"{job.get('tenant', 'default')}/outputs/{filled.get('history_id')}.{(filled.get('ext') or 'xlsx')}"
    stored = excel_blob.put_bytes(out_key, filled["bytes"])

    ready = bump(
        "ready", 100, "Report ready",
        history_id=filled.get("history_id"),
        output_uri=stored.get("uri"),
        matched=filled.get("matched") or [],
        unmatched=filled.get("unmatched") or [],
        sheets=filled.get("sheets") or [],
        validation=report,
        unsupported=list(dict.fromkeys(unsupported)),
        recalculated=bool(recalc.get("recalculated")),
        rows=filled.get("rows") or 0,
        name=filled.get("name"),
        view=filled.get("view"),
        index=indexed,
        lightweight=bool(filled.get("lightweight")),
    )
    delivery.notify_ready(ready)
    excel_jobs._audit({"at": excel_jobs._now(), "type": "excel_job_ready", "job_id": job_id,
                       "history_id": filled.get("history_id"), "user": job.get("user"),
                       "tenant": job.get("tenant"), "rows": filled.get("rows"),
                       "unsupported": unsupported[:20]})
    return ready


def _run_xlsx(job_id, job, doc, model, bump):
    """Map every sheet, let the LLM plan the refresh, patch the original workbook in place."""
    from pulse import excel_blob
    from pulse import excel_jobs
    from excel_engine import engine, summary, understand

    name = doc.get("name") or "workbook.xlsx"
    try:
        built = engine.rebuild_xlsx(
            model, doc["bytes"], name, instruction=job.get("instruction") or "",
            column_map=job.get("column_map"),
            progress=lambda pct, msg: bump("rebuilding" if pct >= 50 else ("planning" if pct >= 35 else "extracting"),
                                           min(pct, 74), msg))
    except Exception as e:
        return bump("failed", 100, "Rebuild failed", error=f"{type(e).__name__}: {e}"[:400])
    wmap, plan = built["wmap"], built["plan"]
    try:
        excel_blob.put_bytes(f"{job.get('tenant', 'default')}/extracts/{job_id}.json",
                             json.dumps(understand.compact(wmap), default=str).encode("utf-8"), "application/json")
    except Exception:
        pass

    bump("validating", 76, "Validating sheets, formulas, tables and pivots")
    report = validate.validate_patch(doc["bytes"], built["bytes"], wmap, plan, built["patched"])
    if report["status"] == "fail":
        return bump("failed", 100, "Validation failed", error="; ".join(report["errors"])[:400],
                    validation=report, plan=plan, sheets=built["sheets"])

    bump("summarising", 85, "Writing the summary")
    facts = summary.facts(wmap, plan, built["patched"], built["catalog"], report)
    text = summary.write(facts)

    built = rebuild.persist_history(built, trigger="excel_engine")
    out_key = f"{job.get('tenant', 'default')}/outputs/{built.get('history_id')}.{built.get('ext') or 'xlsx'}"
    stored = excel_blob.put_bytes(out_key, built["bytes"])
    ready = bump(
        "ready", 100, "Report ready",
        history_id=built.get("history_id"), output_uri=stored.get("uri"),
        matched=built["matched"], unmatched=built["unmatched"], sheets=built["sheets"],
        validation=report, unsupported=built.get("unsupported") or [], recalculated=False,
        rows=built.get("rows") or 0, name=name, view=None, lightweight=False,
        summary=text, plan={"source": plan.get("source"), "llm_error": plan.get("llm_error"),
                            "warnings": (plan.get("warnings") or [])[:40], "sheets": plan["sheets"]},
        flow=understand.flow_text(wmap, limit=6000), facts=facts,
    )
    delivery.notify_ready(ready)
    excel_jobs._audit({"at": excel_jobs._now(), "type": "excel_job_ready", "job_id": job_id,
                       "history_id": built.get("history_id"), "user": job.get("user"),
                       "tenant": job.get("tenant"), "rows": built.get("rows"), "plan_source": plan.get("source")})
    return ready
