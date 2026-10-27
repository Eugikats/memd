"""End-to-end Excel job: extract → index → rebuild → validate → optional Graph recalc → deliver."""
from __future__ import annotations

import json
from pathlib import Path

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

    bump("extracting", 10, "Extracting worksheets, formulas, and layout")
    extracted = extract.extract_workbook(doc["bytes"], doc.get("name") or "workbook.xlsx")
    # Store extract artifact
    extract_key = f"{job.get('tenant', 'default')}/extracts/{job_id}.json"
    excel_blob.put_bytes(extract_key, json.dumps(extracted).encode("utf-8"), "application/json")
    bump("extracting", 25, "Extraction complete", extract=extracted.get("summary"),
         unsupported=extracted.get("unsupported") or [])

    bump("indexing", 35, "Indexing workbook intelligence")
    indexed = search_index.index_extract(job_id, job["document_id"], job.get("tenant") or "default", extracted)
    bump("indexing", 45, f"Indexed {indexed.get('indexed', 0)} chunks ({indexed.get('backend')})")

    if model is None:
        try:
            from pulse import model as M
            model = M.load()
        except Exception as e:
            return bump("failed", 100, "Model unavailable", error=f"Could not load supply chain model: {e}"[:300])

    bump("rebuilding", 55, "Rebuilding template with system figures (deterministic)")
    filled = rebuild.rebuild_from_document(model, job["document_id"], job.get("view"))
    if filled.get("error"):
        return bump("failed", 100, "Rebuild failed", error=filled["error"],
                    unmatched=filled.get("unmatched") or [], sheets=filled.get("sheets") or [])

    bump("validating", 75, "Validating structure and formulas")
    report = validate.validate_rebuild(extracted, filled["bytes"], filled)
    if report["status"] == "fail":
        return bump("failed", 100, "Validation failed", error="; ".join(report["errors"])[:400],
                    validation=report, unsupported=report.get("unsupported") or [])

    # Optional Graph Excel recalc
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
    )
    delivery.notify_ready(ready)
    excel_jobs._audit({"at": excel_jobs._now(), "type": "excel_job_ready", "job_id": job_id,
                       "history_id": filled.get("history_id"), "user": job.get("user"),
                       "tenant": job.get("tenant"), "rows": filled.get("rows"),
                       "unsupported": unsupported[:20]})
    return ready
