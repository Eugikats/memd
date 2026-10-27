"""Async Excel rebuild jobs: Service Bus when configured, else in-process worker thread."""
from __future__ import annotations

import json
import os
import threading
import time
import uuid
from datetime import datetime, timezone
from pathlib import Path

from .bc_client import DATA_DIR
from . import excel_blob

_JOBS: dict[str, dict] = {}
_LOCK = threading.Lock()
_JOB_DIR = Path(os.environ.get("PULSE_EXCEL_JOB_DIR") or (DATA_DIR / "excel" / "jobs"))
_AUDIT = DATA_DIR / "excel_audit.jsonl"
_SB_CONN = os.environ.get("PULSE_EXCEL_SERVICEBUS") or os.environ.get("AZURE_SERVICEBUS_CONNECTION_STRING")
_SB_QUEUE = os.environ.get("PULSE_EXCEL_SERVICEBUS_QUEUE", "excel-jobs")
_MODEL_GETTER = None  # set by app: lambda: STATE["model"]


STATES = ("queued", "extracting", "indexing", "rebuilding", "validating", "ready", "failed")


def set_model_getter(fn):
    global _MODEL_GETTER
    _MODEL_GETTER = fn


def servicebus_configured() -> bool:
    return bool(_SB_CONN)


def _now():
    return datetime.now(timezone.utc).replace(tzinfo=None).isoformat(timespec="seconds")


def _persist(job: dict):
    _JOB_DIR.mkdir(parents=True, exist_ok=True)
    (_JOB_DIR / f"{job['id']}.json").write_text(json.dumps(job, default=str), encoding="utf-8")


def _load_disk(job_id: str) -> dict | None:
    path = _JOB_DIR / f"{job_id}.json"
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return None


def _audit(event: dict):
    _AUDIT.parent.mkdir(parents=True, exist_ok=True)
    with _AUDIT.open("a", encoding="utf-8") as f:
        f.write(json.dumps(event, default=str) + "\n")


def get(job_id: str) -> dict | None:
    with _LOCK:
        job = _JOBS.get(job_id)
    if job:
        return dict(job)
    return _load_disk(job_id)


def update(job_id: str, **fields) -> dict | None:
    with _LOCK:
        job = _JOBS.get(job_id) or _load_disk(job_id)
        if not job:
            return None
        job.update(fields)
        job["updated_at"] = _now()
        _JOBS[job_id] = job
        _persist(job)
        return dict(job)


def create(document_id: str, instruction: str, user: str = "demo", view: str | None = None,
           tenant: str | None = None) -> dict:
    job_id = uuid.uuid4().hex[:12]
    tenant = excel_blob.tenant_prefix(tenant)
    job = {
        "id": job_id,
        "document_id": document_id,
        "instruction": (instruction or "")[:2000],
        "view": view,
        "user": user,
        "tenant": tenant,
        "status": "queued",
        "progress": 0,
        "message": "Queued for Excel Intelligence Engine",
        "created_at": _now(),
        "updated_at": _now(),
        "extract": None,
        "validation": None,
        "unsupported": [],
        "matched": [],
        "unmatched": [],
        "sheets": [],
        "history_id": None,
        "output_uri": None,
        "error": None,
        "recalculated": False,
    }
    with _LOCK:
        _JOBS[job_id] = job
        _persist(job)
    _audit({"at": _now(), "type": "excel_job_created", "job_id": job_id, "user": user,
            "document_id": document_id, "tenant": tenant})
    _enqueue(job_id)
    return dict(job)


def _enqueue(job_id: str):
    if servicebus_configured():
        try:
            from azure.servicebus import ServiceBusClient, ServiceBusMessage
            with ServiceBusClient.from_connection_string(_SB_CONN) as client:
                with client.get_queue_sender(_SB_QUEUE) as sender:
                    sender.send_messages(ServiceBusMessage(json.dumps({"job_id": job_id})))
            return
        except Exception as e:
            update(job_id, message=f"Service Bus unavailable ({type(e).__name__}); running in-process")
    threading.Thread(target=_run_local, args=(job_id,), daemon=True, name=f"excel-job-{job_id}").start()


def _run_local(job_id: str):
    # Small delay so create() can return to the client first.
    time.sleep(0.05)
    try:
        import sys
        from .bc_client import ROOT
        if str(ROOT) not in sys.path:
            sys.path.insert(0, str(ROOT))
        # Wait briefly for the supply-chain model if the app is still booting.
        model = None
        for _ in range(30):
            model = _MODEL_GETTER() if _MODEL_GETTER else None
            if model is not None:
                break
            time.sleep(1)
        from excel_engine.pipeline import run_job
        # If still None, pipeline will attempt model.load() itself.
        run_job(job_id, model=model)
    except Exception as e:
        update(job_id, status="failed", progress=100, error=f"{type(e).__name__}: {e}"[:400],
               message="Excel job failed")
        _audit({"at": _now(), "type": "excel_job_failed", "job_id": job_id, "error": str(e)[:400]})


def list_recent(limit: int = 50) -> list:
    _JOB_DIR.mkdir(parents=True, exist_ok=True)
    rows = []
    for path in sorted(_JOB_DIR.glob("*.json"), key=lambda p: p.stat().st_mtime, reverse=True)[:limit]:
        try:
            rows.append(json.loads(path.read_text(encoding="utf-8")))
        except json.JSONDecodeError:
            pass
    return rows


def audit_rows(limit: int = 200) -> list:
    if not _AUDIT.exists():
        return []
    out = []
    for line in _AUDIT.read_text(encoding="utf-8").splitlines()[-limit:]:
        try:
            out.append(json.loads(line))
        except json.JSONDecodeError:
            pass
    return list(reversed(out))
