"""Delivery hooks for ready Excel jobs: Pulse mail now, Logic Apps / Functions webhook optional."""
from __future__ import annotations

import json
import os
import urllib.request


def logic_apps_configured() -> bool:
    return bool(os.environ.get("PULSE_EXCEL_LOGIC_APP_URL"))


def notify_ready(job: dict) -> dict:
    """POST job summary to Logic App / Function webhook when configured."""
    url = os.environ.get("PULSE_EXCEL_LOGIC_APP_URL")
    if not url:
        return {"delivered": False, "backend": "none"}
    payload = {
        "job_id": job.get("id"),
        "status": job.get("status"),
        "document_id": job.get("document_id"),
        "history_id": job.get("history_id"),
        "user": job.get("user"),
        "tenant": job.get("tenant"),
        "download_path": f"/api/reports/history/{job.get('history_id')}/download" if job.get("history_id") else None,
        "matched": job.get("matched"),
        "unmatched": job.get("unmatched"),
        "unsupported": job.get("unsupported"),
    }
    try:
        req = urllib.request.Request(url, data=json.dumps(payload).encode(),
                                     headers={"Content-Type": "application/json"}, method="POST")
        with urllib.request.urlopen(req, timeout=30) as resp:
            return {"delivered": True, "backend": "logic_app", "status": resp.status}
    except Exception as e:
        return {"delivered": False, "backend": "logic_app", "error": f"{type(e).__name__}: {e}"[:200]}


def email_history(history_id: str, recipients: list, subject: str | None = None) -> dict:
    """Send a completed rebuild via existing Pulse Graph mail."""
    from pulse import agent as A
    return A.send_report(None, history_id=history_id, recipients=recipients)