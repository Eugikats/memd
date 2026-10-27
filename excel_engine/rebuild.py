"""Deterministic template fill from Pulse views — wraps pulse.documents.fill_layout."""
from __future__ import annotations


def rebuild_from_document(model, document_id: str, view: str | None = None, on_progress=None) -> dict:
    from pulse import documents
    return documents.fill_layout(model, document_id, view, on_progress=on_progress)


def persist_history(built: dict, trigger: str = "excel_engine") -> dict:
    """Write rebuilt bytes into report history; return history fields."""
    import uuid
    from pulse import reports as R

    if built.get("error") or not built.get("bytes"):
        return built
    hid = uuid.uuid4().hex[:10]
    ext = built.get("ext") or "xlsx"
    R.FILES.mkdir(parents=True, exist_ok=True)
    (R.FILES / f"{hid}.{ext}").write_bytes(built["bytes"])
    entry = {
        "id": hid,
        "report_id": "upload",
        "name": built.get("name") or "Upload layout",
        "ext": ext,
        "params": {},
        "generated_at": R.now().isoformat(timespec="seconds"),
        "by": "Excel Intelligence Engine",
        "trigger": trigger,
        "rows": built.get("rows") or 0,
        "data_loaded": "",
    }
    hist = R._read(R.HISTORY, [])
    R._write(R.HISTORY, ([entry] + hist)[:500])
    built = dict(built)
    built["history_id"] = hid
    return built
