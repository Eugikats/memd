"""Text extraction for uploaded office files. Numbers in the file are context only."""
import csv
import io
import uuid
from datetime import datetime, timedelta

from openpyxl import load_workbook, Workbook

_STORE = {}
_LIMIT = 12
_TEXT_CAP = 14000


def _trim(text):
    text = (text or "").strip()
    if len(text) > _TEXT_CAP:
        return text[:_TEXT_CAP] + "\n…(truncated)"
    return text


def _xlsx(data):
    book = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    chunks, headers = [], []
    for sheet in book.worksheets[:3]:
        rows = []
        for i, row in enumerate(sheet.iter_rows(values_only=True)):
            if i >= 40:
                break
            cells = ["" if c is None else str(c) for c in row[:20]]
            if any(cells):
                rows.append("\t".join(cells))
            if i == 0:
                headers = [c for c in cells if c]
        if rows:
            chunks.append(f"# {sheet.title}\n" + "\n".join(rows))
    book.close()
    return _trim("\n\n".join(chunks)), headers


def _csv(data):
    text = data.decode("utf-8-sig", errors="replace")
    rows = list(csv.reader(io.StringIO(text)))[:40]
    headers = [c for c in (rows[0] if rows else []) if c]
    body = "\n".join("\t".join(r[:20]) for r in rows)
    return _trim(body), headers


def _docx(data):
    from docx import Document
    doc = Document(io.BytesIO(data))
    lines = [p.text for p in doc.paragraphs if p.text.strip()][:200]
    return _trim("\n".join(lines)), []


def _pptx(data):
    from pptx import Presentation
    deck = Presentation(io.BytesIO(data))
    chunks = []
    for i, slide in enumerate(deck.slides, 1):
        bits = []
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False):
                bits.append(shape.text_frame.text)
        if bits:
            chunks.append(f"# Slide {i}\n" + "\n".join(bits))
        if i >= 30:
            break
    return _trim("\n\n".join(chunks)), []


def extract(name, data):
    ext = (name.rsplit(".", 1)[-1] if "." in name else "").lower()
    if ext in ("txt", "md"):
        return _trim(data.decode("utf-8", errors="replace")), []
    if ext == "csv":
        return _csv(data)
    if ext in ("xlsx", "xlsm"):
        return _xlsx(data)
    if ext == "docx":
        return _docx(data)
    if ext == "pptx":
        return _pptx(data)
    raise ValueError("Upload an Excel, CSV, text, Word, or PowerPoint file.")


def save(name, data):
    if len(data) > 8 * 1024 * 1024:
        raise ValueError("Files must be 8 MB or smaller.")
    text, headers = extract(name, data)
    if not text:
        raise ValueError("No text could be read from that file.")
    now = datetime.now()
    stale = [k for k, v in _STORE.items() if v["at"] < now - timedelta(hours=12)]
    for k in stale:
        _STORE.pop(k, None)
    while len(_STORE) >= _LIMIT:
        _STORE.pop(next(iter(_STORE)))
    doc_id = uuid.uuid4().hex[:10]
    _STORE[doc_id] = {"id": doc_id, "name": name[:180], "text": text, "headers": headers[:30], "at": now}
    return {"id": doc_id, "name": name[:180], "headers": headers[:30], "chars": len(text)}


def get(doc_id):
    return _STORE.get(doc_id)


def context_block(ids):
    parts = []
    for doc_id in ids or []:
        doc = get(doc_id)
        if not doc:
            continue
        parts.append(f"UPLOADED DOCUMENT {doc['name']} (layout and wording only; its numbers are not inventory figures):\n{doc['text']}")
    return "\n\n".join(parts)


def fill_layout(model, doc_id, view):
    """Rebuild an Excel whose columns follow the upload, filled only where a view column matches."""
    from . import views as V
    doc = get(doc_id)
    if not doc or not doc.get("headers"):
        return {"error": "That upload has no header row to follow."}
    if view not in V.VIEW_IDS:
        return {"error": "Unknown view"}
    out = V.render(model, view, {})
    table = (out.get("tables") or [None])[0]
    if not table:
        return {"error": "That view has no table."}
    by_name = {}
    for col in table["columns"]:
        by_name[col["label"].strip().lower()] = col
        by_name[col["field"].strip().lower()] = col
    used, skipped = [], []
    book = Workbook()
    sheet = book.active
    sheet.title = "Report"
    sheet.append(doc["headers"])
    for header in doc["headers"]:
        col = by_name.get(header.strip().lower())
        (used if col else skipped).append(header)
    for row in table["rows"]:
        sheet.append([row.get(by_name[h.strip().lower()]["field"]) if h.strip().lower() in by_name else None for h in doc["headers"]])
    buf = io.BytesIO()
    book.save(buf)
    return {"bytes": buf.getvalue(), "name": doc["name"], "matched": used, "unmatched": skipped,
            "view": out.get("title"), "rows": len(table["rows"])}
