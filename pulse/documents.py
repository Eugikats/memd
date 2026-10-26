"""Uploaded office files: layout, formulas, and wording. Numbers in the file are not inventory figures."""
import csv
import io
import re
import uuid
from datetime import datetime, timedelta

from openpyxl import Workbook, load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter

_STORE = {}
_LIMIT = 12
_TEXT_CAP = 14000


def _trim(text):
    text = (text or "").strip()
    if len(text) > _TEXT_CAP:
        return text[:_TEXT_CAP] + "\n…(truncated)"
    return text


def _ext(name):
    return (name.rsplit(".", 1)[-1] if "." in name else "").lower()


def _formula_note(formula):
    raw = str(formula or "")
    if not raw.startswith("="):
        return None
    body = raw[1:]
    funcs = list(dict.fromkeys(re.findall(r"\b([A-Za-z][A-Za-z0-9.]*)\s*\(", body)))
    sheets = [a or b for a, b in re.findall(r"(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_]*))!", body)]
    if sheets:
        where = "another sheet (" + ", ".join(dict.fromkeys(sheets)) + ")"
    elif re.search(r"[A-Za-z]{1,3}\d*\s*:\s*[A-Za-z]{1,3}", body):
        where = "a range on this sheet"
    else:
        where = "cells on the same row"
    if funcs:
        return f"{', '.join(funcs)} of {where}"
    return f"calculation of {where}"


def _is_formula(value):
    return isinstance(value, str) and value.startswith("=")


def _loose_formulas(ws):
    found = []
    for r in range(1, min(ws.max_row or 1, 20) + 1):
        for c in range(1, min(ws.max_column or 1, 8) + 1):
            value = ws.cell(r, c).value
            if _is_formula(value):
                found.append(value)
                if len(found) >= 4:
                    return found
    return found


def _xlsx_describe(data):
    book = load_workbook(io.BytesIO(data), data_only=False)
    chunks, headers = [], []
    for ws in book.worksheets:
        header_row, labels = _header_row(ws)
        role, inputs, formulas = _sheet_role(ws, header_row, labels)
        if not headers and inputs:
            headers = [h for _, h in inputs]
        lines = [f"# Sheet {ws.title} ({'kept as formulas' if role == 'keep' else 'data sheet, header row ' + str(header_row)})"]
        if inputs:
            lines.append("Input columns: " + ", ".join(h for _, h in inputs))
        if formulas:
            for _col, header, sample in formulas:
                note = _formula_note(sample) or "calculation"
                lines.append(f"Calculated column {header}: {note}. Example: {sample}")
        elif role == "keep":
            for sample in _loose_formulas(ws):
                lines.append(f"Calculated cell: {_formula_note(sample)}. Example: {sample}")
        if len(lines) == 1:
            lines.append("Kept as laid out.")
        chunks.append("\n".join(lines))
    book.close()
    text = _trim("\n\n".join(chunks))
    if text:
        text += "\n\nCached formula results are not inventory figures."
    return text, headers


def _header_row(ws):
    last = min(ws.max_row or 1, 15)
    width = min(ws.max_column or 1, 40)
    for r in range(1, last + 1):
        labels = []
        for c in range(1, width + 1):
            value = ws.cell(r, c).value
            if isinstance(value, str) and value.strip() and not _is_formula(value):
                labels.append((c, value.strip()))
        if len(labels) >= 2:
            return r, labels
    return None, []


def _sheet_role(ws, header_row, labels):
    if not header_row or not labels:
        return "keep", [], []
    inputs, formulas = [], []
    for col, header in labels:
        kinds, sample = [], None
        for r in range(header_row + 1, header_row + 9):
            value = ws.cell(r, col).value
            if value in (None, ""):
                continue
            if _is_formula(value):
                kinds.append("formula")
                sample = sample or value
            else:
                kinds.append("value")
        if kinds and kinds.count("formula") >= kinds.count("value") and "formula" in kinds:
            formulas.append((col, header, sample))
        else:
            inputs.append((col, header))
    cross = any("!" in str(sample or "") for _, _, sample in formulas)
    if formulas and (not inputs or (cross and len(inputs) <= 1)):
        return "keep", inputs, formulas
    return "fill", inputs, formulas


def _csv(data):
    text = data.decode("utf-8-sig", errors="replace")
    rows = list(csv.reader(io.StringIO(text)))[:40]
    headers = [c for c in (rows[0] if rows else []) if c]
    body = "Input columns: " + ", ".join(headers) if headers else ""
    return _trim(body), headers


def _docx(data):
    from docx import Document
    doc = Document(io.BytesIO(data))
    lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()][:80]
    headers = []
    for i, table in enumerate(doc.tables, 1):
        if not table.rows:
            continue
        cols = [c.text.strip() for c in table.rows[0].cells if c.text.strip()]
        if cols and not headers:
            headers = cols
        if cols:
            lines.append(f"# Table {i}\nInput columns: " + ", ".join(cols))
    return _trim("\n".join(lines)), headers


def _pptx(data):
    from pptx import Presentation
    deck = Presentation(io.BytesIO(data))
    chunks, headers = [], []
    for i, slide in enumerate(deck.slides, 1):
        bits = [f"# Slide {i}"]
        for shape in slide.shapes:
            if getattr(shape, "has_table", False):
                row = [c.text.strip() for c in shape.table.rows[0].cells if c.text.strip()] if shape.table.rows else []
                if row:
                    if not headers:
                        headers = row
                    bits.append("Input columns: " + ", ".join(row))
            elif getattr(shape, "has_text_frame", False) and shape.text_frame.text.strip():
                bits.append(shape.text_frame.text.strip())
        if len(bits) > 1:
            chunks.append("\n".join(bits))
        if i >= 30:
            break
    return _trim("\n\n".join(chunks)), headers


def extract(name, data):
    ext = _ext(name)
    if ext in ("txt", "md"):
        return _trim(data.decode("utf-8", errors="replace")), []
    if ext == "csv":
        return _csv(data)
    if ext in ("xlsx", "xlsm"):
        return _xlsx_describe(data)
    if ext == "docx":
        return _docx(data)
    if ext == "pptx":
        return _pptx(data)
    raise ValueError("Upload an Excel, CSV, text, Word, or PowerPoint file.")


def save(name, data):
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
    ext = _ext(name)
    _STORE[doc_id] = {"id": doc_id, "name": name[:180], "text": text, "headers": headers[:30],
                      "bytes": data, "ext": ext, "at": now}
    return {"id": doc_id, "name": name[:180], "headers": headers[:30], "chars": len(text)}


def get(doc_id):
    return _STORE.get(doc_id)


def context_block(ids):
    parts = []
    for doc_id in ids or []:
        doc = get(doc_id)
        if not doc:
            continue
        parts.append(
            f"UPLOADED DOCUMENT {doc['name']} (id {doc['id']}; layout, sheets, and formulas only; "
            f"its numbers are not inventory figures):\n{doc['text']}")
    return "\n\n".join(parts)


def _column_index(columns):
    by = {}
    for col in columns:
        by[str(col["label"]).strip().lower()] = col
        by[str(col["field"]).strip().lower()] = col
    return by


def _catalog(model):
    """Column labels only, so choosing a sheet's view does not export every row of every view."""
    from . import views as V
    token = V.ROW_LIMIT.set(1)
    try:
        out = {}
        for view in V.VIEW_IDS:
            rendered = V.render(model, view, {})
            table = (rendered.get("tables") or [None])[0]
            if table:
                out[view] = {"title": rendered.get("title") or view, "columns": table["columns"]}
        return out
    finally:
        V.ROW_LIMIT.reset(token)


def _materialise(model, view, cache):
    if view in cache:
        return cache[view]
    from . import views as V
    token = V.ROW_LIMIT.set(1_000_000)
    try:
        rendered = V.render(model, view, {})
        table = (rendered.get("tables") or [None])[0]
        cache[view] = {"title": rendered.get("title") or view, "table": table}
    finally:
        V.ROW_LIMIT.reset(token)
    return cache[view]


def _choose_view(catalog, headers, preferred):
    from . import views as V
    names = [h.strip().lower() for h in headers if h and str(h).strip()]
    if preferred and preferred in catalog:
        return preferred
    order = {row[1]: i for i, row in enumerate(V.VIEWS)}
    best, score = None, 0
    for view, info in catalog.items():
        keys = _column_index(info["columns"])
        sc = sum(1 for name in names if name in keys)
        if sc > score or (sc == score and sc and (best is None or order.get(view, 99) < order.get(best, 99))):
            best, score = view, sc
    return best


def _translate(formula, col, src_row, dest_row):
    if dest_row == src_row:
        return formula
    origin = f"{get_column_letter(col)}{src_row}"
    try:
        return Translator(formula, origin=origin).translate_formula(f"{get_column_letter(col)}{dest_row}")
    except Exception:
        return formula


def _write_cell(ws, row, col, value):
    try:
        ws.cell(row, col).value = value
    except AttributeError:
        pass


def _fill_xlsx(model, doc, view):
    book = load_workbook(io.BytesIO(doc["bytes"]), data_only=False)
    catalog = _catalog(model)
    loaded = {}
    if view and view not in catalog:
        book.close()
        return {"error": "Unknown view"}
    sheets, matched, unmatched, kept = [], [], [], []
    total = 0
    titles = []
    for ws in book.worksheets:
        header_row, labels = _header_row(ws)
        role, inputs, formulas = _sheet_role(ws, header_row, labels)
        formula_notes = [f"{header}: {sample}" for _, header, sample in formulas if sample]
        if role == "keep":
            sheets.append({"name": ws.title, "role": "kept", "formulas_kept": formula_notes})
            kept.extend(formula_notes)
            continue
        headers = [h for _, h in labels]
        chosen = _choose_view(catalog, headers, view)
        info = _materialise(model, chosen, loaded) if chosen else None
        by = _column_index(info["table"]["columns"]) if info and info.get("table") else {}
        used = [h for h in headers if h.strip().lower() in by]
        skipped = [h for h in headers if h.strip().lower() not in by]
        if not used:
            sheets.append({"name": ws.title, "role": "kept", "formulas_kept": formula_notes, "unmatched": skipped})
            kept.extend(formula_notes)
            continue
        rows = info["table"]["rows"]
        cols = [c for c, _h in labels]
        last = header_row
        blank = 0
        for r in range(header_row + 1, min(ws.max_row or header_row, header_row + 5000) + 1):
            if any(ws.cell(r, c).value not in (None, "") for c in cols):
                last, blank = r, 0
            else:
                blank += 1
                if blank >= 2 and last > header_row:
                    break
        samples = {}
        for col, _header, sample in formulas:
            src = header_row + 1
            for r in range(header_row + 1, max(last, header_row + 1) + 1):
                value = ws.cell(r, col).value
                if _is_formula(value):
                    src, sample = r, value
                    break
            if sample:
                samples[col] = (src, sample)
        for i, row in enumerate(rows):
            r = header_row + 1 + i
            for col, header in inputs:
                key = header.strip().lower()
                _write_cell(ws, r, col, row.get(by[key]["field"]) if key in by else None)
            for col, (src, sample) in samples.items():
                if r == src:
                    continue
                existing = ws.cell(r, col).value
                if _is_formula(existing):
                    continue
                _write_cell(ws, r, col, _translate(sample, col, src, r))
        end = header_row + len(rows)
        for r in range(end + 1, last + 1):
            for col in cols:
                _write_cell(ws, r, col, None)
        matched.extend(used)
        unmatched.extend(skipped)
        kept.extend(formula_notes)
        total += len(rows)
        titles.append(info["title"])
        sheets.append({"name": ws.title, "role": "filled", "view": info["title"], "rows": len(rows),
                       "matched": used, "unmatched": skipped, "formulas_kept": formula_notes})
    buf = io.BytesIO()
    book.save(buf)
    book.close()
    if not any(s["role"] == "filled" for s in sheets):
        return {"error": "No sheet in that file had a header matching a system column."}
    return {"bytes": buf.getvalue(), "ext": "xlsx", "name": doc["name"], "matched": list(dict.fromkeys(matched)),
            "unmatched": list(dict.fromkeys(unmatched)), "sheets": sheets, "formulas_kept": kept,
            "view": ", ".join(dict.fromkeys(titles)), "rows": total}


def _fill_docx(model, doc, view):
    from docx import Document
    catalog = _catalog(model)
    loaded = {}
    if view and view not in catalog:
        return {"error": "Unknown view"}
    document = Document(io.BytesIO(doc["bytes"]))
    sheets, matched, unmatched = [], [], []
    total, titles = 0, []
    for i, table in enumerate(document.tables, 1):
        if not table.rows:
            continue
        headers = [c.text.strip() for c in table.rows[0].cells]
        chosen = _choose_view(catalog, [h for h in headers if h], view)
        info = _materialise(model, chosen, loaded) if chosen else None
        by = _column_index(info["table"]["columns"]) if info and info.get("table") else {}
        used = [h for h in headers if h.strip().lower() in by]
        if not used:
            sheets.append({"name": f"Table {i}", "role": "kept"})
            continue
        data = info["table"]["rows"]
        while len(table.rows) - 1 < len(data):
            table.add_row()
        body = list(table.rows)[1:]
        for n, source in enumerate(data):
            for cell, header in zip(body[n].cells, headers):
                key = header.strip().lower()
                value = source.get(by[key]["field"]) if key in by else None
                cell.text = "" if value is None else str(value)
        for row in body[len(data):]:
            for cell in row.cells:
                cell.text = ""
        matched.extend(used)
        unmatched.extend(h for h in headers if h and h.strip().lower() not in by)
        total += len(data)
        titles.append(info["title"])
        sheets.append({"name": f"Table {i}", "role": "filled", "view": info["title"], "rows": len(data), "matched": used})
    if not any(s["role"] == "filled" for s in sheets):
        return {"error": "No table in that document had a header matching a system column."}
    buf = io.BytesIO()
    document.save(buf)
    return {"bytes": buf.getvalue(), "ext": "docx", "name": doc["name"], "matched": list(dict.fromkeys(matched)),
            "unmatched": list(dict.fromkeys(unmatched)), "sheets": sheets, "formulas_kept": [],
            "view": ", ".join(dict.fromkeys(titles)), "rows": total}


def _fill_pptx(model, doc, view):
    from pptx import Presentation
    catalog = _catalog(model)
    loaded = {}
    if view and view not in catalog:
        return {"error": "Unknown view"}
    deck = Presentation(io.BytesIO(doc["bytes"]))
    sheets, matched, unmatched = [], [], []
    total, titles = 0, []
    seen = 0
    for slide in deck.slides:
        for shape in slide.shapes:
            if not getattr(shape, "has_table", False) or not shape.table.rows:
                continue
            seen += 1
            table = shape.table
            headers = [c.text.strip() for c in table.rows[0].cells]
            chosen = _choose_view(catalog, [h for h in headers if h], view)
            info = _materialise(model, chosen, loaded) if chosen else None
            by = _column_index(info["table"]["columns"]) if info and info.get("table") else {}
            used = [h for h in headers if h.strip().lower() in by]
            if not used:
                sheets.append({"name": f"Slide table {seen}", "role": "kept"})
                continue
            data = info["table"]["rows"]
            capacity = max(0, len(table.rows) - 1)
            for n in range(capacity):
                source = data[n] if n < len(data) else None
                for c, header in enumerate(headers):
                    key = header.strip().lower()
                    if source is None or key not in by or source.get(by[key]["field"]) is None:
                        table.cell(n + 1, c).text = ""
                    else:
                        table.cell(n + 1, c).text = str(source.get(by[key]["field"]))
            written = min(capacity, len(data))
            matched.extend(used)
            unmatched.extend(h for h in headers if h and h.strip().lower() not in by)
            total += written
            titles.append(info["title"])
            sheets.append({"name": f"Slide table {seen}", "role": "filled", "view": info["title"],
                           "rows": written, "matched": used})
    if not any(s["role"] == "filled" for s in sheets):
        return {"error": "No table in that presentation had a header matching a system column."}
    buf = io.BytesIO()
    deck.save(buf)
    return {"bytes": buf.getvalue(), "ext": "pptx", "name": doc["name"], "matched": list(dict.fromkeys(matched)),
            "unmatched": list(dict.fromkeys(unmatched)), "sheets": sheets, "formulas_kept": [],
            "view": ", ".join(dict.fromkeys(titles)), "rows": total}


def _fill_flat(model, doc, view):
    catalog = _catalog(model)
    headers = doc.get("headers") or []
    if not headers:
        return {"error": "That upload has no header row to follow."}
    chosen = _choose_view(catalog, headers, view)
    info = _materialise(model, chosen, {}) if chosen else None
    if not info:
        return {"error": "No system column matches that header row."}
    by = _column_index(info["table"]["columns"])
    used = [h for h in headers if h.strip().lower() in by]
    skipped = [h for h in headers if h.strip().lower() not in by]
    if not used:
        return {"error": "No system column matches that header row."}
    book = Workbook()
    sheet = book.active
    sheet.title = "Report"
    sheet.append(headers)
    for row in info["table"]["rows"]:
        sheet.append([row.get(by[h.strip().lower()]["field"]) if h.strip().lower() in by else None for h in headers])
    buf = io.BytesIO()
    book.save(buf)
    return {"bytes": buf.getvalue(), "ext": "xlsx", "name": doc["name"], "matched": used, "unmatched": skipped,
            "sheets": [{"name": "Report", "role": "filled", "view": info["title"], "rows": len(info["table"]["rows"]),
                        "matched": used, "unmatched": skipped, "formulas_kept": []}],
            "formulas_kept": [], "view": info["title"], "rows": len(info["table"]["rows"])}


def fill_layout(model, doc_id, view=None):
    """Copy the uploaded file and fill matching columns from system views. Formula cells stay formulas."""
    doc = get(doc_id)
    if not doc or "bytes" not in doc:
        return {"error": "That upload is no longer available. Attach it again."}
    ext = doc.get("ext")
    if ext in ("xlsx", "xlsm"):
        return _fill_xlsx(model, doc, view)
    if ext == "docx":
        return _fill_docx(model, doc, view)
    if ext == "pptx":
        return _fill_pptx(model, doc, view)
    if ext in ("csv", "txt", "md"):
        return _fill_flat(model, doc, view)
    return {"error": "That file has no table layout to follow."}
