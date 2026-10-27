"""Uploaded office files: layout, formulas, and wording. Numbers in the file are not inventory figures."""
import csv
import io
import re
import uuid
from datetime import datetime, timedelta

from openpyxl import Workbook, load_workbook

_STORE = {}
_LIMIT = 12
_TEXT_CAP = 24000


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


def _looks_like_xlsx(data):
    return isinstance(data, (bytes, bytearray)) and len(data) >= 4 and data[:2] == b"PK"


def _xlsx_zip_fallback(data):
    """Sheet names only — used when openpyxl cannot open a complex workbook on upload."""
    import zipfile
    from xml.etree import ElementTree as ET
    if not _looks_like_xlsx(data):
        raise ValueError("That Excel file could not be opened.")
    names = []
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        if "[Content_Types].xml" not in zf.namelist():
            raise ValueError("That Excel file could not be opened.")
        if "xl/workbook.xml" in zf.namelist():
            root = ET.fromstring(zf.read("xl/workbook.xml"))
            for el in root.iter():
                if el.tag.endswith("sheet") and el.attrib.get("name"):
                    names.append(el.attrib["name"])
    chunks = [f"# Sheet {n}\nKept as laid out." for n in (names or ["Workbook"])]
    text = _trim("\n\n".join(chunks) + "\n\nCached formula results are not inventory figures.")
    return text, []


def _xlsx_describe_rows(title, matrix):
    """matrix is a list of row tuples from the first rows of a sheet."""
    header_idx, labels = None, []
    for i, row in enumerate(matrix[:15]):
        found = []
        for c, value in enumerate(row[:40], start=1):
            if isinstance(value, str) and value.strip() and not _is_formula(value):
                found.append((c, value.strip()))
        if len(found) >= 2:
            header_idx, labels = i, found
            break
    if header_idx is None:
        loose = [v for row in matrix[:20] for v in row[:8] if _is_formula(v)][:4]
        lines = [f"# Sheet {title} (kept as formulas)"]
        for sample in loose:
            lines.append(f"Calculated cell: {_formula_note(sample)}. Example: {sample}")
        if len(lines) == 1:
            lines.append("Kept as laid out.")
        return "\n".join(lines), []
    inputs, formulas = [], []
    for col, header in labels:
        kinds, sample = [], None
        for row in matrix[header_idx + 1:header_idx + 9]:
            value = row[col - 1] if col - 1 < len(row) else None
            if value in (None, ""):
                continue
            if _is_formula(value):
                kinds.append("formula")
                sample = sample or value
            else:
                kinds.append("value")
        if kinds and kinds.count("formula") >= kinds.count("value") and "formula" in kinds:
            formulas.append((header, sample))
        else:
            inputs.append(header)
    cross = any("!" in str(sample or "") for _, sample in formulas)
    role = "keep" if formulas and (not inputs or (cross and len(inputs) <= 1)) else "fill"
    lines = [f"# Sheet {title} ({'kept as formulas' if role == 'keep' else 'data sheet, header row ' + str(header_idx + 1)})"]
    if inputs:
        lines.append("Input columns: " + ", ".join(inputs))
    for header, sample in formulas:
        lines.append(f"Calculated column {header}: {_formula_note(sample) or 'calculation'}. Example: {sample}")
    if len(lines) == 1:
        lines.append("Kept as laid out.")
    return "\n".join(lines), inputs if role == "fill" else []


def _xlsx_describe(data):
    """Stream the workbook on upload so large planning files do not need a full in-memory load."""
    chunks, headers = [], []
    try:
        book = load_workbook(io.BytesIO(data), read_only=True, data_only=False, keep_links=False)
    except Exception:
        return _xlsx_zip_fallback(data)
    try:
        for ws in book.worksheets[:40]:
            try:
                matrix = []
                for i, row in enumerate(ws.iter_rows(values_only=True)):
                    if i >= 40:
                        break
                    matrix.append(tuple("" if c is None else c for c in row[:40]))
                block, sheet_headers = _xlsx_describe_rows(ws.title, matrix)
                chunks.append(block)
                if not headers and sheet_headers:
                    headers = sheet_headers
            except Exception:
                chunks.append(f"# Sheet {getattr(ws, 'title', 'Sheet')}\nKept as laid out.")
    finally:
        book.close()
    text = _trim("\n\n".join(chunks))
    if not text:
        return _xlsx_zip_fallback(data)
    return text + "\n\nCached formula results are not inventory figures.", headers


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


def _workbook_flow_fast(name, data):
    """
    Capped, read_only sheet/formula summary for chat context.
    Must stay cheap — never full openpyxl load (that 504s large planning files on upload).
    """
    unsupported = []
    try:
        import sys
        from .bc_client import ROOT
        if str(ROOT) not in sys.path:
            sys.path.insert(0, str(ROOT))
        from excel_engine import understand
        wmap = understand.workbook_map(data, name, fast=True)
        return understand.flow_text(wmap, limit=18000), wmap.get("unsupported") or []
    except Exception:
        pass
    try:
        from excel_engine.extract import _zip_feature_flags
        unsupported = _zip_feature_flags(data)
    except Exception:
        pass
    lines = [f"Workbook: {name}"]
    try:
        book = load_workbook(io.BytesIO(data), read_only=True, data_only=False, keep_links=False)
    except Exception as e:
        text, _ = _xlsx_zip_fallback(data)
        return (text or f"(Could not open workbook: {type(e).__name__})"), unsupported
    try:
        sheets = book.worksheets[:25]
        lines.append("Sheets (" + str(len(book.worksheets)) + "): " + ", ".join(ws.title for ws in sheets))
        if len(book.worksheets) > 25:
            lines.append(f"…and {len(book.worksheets) - 25} more sheets")
        for ws in sheets:
            try:
                headers, formulas = [], []
                for i, row in enumerate(ws.iter_rows(max_row=40, max_col=30, values_only=False)):
                    for cell in row:
                        value = cell.value
                        if isinstance(value, str) and value.strip() and not _is_formula(value) and i < 15:
                            if len(headers) < 20 and value.strip() not in headers:
                                headers.append(value.strip())
                        if _is_formula(value) and len(formulas) < 8:
                            formulas.append(f"{cell.coordinate}={str(value)[:120]}")
                    if i >= 39:
                        break
                bit = f"  {ws.title}"
                if headers:
                    bit += " headers: " + ", ".join(headers[:12])
                if formulas:
                    bit += " | formulas: " + "; ".join(formulas[:6])
                lines.append(bit)
            except Exception:
                lines.append(f"  {getattr(ws, 'title', 'Sheet')}: (skipped)")
    finally:
        book.close()
    if unsupported:
        lines.append("Unsupported / caution: " + "; ".join(unsupported[:8]))
    summary = "\n".join(lines)
    if len(summary) > 8000:
        summary = summary[:8000] + "\n…(workbook flow truncated)"
    return summary, unsupported


def save(name, data):
    """Attach path must stay under gateway timeouts — light describe + store bytes only."""
    from . import excel_blob
    ext = _ext(name)
    try:
        text, headers = extract(name, data)
    except Exception as e:
        if ext in ("xlsx", "xlsm") and _looks_like_xlsx(data):
            text, headers = _xlsx_zip_fallback(data)
            text = (text or "Excel workbook attached.") + f"\n(Layout summary was limited: {type(e).__name__}.)"
        else:
            raise
    if not text:
        if ext in ("xlsx", "xlsm") and _looks_like_xlsx(data):
            text, headers = _xlsx_zip_fallback(data)
        else:
            raise ValueError("No text could be read from that file.")
    now = datetime.now()
    stale = [k for k, v in _STORE.items() if v["at"] < now - timedelta(hours=12)]
    for k in stale:
        _STORE.pop(k, None)
    while len(_STORE) >= _LIMIT:
        _STORE.pop(next(iter(_STORE)))
    doc_id = uuid.uuid4().hex[:10]
    tenant = excel_blob.tenant_prefix()
    blob_key = f"{tenant}/originals/{doc_id}/{name[:120]}"
    stored = excel_blob.put_bytes(blob_key, data)
    # Keep bytes in memory for this process so ask/rebuild need not re-download immediately.
    _STORE[doc_id] = {"id": doc_id, "name": name[:180], "text": text, "headers": headers[:30],
                      "blob_key": stored["key"], "uri": stored["uri"], "backend": stored["backend"],
                      "ext": ext, "at": now, "size": len(data), "unsupported": [],
                      "has_flow": False, "bytes": data}
    return {"id": doc_id, "name": name[:180], "headers": headers[:30], "chars": len(text),
            "storage": stored["backend"], "size": len(data), "has_flow": False}


def cache_bytes(doc_id, data):
    doc = _STORE.get(doc_id)
    if doc is not None:
        doc["bytes"] = data


def get(doc_id):
    doc = _STORE.get(doc_id)
    if not doc:
        return None
    if doc.get("bytes"):
        return doc
    key = doc.get("blob_key") or doc.get("uri")
    if not key:
        return doc
    try:
        from . import excel_blob
        doc["bytes"] = excel_blob.get_bytes(key)
    except Exception:
        pass
    return doc


def context_block(ids):
    parts = []
    for doc_id in ids or []:
        doc = get(doc_id)
        if not doc:
            continue
        # Lazy, capped flow for the agent — never blocks the upload HTTP request.
        if doc.get("ext") in ("xlsx", "xlsm") and not doc.get("has_flow") and doc.get("bytes"):
            try:
                flow, unsupported = _workbook_flow_fast(doc.get("name") or "workbook.xlsx", doc["bytes"])
                if flow:
                    doc["text"] = _trim((doc.get("text") or "") + "\n\n# Workbook flow (sheets, formulas, structure)\n" + flow)
                    doc["has_flow"] = True
                    doc["unsupported"] = (unsupported or [])[:20]
            except Exception as e:
                doc["has_flow"] = True  # do not retry every ask
                doc["text"] = _trim((doc.get("text") or "") + f"\n\n(Workbook flow skipped: {type(e).__name__}.)")
        parts.append(
            f"UPLOADED DOCUMENT {doc['name']} (id {doc['id']}; layout, sheets, formulas, and workbook flow only; "
            f"its numbers are not inventory figures):\n{doc['text']}")
        if doc.get("unsupported"):
            parts.append("Unsupported / caution flags: " + "; ".join(doc["unsupported"][:12]))
    return "\n\n".join(parts)


def _norm_header(text):
    s = str(text or "").strip().lower()
    s = re.sub(r"[_\-/]+", " ", s)
    s = re.sub(r"[^\w\s.]+", "", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


# Template header phrases → conceptual key; resolved to a view field that exists on the chosen view.
_HEADER_ALIASES = {
    "item": "item_no", "item code": "item_no", "item no": "item_no", "item no.": "item_no",
    "item number": "item_no", "itemnumber": "item_no", "sku": "item_no", "items": "item_no",
    "stock": "stock_qty", "inventory": "stock_qty", "on hand": "stock_qty", "stock on hand": "stock_qty",
    "qty on hand": "stock_qty", "quantity on hand": "stock_qty", "inventory qty": "stock_qty",
    "cover days": "cover", "coverdays": "cover", "days of cover": "cover", "cover day": "cover",
    "production days cover": "cover", "prod days cover": "cover", "coverdays + intransit": "cover",
    "status": "status", "cover status": "cover_status",
    "description": "description", "desc": "description",
    "in transit qty": "incoming_qty", "intransit qty": "incoming_qty", "incoming": "incoming_qty",
    "incoming po": "incoming_qty", "on order": "incoming_qty",
    "average cons": "daily_usage", "avg cons": "daily_usage", "max cons": "daily_usage",
    "daily usage": "daily_usage", "usage / day": "daily_usage", "used / day": "daily_usage",
}


def _column_index(columns):
    """Exact label/field keys plus alias keys that resolve to fields present on this view."""
    by = {}
    fields = {str(c["field"]): c for c in columns}
    for col in columns:
        by[str(col["label"]).strip().lower()] = col
        by[str(col["field"]).strip().lower()] = col
        by[_norm_header(col["label"])] = col
        by[_norm_header(col["field"])] = col
    # Conceptual cover → whichever cover field this view exposes
    cover_col = fields.get("cover_days") or fields.get("cover_prod_days")
    status_col = fields.get("cover_status") or fields.get("supply_status") or fields.get("risk_level")
    for alias, concept in _HEADER_ALIASES.items():
        if concept == "cover" and cover_col:
            by[alias] = cover_col
        elif concept == "status" and status_col and alias in ("status",):
            # Prefer explicit Status label if present, else cover/supply status
            if "status" in by and by["status"].get("field") == "status":
                continue
            by[alias] = fields.get("status") or status_col
        elif concept == "cover_status" and fields.get("cover_status"):
            by[alias] = fields["cover_status"]
        elif concept in fields:
            by[alias] = fields[concept]
    return by


def _resolve_header(header, by, column_map=None):
    """Return view column dict for a template header, or None."""
    raw = str(header or "").strip()
    if not raw:
        return None
    if column_map:
        mapped = column_map.get(raw) or column_map.get(raw.lower()) or column_map.get(_norm_header(raw))
        if mapped:
            # column_map values are field names
            for col in by.values():
                if col.get("field") == mapped:
                    return col
            # direct field key in index
            if str(mapped).lower() in by:
                return by[str(mapped).lower()]
    key = raw.lower()
    if key in by:
        return by[key]
    norm = _norm_header(raw)
    if norm in by:
        return by[norm]
    return None


def _match_score(headers, columns):
    by = _column_index(columns)
    return sum(1 for h in headers if h and _resolve_header(h, by))


_CATALOG_CACHE = None
_FILL_ROW_CAP = 2000


def _catalog(model):
    """Column labels only, cached for the process so each rebuild does not re-render every view."""
    global _CATALOG_CACHE
    if _CATALOG_CACHE is not None:
        return _CATALOG_CACHE
    from . import views as V
    token = V.ROW_LIMIT.set(1)
    try:
        out = {}
        for view in V.VIEW_IDS:
            rendered = V.render(model, view, {})
            table = (rendered.get("tables") or [None])[0]
            if table:
                out[view] = {"title": rendered.get("title") or view, "columns": table["columns"]}
        _CATALOG_CACHE = out
        return out
    finally:
        V.ROW_LIMIT.reset(token)


def _materialise(model, view, cache):
    if view in cache:
        return cache[view]
    from . import views as V
    token = V.ROW_LIMIT.set(_FILL_ROW_CAP)
    try:
        rendered = V.render(model, view, {})
        table = (rendered.get("tables") or [None])[0]
        cache[view] = {"title": rendered.get("title") or view, "table": table}
    finally:
        V.ROW_LIMIT.reset(token)
    return cache[view]


def _hint_view(doc, headers):
    """Prefer RM / inventory views when the file name (or clear header wording) points at that pack.

    Do not treat a lone 'Cover days' header as an RM hint — that label exists on inventory too.
    """
    name = str(doc.get("name") or "").lower()
    hdr = " ".join(str(h) for h in (headers or [])).lower()
    blob = f"{name} {hdr}"
    if re.search(r"\brm\b|raw.?material", blob) or re.search(r"planning", name):
        return "rm_consumption_cover"
    if re.search(r"inventory|stock\b", blob):
        return "inventory_position"
    return None


def _choose_view(catalog, headers, preferred, doc=None):
    """Pick the view with the best alias-aware header score; RM/inventory hint is a tie-break only."""
    from . import views as V
    names = [h for h in headers if h and str(h).strip()]
    if preferred and preferred in catalog:
        return preferred
    hinted = _hint_view(doc or {}, headers)
    order = {row[1]: i for i, row in enumerate(V.VIEWS)}
    best, score = None, 0
    for view, info in catalog.items():
        sc = _match_score(names, info["columns"])
        if sc < 1:
            continue
        if sc > score:
            best, score = view, sc
            continue
        if sc < score:
            continue
        # Tie: prefer filename/header hint, else catalogue order
        if hinted and view == hinted:
            best = view
        elif hinted and best == hinted:
            pass
        elif best is None or order.get(view, 99) < order.get(best, 99):
            best = view
    return best


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
        chosen = _choose_view(catalog, [h for h in headers if h], view, doc)
        info = _materialise(model, chosen, loaded) if chosen else None
        by = _column_index(info["table"]["columns"]) if info and info.get("table") else {}
        used = [h for h in headers if _resolve_header(h, by)]
        if not used:
            sheets.append({"name": f"Table {i}", "role": "kept"})
            continue
        data = (info["table"]["rows"] or [])[:_FILL_ROW_CAP]
        while len(table.rows) - 1 < len(data):
            table.add_row()
        body = list(table.rows)[1:]
        for n, source in enumerate(data):
            for cell, header in zip(body[n].cells, headers):
                resolved = _resolve_header(header, by)
                value = source.get(resolved["field"]) if resolved else None
                cell.text = "" if value is None else str(value)
        for row in body[len(data):]:
            for cell in row.cells:
                cell.text = ""
        matched.extend(used)
        unmatched.extend(h for h in headers if h and not _resolve_header(h, by))
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
            chosen = _choose_view(catalog, [h for h in headers if h], view, doc)
            info = _materialise(model, chosen, loaded) if chosen else None
            by = _column_index(info["table"]["columns"]) if info and info.get("table") else {}
            used = [h for h in headers if _resolve_header(h, by)]
            if not used:
                sheets.append({"name": f"Slide table {seen}", "role": "kept"})
                continue
            data = (info["table"]["rows"] or [])[:_FILL_ROW_CAP]
            capacity = max(0, len(table.rows) - 1)
            for n in range(capacity):
                source = data[n] if n < len(data) else None
                for c, header in enumerate(headers):
                    resolved = _resolve_header(header, by)
                    if source is None or not resolved or source.get(resolved["field"]) is None:
                        table.cell(n + 1, c).text = ""
                    else:
                        table.cell(n + 1, c).text = str(source.get(resolved["field"]))
            written = min(capacity, len(data))
            matched.extend(used)
            unmatched.extend(h for h in headers if h and not _resolve_header(h, by))
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
    chosen = _choose_view(catalog, headers, view, doc)
    info = _materialise(model, chosen, {}) if chosen else None
    if not info:
        return {"error": "No system column matches that header row."}
    by = _column_index(info["table"]["columns"])
    used = [h for h in headers if _resolve_header(h, by)]
    skipped = [h for h in headers if h and not _resolve_header(h, by)]
    if not used:
        return {"error": "No system column matches that header row."}
    rows = (info["table"]["rows"] or [])[:_FILL_ROW_CAP]
    book = Workbook()
    sheet = book.active
    sheet.title = "Report"
    sheet.append(headers)
    for row in rows:
        sheet.append([
            (lambda r: row.get(r["field"]) if r else None)(_resolve_header(h, by))
            for h in headers
        ])
    buf = io.BytesIO()
    book.save(buf)
    return {"bytes": buf.getvalue(), "ext": "xlsx", "name": doc["name"], "matched": used, "unmatched": skipped,
            "sheets": [{"name": "Report", "role": "filled", "view": info["title"], "rows": len(rows),
                        "matched": used, "unmatched": skipped, "formulas_kept": []}],
            "formulas_kept": [], "view": info["title"], "rows": len(rows)}


def fill_layout(model, doc_id, view=None, on_progress=None, column_map=None):
    """Copy the uploaded file and fill matching columns from system views. Formula cells stay formulas."""
    doc = get(doc_id)
    if not doc or "bytes" not in doc:
        return {"error": "That upload is no longer available. Attach it again."}
    cmap = column_map if isinstance(column_map, dict) else None
    ext = doc.get("ext")
    if ext in ("xlsx", "xlsm"):
        from excel_engine import engine
        try:
            built = engine.rebuild_xlsx(model, doc["bytes"], doc["name"], column_map=cmap,
                                        progress=(lambda pct, msg: on_progress(pct, 100, msg)) if on_progress else None)
        except Exception as e:
            return {"error": f"That workbook could not be rebuilt ({type(e).__name__}: {e})"[:300]}
        return {k: built[k] for k in ("bytes", "ext", "name", "matched", "unmatched", "sheets", "rows", "unsupported")}
    if ext == "docx":
        return _fill_docx(model, doc, view)
    if ext == "pptx":
        return _fill_pptx(model, doc, view)
    if ext in ("csv", "txt", "md"):
        return _fill_flat(model, doc, view)
    return {"error": "That file has no table layout to follow."}
