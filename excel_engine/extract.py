"""Extract workbook structure, formulas, named ranges, and unsupported-feature flags."""
from __future__ import annotations

import io
import re
import zipfile
from xml.etree import ElementTree as ET


def _is_formula(value):
    return isinstance(value, str) and value.startswith("=")


def extract_workbook(data: bytes, name: str = "workbook.xlsx") -> dict:
    """Return a structure summary suitable for search indexing and validation."""
    unsupported = []
    sheets = []
    named_ranges = []
    formulas = []
    charts = 0
    try:
        from openpyxl import load_workbook
        book = load_workbook(io.BytesIO(data), read_only=False, data_only=False, keep_links=False)
    except Exception as e:
        return _zip_extract(data, name, note=f"openpyxl failed: {type(e).__name__}")

    try:
        for defn in getattr(book, "defined_names", {}).values() if hasattr(book, "defined_names") else []:
            try:
                named_ranges.append({"name": defn.name, "attr": str(getattr(defn, "attr_text", "") or "")[:200]})
            except Exception:
                pass
        for ws in book.worksheets:
            header_row, labels = None, []
            width = min(ws.max_column or 1, 40)
            for r in range(1, min(ws.max_row or 1, 15) + 1):
                found = []
                for c in range(1, width + 1):
                    value = ws.cell(r, c).value
                    if isinstance(value, str) and value.strip() and not _is_formula(value):
                        found.append({"col": c, "header": value.strip()})
                if len(found) >= 2:
                    header_row, labels = r, found
                    break
            sheet_formulas = []
            sample_cells = 0
            for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row or 1, 80),
                                    max_col=min(ws.max_column or 1, 40)):
                for cell in row:
                    sample_cells += 1
                    if _is_formula(cell.value):
                        entry = {"sheet": ws.title, "cell": cell.coordinate, "formula": str(cell.value)[:500]}
                        sheet_formulas.append(entry)
                        formulas.append(entry)
                        if "!" in str(cell.value) and re.search(r"\[.+\]", str(cell.value)):
                            unsupported.append(f"External link formula on {ws.title}!{cell.coordinate}")
            try:
                charts += len(getattr(ws, "_charts", []) or [])
            except Exception:
                pass
            sheets.append({
                "name": ws.title,
                "header_row": header_row,
                "headers": [h["header"] for h in labels],
                "header_map": labels,
                "max_row": ws.max_row,
                "max_col": ws.max_column,
                "formula_count": len(sheet_formulas),
                "formula_samples": sheet_formulas[:20],
            })
        if charts:
            unsupported.append(f"{charts} chart(s) detected — chart recreation is best-effort / may be preserved as-is")
        # Pivot / vba hints via zip
        extras = _zip_feature_flags(data)
        unsupported.extend(extras)
    finally:
        book.close()

    return {
        "name": name,
        "sheets": sheets,
        "named_ranges": named_ranges[:100],
        "formulas": formulas[:500],
        "formula_count": len(formulas),
        "unsupported": list(dict.fromkeys(unsupported)),
        "summary": _summary_text(name, sheets, named_ranges, formulas, unsupported),
    }


def _zip_feature_flags(data: bytes) -> list:
    flags = []
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            names = zf.namelist()
            if any(n.startswith("xl/pivotCache") or n.startswith("xl/pivotTables") for n in names):
                flags.append("Pivot tables/caches present — values may need Excel to refresh")
            if "xl/vbaProject.bin" in names:
                flags.append("VBA macros present — macros are preserved but not executed by Pulse")
            if any("externalLinks" in n for n in names):
                flags.append("External workbook links present")
    except Exception:
        pass
    return flags


def _zip_extract(data: bytes, name: str, note: str = "") -> dict:
    sheets = []
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            if "xl/workbook.xml" in zf.namelist():
                root = ET.fromstring(zf.read("xl/workbook.xml"))
                for el in root.iter():
                    if el.tag.endswith("sheet") and el.attrib.get("name"):
                        sheets.append({"name": el.attrib["name"], "headers": [], "formula_count": 0,
                                       "header_row": None, "header_map": [], "max_row": None, "max_col": None,
                                       "formula_samples": []})
    except Exception:
        sheets = [{"name": "Workbook", "headers": [], "formula_count": 0, "header_row": None,
                   "header_map": [], "max_row": None, "max_col": None, "formula_samples": []}]
    unsupported = _zip_feature_flags(data)
    if note:
        unsupported.append(note)
    return {
        "name": name,
        "sheets": sheets,
        "named_ranges": [],
        "formulas": [],
        "formula_count": 0,
        "unsupported": unsupported,
        "summary": _summary_text(name, sheets, [], [], unsupported),
    }


def _summary_text(name, sheets, named_ranges, formulas, unsupported) -> str:
    lines = [f"Workbook: {name}", f"Sheets ({len(sheets)}): " + ", ".join(s["name"] for s in sheets[:30])]
    for s in sheets[:12]:
        if s.get("headers"):
            lines.append(f"  {s['name']} headers: " + ", ".join(s["headers"][:20]))
        if s.get("formula_count"):
            lines.append(f"  {s['name']} formulas: {s['formula_count']}")
    if named_ranges:
        lines.append("Named ranges: " + ", ".join(n["name"] for n in named_ranges[:20]))
    if formulas:
        lines.append(f"Indexed formulas: {min(len(formulas), 500)}")
    if unsupported:
        lines.append("Unsupported / caution: " + "; ".join(unsupported[:12]))
    return "\n".join(lines)
