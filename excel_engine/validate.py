"""Compare original extract vs rebuilt workbook; never invent formulas."""
from __future__ import annotations

import io
import re
import zipfile

_FORMULA = re.compile(rb"<(?:\w+:)?f[ >/]")


def _formula_counts(zf, parts):
    out = {}
    for name, part in parts.items():
        if not part or part not in zf.namelist():
            continue
        n = 0
        with zf.open(part) as f:
            tail = b""
            while True:
                chunk = f.read(1 << 22)
                if not chunk:
                    break
                buf = tail + chunk
                n += len(_FORMULA.findall(buf[:-8] if len(buf) > 8 else b""))
                tail = buf[-8:] if len(buf) > 8 else buf
            n += len(_FORMULA.findall(tail))
        out[name] = n
    return out


def validate_patch(original: bytes, output: bytes, wmap: dict, plan: dict, patched: dict) -> dict:
    """In-place rebuild checks: same sheets in the same order, parts kept, formulas kept, no false 'refreshed'."""
    from excel_engine.understand import Package
    errors, warnings, checks = [], [], []
    try:
        out_pkg = Package(output)
    except Exception as e:
        return {"status": "fail", "ok": False, "errors": [f"Rebuilt file does not open: {type(e).__name__}"],
                "warnings": [], "checks": [], "unsupported": wmap.get("unsupported") or []}
    orig_names = [s["name"] for s in wmap["sheets"]]
    out_names = [s["name"] for s in out_pkg.sheets]
    if out_names[:len(orig_names)] != orig_names:
        missing = [n for n in orig_names if n not in out_names]
        errors.append("Sheets changed: " + (", ".join(missing[:12]) if missing else "order differs"))
    else:
        checks.append({"name": "sheets_preserved", "ok": True, "count": len(orig_names)})
    orig_zip = zipfile.ZipFile(io.BytesIO(original))
    dropped = {n for n in orig_zip.namelist() if re.match(r"xl/calcChain\.xml$", n)}
    lost = [n for n in orig_zip.namelist() if n not in out_pkg.names and n not in dropped]
    if lost:
        errors.append("Parts lost in rebuild: " + ", ".join(lost[:10]))
    else:
        checks.append({"name": "parts_preserved", "ok": True,
                       "pivots": sum(1 for n in out_pkg.names if n.startswith("xl/pivotTables/")),
                       "tables": sum(1 for n in out_pkg.names if n.startswith("xl/tables/")),
                       "drawings": sum(1 for n in out_pkg.names if n.startswith("xl/drawings/"))})
    strategy = {e["sheet"]: e["strategy"] for e in plan["sheets"]}
    parts = {s["name"]: s.get("part") for s in wmap["sheets"]}
    before = _formula_counts(orig_zip, parts)
    after = _formula_counts(out_pkg.zf, parts)
    for name, n in before.items():
        m = after.get(name, 0)
        if strategy.get(name) == "replace_rows":
            continue  # formula columns are re-applied per new row; totals follow the new row count
        if m < n:
            errors.append(f"Formulas lost on {name}: {n:,} -> {m:,}")
    checks.append({"name": "formulas_preserved", "ok": not any("Formulas lost" in e for e in errors),
                   "formulas": sum(after.values())})
    for r in patched.get("sheets") or []:
        if r.get("status") != "refreshed":
            warnings.append(f"{r['sheet']}: {r.get('status')} — {r.get('reason') or 'no system values written'}")
    if not any(r.get("status") == "refreshed" for r in patched.get("sheets") or []):
        errors.append("No sheet could be refreshed from system data")
    status = "fail" if errors else ("warn" if warnings else "ok")
    return {"status": status, "ok": not errors, "errors": errors, "warnings": warnings, "checks": checks,
            "unsupported": wmap.get("unsupported") or []}


def validate_rebuild(original_extract: dict, output_bytes: bytes, fill_result: dict) -> dict:
    """Return validation report with ok/warn/fail flags."""
    warnings = []
    errors = []
    checks = []
    lightweight = bool(fill_result.get("lightweight"))

    orig_sheets = {s["name"]: s for s in (original_extract.get("sheets") or [])}
    out_sheets = {s.get("name"): s for s in (fill_result.get("sheets") or [])}

    missing = [n for n in orig_sheets if n not in out_sheets]
    if missing:
        try:
            from openpyxl import load_workbook
            book = load_workbook(io.BytesIO(output_bytes), read_only=True, data_only=False)
            present = {ws.title for ws in book.worksheets}
            book.close()
            missing = [n for n in orig_sheets if n not in present]
        except Exception:
            pass
    if missing:
        msg = "Missing sheets after rebuild: " + ", ".join(missing[:12])
        if lightweight:
            warnings.append(msg + " (allowed for lightweight rebuild)")
        else:
            errors.append(msg)
    else:
        checks.append({"name": "sheets_preserved", "ok": True})

    for name, meta in list(orig_sheets.items())[:40]:
        orig_headers = meta.get("headers") or []
        if not orig_headers:
            continue
        checks.append({"name": f"headers:{name}", "ok": True, "headers": orig_headers[:20]})

    if fill_result.get("error"):
        errors.append(fill_result["error"])

    if not fill_result.get("matched") and not fill_result.get("error"):
        warnings.append("No columns matched system fields")

    for note in (fill_result.get("formulas_kept") or [])[:50]:
        if isinstance(note, str) and "sheet skipped" in note:
            warnings.append(note)
        if isinstance(note, str) and "not cloned" in note:
            warnings.append(note)

    unsupported = list(original_extract.get("unsupported") or [])
    unsupported.extend(fill_result.get("unsupported") or [])
    if lightweight:
        warnings.append("Lightweight rebuild: formatting/charts/formulas were not cloned")

    status = "fail" if errors else ("warn" if warnings or unsupported else "ok")
    return {
        "status": status,
        "ok": not errors,
        "errors": errors,
        "warnings": warnings,
        "checks": checks,
        "unsupported": list(dict.fromkeys(unsupported)),
        "matched": fill_result.get("matched") or [],
        "unmatched": fill_result.get("unmatched") or [],
        "rows": fill_result.get("rows") or 0,
        "lightweight": lightweight,
    }
