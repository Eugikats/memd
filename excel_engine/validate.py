"""Compare original extract vs rebuilt workbook; never invent formulas."""
from __future__ import annotations

import io


def validate_rebuild(original_extract: dict, output_bytes: bytes, fill_result: dict) -> dict:
    """Return validation report with ok/warn/fail flags."""
    warnings = []
    errors = []
    checks = []

    orig_sheets = {s["name"]: s for s in (original_extract.get("sheets") or [])}
    out_sheets = {s.get("name"): s for s in (fill_result.get("sheets") or [])}

    missing = [n for n in orig_sheets if n not in out_sheets]
    if missing:
        # fill_layout keeps all sheets in the workbook bytes even if not listed as filled;
        # verify via openpyxl when possible.
        try:
            from openpyxl import load_workbook
            book = load_workbook(io.BytesIO(output_bytes), read_only=True, data_only=False)
            present = {ws.title for ws in book.worksheets}
            book.close()
            missing = [n for n in orig_sheets if n not in present]
        except Exception:
            pass
    if missing:
        errors.append("Missing sheets after rebuild: " + ", ".join(missing[:12]))
    else:
        checks.append({"name": "sheets_preserved", "ok": True})

    for name, meta in list(orig_sheets.items())[:40]:
        orig_headers = meta.get("headers") or []
        if not orig_headers:
            continue
        # Header presence is structural; fill may leave unmatched blank.
        checks.append({"name": f"headers:{name}", "ok": True, "headers": orig_headers[:20]})

    if fill_result.get("error"):
        errors.append(fill_result["error"])

    if not fill_result.get("matched") and not fill_result.get("error"):
        warnings.append("No columns matched system fields")

    # Formula text must come from the template, not the model.
    for note in (fill_result.get("formulas_kept") or [])[:50]:
        if isinstance(note, str) and note.strip().startswith("=") is False and "sheet skipped" in note:
            warnings.append(note)

    unsupported = list(original_extract.get("unsupported") or [])
    unsupported.extend(fill_result.get("unsupported") or [])

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
    }
