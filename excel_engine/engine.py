"""Excel rebuild: understand the workbook, plan the fill (LLM), patch the original in place."""
from __future__ import annotations

from excel_engine import planner, understand, xlsx_patch
from excel_engine.sources import Sources


def rebuild_xlsx(model, data: bytes, name: str, instruction: str = "", column_map=None, progress=None) -> dict:
    def bump(pct, msg):
        if progress:
            try:
                progress(pct, msg)
            except Exception:
                pass

    bump(12, "Reading every sheet: headers, formulas, tables, pivots, data flow")
    wmap = understand.workbook_map(
        data, name, on_progress=lambda i, n, s: bump(12 + int(18 * i / max(n, 1)), f"Analysing sheet {i}/{n}: {s}"))
    src = Sources(model)
    catalog = src.catalog()
    bump(35, "Planning the rebuild from the workbook flow (AI)")
    plan = planner.make_plan(wmap, catalog, instruction, column_map)
    todo = [e for e in plan["sheets"] if e["strategy"] != "keep"]
    bump(50, f"Plan ready ({plan['source']}): refreshing {len(todo)} sheet(s), keeping "
             f"{len(plan['sheets']) - len(todo)}")
    patched = xlsx_patch.patch(
        data, wmap, plan, src,
        on_progress=lambda i, n, s: bump(55 + int(20 * i / max(n, 1)), f"Refreshing sheet {i}/{n}: {s}"))
    results = {r["sheet"]: r for r in patched["sheets"]}
    by = {s["name"]: s for s in wmap["sheets"]}
    matched, unmatched, sheets = [], [], []
    for e in plan["sheets"]:
        s = by.get(e["sheet"], {})
        headers = {c["col"]: c["header"] for c in s.get("columns") or []}
        r = results.get(e["sheet"], {})
        written = [headers.get(c["col"], c["col"]) for c in e["columns"]
                   if c["action"] in ("fill", "fill_month", "fill_location")]
        left = [headers.get(c["col"], c["col"]) for c in e["columns"] if c["action"] == "leave"]
        if e["strategy"] != "keep" and r.get("status") == "refreshed":
            matched += written
            unmatched += left
        sheets.append({"name": e["sheet"], "role": "filled" if r.get("status") == "refreshed" else "kept",
                       "strategy": e["strategy"], "purpose": e.get("purpose"), "status": r.get("status") or "kept",
                       "rows": r.get("rows") or 0, "matched": written[:60], "unmatched": left[:30],
                       "reason": r.get("reason") or e.get("reason")})
    rows = sum(r.get("rows") or 0 for r in patched["sheets"] if r.get("status") == "refreshed")
    return {"bytes": patched["bytes"], "ext": "xlsm" if name.lower().endswith(".xlsm") else "xlsx", "name": name,
            "wmap": wmap, "plan": plan, "patched": patched, "catalog": catalog,
            "matched": list(dict.fromkeys(matched)), "unmatched": list(dict.fromkeys(unmatched)),
            "sheets": sheets, "rows": rows, "unsupported": wmap.get("unsupported") or []}
