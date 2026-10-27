"""Chat reply for a finished rebuild: what each sheet is, what was refreshed, what was kept and why."""
from __future__ import annotations

import json
import re

from excel_engine import llm

SYSTEM = """You write the chat reply of Pulse, a supply-chain assistant, after it rebuilt a user's Excel workbook
from Business Central data. You receive FACTS (JSON) about what was done. Write a short, specific markdown reply:
- First sentence: what was produced (same workbook, all sheets kept, which source sheets were refreshed).
- One bullet per refreshed sheet: what the sheet is, which columns now hold system figures, rows/items matched,
  and for ledger dumps the entry types and date range.
- One line naming the calculated sheets, pivots and charts that update from those sources when the file is opened in
  Excel (pivots refresh and formulas recalculate on open).
- One line on what was kept as the user's input and why (forecasts, trackers, manual comments), and anything that
  could not be refreshed with the reason. Mention the "Not in template" sheet if present.
Use only numbers that appear in FACTS. No preamble, no apology, under 220 words. Do not add a download link."""


def _downstream(wmap, names):
    feeds = {}
    for s in wmap["sheets"]:
        feeds[s["name"]] = [re.sub(r" \(pivot .*\)$", "", f) for f in s.get("feeds") or []]
    seen, todo = set(), list(names)
    while todo:
        n = todo.pop()
        for f in feeds.get(n, []):
            if f not in seen and f not in names:
                seen.add(f)
                todo.append(f)
    order = {s["name"]: i for i, s in enumerate(wmap["sheets"])}
    return sorted(seen, key=lambda x: order.get(x, 999))


def facts(wmap, plan, patched, catalog, validation=None):
    by = {s["name"]: s for s in wmap["sheets"]}
    results = {r["sheet"]: r for r in patched.get("sheets") or []}
    refreshed, kept, problems = [], [], []
    for e in plan["sheets"]:
        s = by.get(e["sheet"], {})
        headers = {c["col"]: c["header"] for c in s.get("columns") or []}
        if e["strategy"] == "keep":
            kept.append({"sheet": e["sheet"], "purpose": e.get("purpose"), "reason": e.get("reason"),
                         "hidden": s.get("state") != "visible"})
            continue
        r = results.get(e["sheet"], {})
        cols = []
        for c in e["columns"]:
            if c["action"] == "fill":
                cols.append(f"{headers.get(c['col'], c['col'])} <- {c['field']}")
            elif c["action"] == "fill_location":
                cols.append(f"{headers.get(c['col'], c['col'])} <- stock at {c['location'] or '(no location)'}")
            elif c["action"] == "fill_month":
                cols.append(f"{headers.get(c['col'], c['col'])} <- {c['measure']} {c['month']}")
        left = [f"{headers.get(c['col'], c['col'])}: {c['reason']}" for c in e["columns"]
                if c["action"] == "leave" and c.get("reason")]
        item = {"sheet": e["sheet"], "purpose": e.get("purpose"), "strategy": e["strategy"],
                "status": r.get("status"), "columns": cols[:60], "left": left[:20],
                "formula_columns_kept": sum(1 for c in e["columns"] if c["action"] == "keep_formula")}
        for k in ("rows", "matched", "unknown", "cells_written", "previous_rows", "first_date", "last_date",
                  "entry_types", "item_prefixes", "reason", "unknown_sample"):
            if r.get(k) not in (None, [], ""):
                item[k] = r[k]
        (refreshed if r.get("status") == "refreshed" else problems).append(item)
    return {
        "workbook": wmap["name"], "sheets_total": len(wmap["sheets"]), "plan_source": plan.get("source"),
        "reporting_date": catalog.get("reporting_date"),
        "history": {"first": catalog.get("history_first_date"), "last": catalog.get("history_last_date")},
        "refreshed": refreshed, "not_refreshed": problems, "kept": kept,
        "recalculates_on_open": _downstream(wmap, [r["sheet"] for r in refreshed]),
        "missing_items_sheet": patched.get("missing_sheet"), "missing_items": patched.get("missing_items", 0),
        "validation": {k: (validation or {}).get(k) for k in ("status", "errors", "warnings")} if validation else None,
        "flags": wmap.get("unsupported") or [],
    }


def write(f: dict) -> str:
    if llm.available():
        try:
            text = llm.chat_text(SYSTEM, "FACTS:\n" + json.dumps(f, default=str)[:30000])
            if text:
                return text
        except Exception:
            pass
    return fallback(f)


def fallback(f: dict) -> str:
    ref = f["refreshed"]
    lines = []
    if ref:
        lines.append(f"I rebuilt **{f['workbook']}** with all {f['sheets_total']} sheets kept and refreshed "
                     f"{len(ref)} source sheet(s) from Business Central (data as of {f['reporting_date']}).")
    else:
        lines.append(f"I could not refresh any sheet of **{f['workbook']}** from system data; the file is unchanged.")
    for r in ref:
        bit = f"- **{r['sheet']}**"
        if r.get("purpose"):
            bit += f" ({r['purpose']})"
        if r["strategy"] == "replace_rows":
            bit += (f": {r.get('rows', 0):,} ledger rows"
                    + (f" from {r.get('first_date')} to {r.get('last_date')}" if r.get("first_date") else "")
                    + (f", entry types {', '.join(r.get('entry_types') or [])}" if r.get("entry_types") else "")
                    + f" (template had {r.get('previous_rows', 0):,})")
        else:
            bit += f": {r.get('matched', 0):,} of {r.get('rows', 0):,} template items matched"
            if r.get("unknown"):
                bit += f", {r['unknown']} not found in BC"
        cols = r.get("columns") or []
        if cols:
            shown = cols[:6]
            bit += "; columns " + "; ".join(shown) + (f" and {len(cols) - 6} more" if len(cols) > 6 else "")
        lines.append(bit)
    if f.get("recalculates_on_open"):
        lines.append("Recalculate from these when opened in Excel (pivots refresh on open): "
                     + ", ".join(f["recalculates_on_open"][:15]) + ".")
    for p in f.get("not_refreshed") or []:
        lines.append(f"- Not refreshed: **{p['sheet']}** — {p.get('reason') or p.get('status')}")
    kept = [k["sheet"] for k in f.get("kept") or []]
    if kept:
        lines.append("Kept as your inputs (forecasts, trackers, lookups, calculated sheets): "
                     + ", ".join(kept[:15]) + (f" and {len(kept) - 15} more" if len(kept) > 15 else "") + ".")
    if f.get("missing_items_sheet"):
        lines.append(f"{f['missing_items']:,} active system items not in your template are listed on the "
                     f"**{f['missing_items_sheet']}** sheet.")
    return "\n".join(lines)
