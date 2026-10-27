"""Fill plan: which sheets to refresh from system data, how, and which columns to leave alone.

The LLM reads the workbook map (structure, formulas, data flow) plus the system field catalog and returns a
JSON plan. Code validates every entry against the catalog and the template (formula columns are never written).
When the LLM is unavailable, a heuristic plan built from the same rules is used.
"""
from __future__ import annotations

import json
import re

from excel_engine import llm, understand
from excel_engine.sources import ITEM_FIELDS, LEDGER_FIELDS, MONTH_MEASURES, TEXT_FIELDS

STRATEGIES = ("keyed_update", "replace_rows", "keep")
ACTIONS = ("fill", "fill_month", "fill_location", "keep_formula", "leave")

PLAN_SCHEMA = {
    "type": "object", "additionalProperties": False, "required": ["sheets"],
    "properties": {"sheets": {"type": "array", "items": {
        "type": "object", "additionalProperties": False,
        "required": ["sheet", "purpose", "strategy", "key_col", "entry_types", "item_prefixes",
                     "append_missing", "reason", "columns"],
        "properties": {
            "sheet": {"type": "string"},
            "purpose": {"type": "string"},
            "strategy": {"type": "string", "enum": list(STRATEGIES)},
            "key_col": {"type": ["string", "null"]},
            "entry_types": {"type": "array", "items": {"type": "string"}},
            "item_prefixes": {"type": "array", "items": {"type": "string"}},
            "append_missing": {"type": "boolean"},
            "reason": {"type": "string"},
            "columns": {"type": "array", "items": {
                "type": "object", "additionalProperties": False,
                "required": ["col", "action", "field", "measure", "month", "location", "reason"],
                "properties": {
                    "col": {"type": "string"},
                    "action": {"type": "string", "enum": list(ACTIONS)},
                    "field": {"type": ["string", "null"]},
                    "measure": {"type": ["string", "null"]},
                    "month": {"type": ["string", "null"]},
                    "location": {"type": ["string", "null"]},
                    "reason": {"type": ["string", "null"]}}}}}}}},
}

SYSTEM = """You are the Excel planning engine of Pulse, a supply-chain analytics system for a plastics manufacturer
running Microsoft Dynamics 365 Business Central. A user uploaded their own planning workbook and asked for a report
"in the same format" built from system data. You receive a map of the workbook (every sheet, its header row, column
types, sample formulas, tables, pivots, and which sheets read from which) and a catalog of system data.

Produce a fill plan that recreates the workbook with current system data while keeping its exact pattern:
- Think in data flow. Most planning workbooks have SOURCE sheets (raw ledger dumps, stock-by-location matrices,
  item lists with manual inputs) that feed CALCULATED sheets through VLOOKUP/XLOOKUP/SUMIF and pivots. Refresh the
  sources; calculated sheets, pivots and charts then recalculate by themselves when the file is opened.
- strategy "replace_rows": the sheet's rows are ledger entries (posting date, entry type, item, quantity...). Its
  rows are replaced by system ledger rows. Map every non-formula column to a ledger field (action "fill") or
  "leave"; formula columns are "keep_formula" (they are re-applied to every new row). entry_types = the entry types
  that appear in the template's Entry Type column (use the BC names from the catalog). item_prefixes = the item
  prefixes of the template's item column.
- strategy "keyed_update": the sheet is a list of items keyed by an item-number column (key_col). Keep the
  template's own rows and update only input (non-formula) columns that correspond to a system figure: item fields
  ("fill"), stock at a location ("fill_location" with the location code from the header; "" for an
  UNSPECIFIED / blank-location column), or monthly history
  ("fill_month" with measure and month YYYY-MM). Set append_missing true only for sheets that act as a master
  item list for the workbook (e.g. the stock matrix or the main MRP list), so system items absent from the template
  are listed on a separate sheet.
- strategy "keep": forecasts, projections, budgets, targets, BOM/ratio matrices, manual trackers (shipments, BL
  numbers, comments), lookup/mapping lists, notes, and sheets that are purely calculated or pivot output.
- Never fill a formula column. Never fill descriptive text the user maintains (application, comments, remarks).
- fill_month only for actual history (not forecast/projection), month within catalog.history_months and not after
  the reporting month. Resolve fiscal years from names: FY26 = Jul 2025 to Jun 2026. Months outside the history
  are "leave" with reason.
- Prefer exact meaning over header similarity: "Inventory"/"Stock" = stock_qty, "IN TRANSIT QTY"/"On order" =
  incoming_qty, "Average Cons"/"Max Cons" (monthly) = avg_monthly_consumed / max_monthly_consumed, "Cover days" =
  cover_days. Use the column hints given by the user/agent when present.
- Include EVERY sheet of the workbook exactly once with a one-line purpose. For keep sheets, columns may be empty.
- Never invent numbers; you only choose mappings.
"""


def understand_norm(text):
    s = str(text or "").strip().lower()
    s = re.sub(r"[\u2018\u2019'`]", "", s)
    s = re.sub(r"[^a-z0-9%]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


_ITEM_ALIASES = {
    "inventory": "stock_qty", "stock": "stock_qty", "on hand": "stock_qty", "stock on hand": "stock_qty",
    "qty on hand": "stock_qty", "closing stock": "stock_qty", "inventory qty": "stock_qty", "soh": "stock_qty",
    "in transit qty": "incoming_qty", "intransit qty": "incoming_qty", "in transit": "incoming_qty",
    "intransit": "incoming_qty", "on order": "incoming_qty", "incoming": "incoming_qty", "open po": "incoming_qty",
    "cover days": "cover_days", "days cover": "cover_days", "days of cover": "cover_days", "coverdays": "cover_days",
    "production days cover": "cover_prod_days", "prod days cover": "cover_prod_days",
    "average cons": "avg_monthly_consumed", "avg cons": "avg_monthly_consumed",
    "average consumption": "avg_monthly_consumed", "max cons": "max_monthly_consumed",
    "maximum consumption": "max_monthly_consumed", "average sales": "avg_monthly_sold",
    "max sales": "max_monthly_sold", "daily usage": "daily_usage", "usage per day": "daily_usage",
    "unit cost": "unit_cost", "stock value": "stock_value", "cover status": "cover_status",
    "uom": "uom", "unit of measure": "uom",
}
_LEDGER_ALIASES = {
    "posting date": "posting_date", "date": "posting_date", "entry type": "entry_type",
    "document type": "document_type", "document no": "document_no", "doc no": "document_no",
    "item no": "item_no", "item code": "item_no", "item": "item_no", "no": "item_no", "item number": "item_no",
    "description": "description", "item description": "description", "location code": "location",
    "location": "location", "quantity": "quantity", "qty": "quantity", "abs quantity": "abs_quantity",
    "remaining quantity": "remaining_quantity", "cost amount actual": "cost_amount",
    "sales amount actual": "sales_amount", "entry no": "entry_no", "item category code": "item_category",
    "open": "open", "month": "month_name", "month year": "month_year", "year": "year",
}
_KEEP_SHEETS = re.compile(r"forecast|projected|projection|budget|target|plan\b|matrix|key\b|notes?\b", re.I)


def _month_value(col_month, sheet_name, file_name, catalog):
    """Resolve a month header to YYYY-MM using its year, the fiscal year in the names, or the latest past month."""
    month, year = col_month.get("month"), col_month.get("year")
    rep = catalog["reporting_date"][:7]
    if not year:
        fy = re.search(r"\bFY\s?'?(\d{2,4})\b", f"{sheet_name} {file_name}", re.I)
        if fy:
            end = int(fy.group(1))
            end = end + 2000 if end < 100 else end
            year = end - 1 if month >= 7 else end
        else:
            ry, rm = int(rep[:4]), int(rep[5:7])
            year = ry if month <= rm else ry - 1
    return f"{year:04d}-{month:02d}"


def _measure_for(sheet_name, prefixes):
    name = sheet_name.lower()
    if re.search(r"sale|sold|demand|dispatch", name):
        return "sold"
    if re.search(r"cons|usage|issue", name):
        return "consumed"
    pre = set(prefixes or [])
    if pre and pre <= {"FG-"}:
        return "sold"
    if pre and not pre & {"FG-"}:
        return "consumed"
    return None


def heuristic_plan(wmap, catalog, hints=None):
    hints = {understand_norm(k): v for k, v in (hints or {}).items()}
    locations = set(catalog.get("locations") or [])
    sheets = []
    for s in wmap["sheets"]:
        cols = s.get("columns") or []
        entry = {"sheet": s["name"], "purpose": "", "strategy": "keep", "key_col": s.get("key_col"),
                 "entry_types": [], "item_prefixes": s.get("key_prefixes") or [], "append_missing": False,
                 "reason": "", "columns": []}
        ledger_map = {}
        for c in cols:
            n = understand_norm(c["header"])
            f = hints.get(n) if hints.get(n) in LEDGER_FIELDS else _LEDGER_ALIASES.get(n)
            if f and c["kind"] != "formula":
                ledger_map[c["col"]] = f
        lf = set(ledger_map.values())
        if {"posting_date", "item_no"} <= lf and lf & {"quantity", "abs_quantity"} and "entry_type" in lf:
            entry.update(strategy="replace_rows", purpose="Item ledger entry dump",
                         reason="Rows are ledger entries; replaced with system ledger rows for the same entry "
                                "types and item groups.")
            et_col = next((c for c in cols if ledger_map.get(c["col"]) == "entry_type"), None)
            known = set(catalog.get("entry_types") or [])
            if et_col and et_col.get("distinct"):
                entry["entry_types"] = [v for v in et_col["distinct"] if not known or v in known]
            for c in cols:
                if c["kind"] == "formula":
                    entry["columns"].append(_col(c["col"], "keep_formula"))
                elif c["col"] in ledger_map:
                    entry["columns"].append(_col(c["col"], "fill", field=ledger_map[c["col"]]))
                else:
                    entry["columns"].append(_col(c["col"], "leave", reason="no system ledger field"))
            sheets.append(entry)
            continue
        if not s.get("key_col") or _KEEP_SHEETS.search(s["name"]):
            entry["purpose"] = "Reference / calculated sheet"
            entry["reason"] = "No item key column or a forecast/plan/matrix sheet; kept as is."
            sheets.append(entry)
            continue
        measure = _measure_for(s["name"], s.get("key_prefixes"))
        actions = []
        for c in cols:
            if c["col"] == s["key_col"]:
                continue
            if c["kind"] == "formula":
                actions.append(_col(c["col"], "keep_formula"))
                continue
            n = understand_norm(c["header"])
            field = hints.get(n) if hints.get(n) in ITEM_FIELDS else _ITEM_ALIASES.get(n)
            loc = _location_code(c["header"])
            if loc is not None and (loc in locations or loc == ""):
                actions.append(_col(c["col"], "fill_location", location=loc))
            elif field and field not in TEXT_FIELDS - {"cover_status"}:
                actions.append(_col(c["col"], "fill", field=field))
            elif c.get("month") and measure:
                month = _month_value(c["month"], s["name"], wmap["name"], catalog)
                actions.append(_col(c["col"], "fill_month", measure=measure, month=month))
        if any(a["action"] in ("fill", "fill_month", "fill_location") for a in actions):
            feeds = s.get("feeds") or []
            entry.update(strategy="keyed_update", columns=actions,
                         purpose="Item list keyed by " + s["key_col"],
                         append_missing=bool(feeds) and (s.get("key_count") or 0) >= 50,
                         reason="Template item rows kept; system figures written into input columns.")
            prefixes = s.get("key_prefixes") or []
            entry["item_prefixes"] = prefixes
        else:
            entry["purpose"] = "Item list with manual or calculated columns"
            entry["reason"] = "No input column maps to a system figure."
        sheets.append(entry)
    return {"sheets": sheets}


_NO_LOCATION = re.compile(r"^\(?(unspecified|blank|no location|none)\)?$", re.I)


def _location_code(header):
    """Location code a stock-matrix header stands for; '' is BC's blank location code."""
    text = str(header if header is not None else "").strip()
    if not text:
        return None
    if _NO_LOCATION.match(text):
        return ""
    return text.upper() if re.match(r"^[A-Za-z0-9][A-Za-z0-9 _.\-/]{0,19}$", text) else None


def _col(col, action, field=None, measure=None, month=None, location=None, reason=None):
    return {"col": col, "action": action, "field": field, "measure": measure, "month": month,
            "location": location, "reason": reason}


def llm_plan(wmap, catalog, instruction="", hints=None):
    compact = understand.compact(wmap)
    cat = {"item_fields": catalog["item_fields"], "month_measures": catalog["month_measures"],
           "ledger_fields": catalog["ledger_fields"], "locations": catalog["locations"],
           "history_months": catalog["history_months"], "reporting_date": catalog["reporting_date"],
           "entry_types": catalog["entry_types"]}
    user = ("USER REQUEST: " + (instruction or "Create a report from system data in the same format.") + "\n\n"
            + ("COLUMN HINTS (template header -> system field): " + json.dumps(hints) + "\n\n" if hints else "")
            + "SYSTEM CATALOG:\n" + json.dumps(cat) + "\n\nWORKBOOK MAP:\n" + json.dumps(compact))
    return llm.chat_json(SYSTEM, user, PLAN_SCHEMA, name="fill_plan")


def validate_plan(raw, wmap, catalog, file_name=""):
    """Drop or downgrade anything not backed by the template or the catalog. Returns (plan, warnings)."""
    warnings = []
    by_name = {s["name"]: s for s in wmap["sheets"]}
    locations = set(catalog.get("locations") or [])
    months = set(catalog.get("history_months") or [])
    rep_month = catalog["reporting_date"][:7]
    known_types = set(catalog.get("entry_types") or [])
    seen = set()
    out = []
    for entry in (raw or {}).get("sheets") or []:
        name = entry.get("sheet")
        s = by_name.get(name)
        if not s or name in seen:
            continue
        seen.add(name)
        e = {"sheet": name, "purpose": (entry.get("purpose") or "")[:200],
             "strategy": entry.get("strategy") if entry.get("strategy") in STRATEGIES else "keep",
             "key_col": entry.get("key_col"), "entry_types": [], "item_prefixes": [],
             "append_missing": bool(entry.get("append_missing")), "reason": (entry.get("reason") or "")[:300],
             "columns": []}
        cols = {c["col"]: c for c in s.get("columns") or []}
        if e["strategy"] == "keyed_update":
            if e["key_col"] not in cols:
                e["key_col"] = s.get("key_col")
            if not e["key_col"]:
                warnings.append(f"{name}: no item key column; kept")
                e["strategy"] = "keep"
        prefixes = [p for p in entry.get("item_prefixes") or [] if re.match(r"^[A-Z]{1,5}-?$", str(p))]
        e["item_prefixes"] = [p if p.endswith("-") else p + "-" for p in prefixes] or (s.get("key_prefixes") or [])
        if e["strategy"] == "replace_rows":
            types = [t for t in entry.get("entry_types") or [] if not known_types or t in known_types]
            e["entry_types"] = types
        for c in entry.get("columns") or []:
            col = cols.get(c.get("col"))
            if not col:
                continue
            action = c.get("action") if c.get("action") in ACTIONS else "leave"
            item = _col(col["col"], action, c.get("field"), c.get("measure"), c.get("month"), c.get("location"),
                        (c.get("reason") or None))
            if col["kind"] == "formula" and action in ("fill", "fill_month", "fill_location"):
                warnings.append(f"{name}!{col['col']} is a formula column; kept")
                item.update(action="keep_formula", field=None)
            if e["strategy"] == "keyed_update" and col["col"] == e["key_col"]:
                continue
            if item["action"] == "fill":
                catalog_fields = LEDGER_FIELDS if e["strategy"] == "replace_rows" else ITEM_FIELDS
                if item["field"] not in catalog_fields:
                    warnings.append(f"{name}!{col['col']}: unknown field {item['field']}")
                    item.update(action="leave", reason=f"no system field ({item['field']})", field=None)
            elif item["action"] == "fill_location":
                loc = "" if item["location"] == "" else _location_code(item["location"] or col["header"])
                if loc is None or (loc and loc not in locations):
                    item.update(action="leave", reason=f"location {loc} not in Business Central")
                else:
                    item["location"] = loc
            elif item["action"] == "fill_month":
                month = item["month"] or (col.get("month") and _month_value(col["month"], name, file_name, catalog))
                if item["measure"] not in MONTH_MEASURES or not month or not re.match(r"^\d{4}-\d{2}$", str(month)):
                    item.update(action="leave", reason="month or measure not recognised")
                elif month > rep_month:
                    item.update(action="leave", reason=f"{month} is a future month (forecast input kept)")
                elif month not in months:
                    item.update(action="leave", reason=f"{month} is outside the 12-month system history")
                else:
                    item["month"] = month
            e["columns"].append(item)
        writes = [c for c in e["columns"] if c["action"] in ("fill", "fill_month", "fill_location")]
        if e["strategy"] == "replace_rows":
            fields = {c["field"] for c in writes}
            if not ({"item_no"} <= fields and fields & {"quantity", "abs_quantity"}):
                warnings.append(f"{name}: ledger replace needs item and quantity columns; kept")
                e["strategy"] = "keep"
            # formula columns of a replaced table are re-applied to every new row
            listed = {c["col"] for c in e["columns"]}
            for col in s.get("columns") or []:
                if col["kind"] == "formula" and col["col"] not in listed:
                    e["columns"].append(_col(col["col"], "keep_formula"))
        elif e["strategy"] == "keyed_update" and not writes:
            e["strategy"] = "keep"
            e["reason"] = e["reason"] or "No input column maps to a system figure."
        out.append(e)
    for s in wmap["sheets"]:
        if s["name"] not in seen:
            out.append({"sheet": s["name"], "purpose": "", "strategy": "keep", "key_col": None, "entry_types": [],
                        "item_prefixes": [], "append_missing": False, "reason": "Not in plan; kept as is.",
                        "columns": []})
    order = {s["name"]: i for i, s in enumerate(wmap["sheets"])}
    out.sort(key=lambda e: order.get(e["sheet"], 999))
    return {"sheets": out}, warnings


def make_plan(wmap, catalog, instruction="", hints=None):
    """LLM plan when Azure OpenAI is configured, validated; heuristic plan otherwise or on failure."""
    source, error = "heuristic", None
    raw = None
    if llm.available():
        try:
            raw = llm_plan(wmap, catalog, instruction, hints)
            source = "llm"
        except Exception as e:
            error = f"{type(e).__name__}: {str(e)[:200]}"
    if raw is None:
        raw = heuristic_plan(wmap, catalog, hints)
    plan, warnings = validate_plan(raw, wmap, catalog, wmap.get("name") or "")
    if source == "llm" and not any(e["strategy"] != "keep" for e in plan["sheets"]):
        # LLM found nothing to refresh; fall back to the rule-based plan rather than returning an empty rebuild.
        plan, w2 = validate_plan(heuristic_plan(wmap, catalog, hints), wmap, catalog, wmap.get("name") or "")
        warnings += ["LLM plan had no refreshable sheet; rule-based plan used"] + w2
        source = "heuristic"
    plan["source"] = source
    plan["llm_error"] = error
    plan["warnings"] = warnings[:50]
    return plan
