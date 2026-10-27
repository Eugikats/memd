"""Raw Material Planning: the Supply Chain RM Planning workbook, calculated from Business Central.

Stock, consumption and open purchase orders come from the model. Planner inputs that Business Central
does not hold (forecasts, the resin recipe matrix, alternatives, the shipment tracker, comments and
targets) come from the workbook pack. Formulas match the workbook; the 90-day cover math is not used.
"""
from __future__ import annotations

import gzip
import json
import math
import re
from pathlib import Path

import numpy as np
import pandas as pd

PACK_PATH = Path(__file__).with_name("rm_pack.json.gz")

# Consmptn Dump keeps these entry types, then takes the absolute quantity.
DUMP_TYPES = ("Consumption", "Negative Adjmt.", "Positive Adjmt.")
# Sheet6's pivot is Consumption only.
SHEET6_TYPES = ("Consumption",)

# Stock sheet, column D through BD, then the three formula columns.
LOCATIONS = [
    "", "ADMIN", "DIGITAL", "HO-STORE", "JENI", "JUMIA", "KISUGUFOS", "MAINT-BUG", "MIT", "OPBFG",
    "QC REWORK", "RECYCLE", "ROUTEWH", "SHIP1BRAND", "SHIP1PR", "SHIP1RM", "SHIP1SPN", "SHIP1SPU",
    "SHIP1WIP", "SHIP2BRAND", "SHIP2FG", "SHIP2FG-1F", "SHIP2FG-PR", "SHIP2PR", "SHIP2RM", "SHIP3BOND",
    "SHIP3BRAND", "SHIP3FG", "SHIP3FG-1F", "SHIP3PR", "SHIP3RM", "SHIP4BRAND", "SHIP4FG", "SHIP4PR",
    "SHIP4RM", "SHIP4RM-PR", "SHIP5BOND", "SHIP5BRAND", "SHIP5FG", "SHIP5FG-1F", "SHIP5FG-2F",
    "SHIP5FG-BD", "SHIP5FGJFD", "SHIP5PR", "SHIP5RM", "SHIP5SP", "SHIP5SUND", "SHIP6FG", "STRIKERS",
    "TFUMA", "UBLFOS", "UMA- MIU", "UMAFOS",
]
LOCATION_LABEL = {"": "UNSPECIFIED"}
# SUM(UNSPECIFIED : UMA- MIU) — UMAFOS is the next column and is outside the sum.
TOTAL_LOCS = [c for c in LOCATIONS if c != "UMAFOS"]
SUM_RM = ["SHIP1RM", "SHIP2RM", "SHIP3RM", "SHIP4RM", "SHIP5BOND", "SHIP5RM"]
RITAH = ["SHIP1SPU", "SHIP6FG", "SHIP5PR", "SHIP5FGJFD", "SHIP5FG", "SHIP4RM", "SHIP4PR", "SHIP4FG",
         "SHIP3PR", "SHIP3FG-1F", "SHIP3BRAND", "SHIP3FG", "SHIP2FG-PR", "SHIP5BRAND"]

F_AVG = "Average of months with movement = SUM of monthly absolute quantity / COUNT of those months (pivot Average)."
F_MAX = "Maximum monthly absolute quantity (pivot Maximum)."
F_INV = "SUM RM = SHIP1RM + SHIP2RM + SHIP3RM + SHIP4RM + SHIP5BOND + SHIP5RM."
F_MIN = "ROUNDUP(((Average Cons / 26 * 30) / 25), 0) * 25."
F_MAX45 = "Minimum (30 days) * 1.5."
F_OPEN = "Inventory - Average Cons * 2.5 + IN TRANSIT QTY."
F_PROJ = "INDEX of RM MATRIX row 134, matched on the item code in row 1. Planner recipe, not a BC figure."
F_PUR = "Anticipated opening (Dec) - projected consumption (Jan)."
F_NOV = ("IF the purchase quantity is zero or negative, Inventory / 26 * 15; otherwise the purchase quantity. "
         "This is the workbook's Opening - Nov formula.")
F_ORDER = "ROUNDUP((ABS(Purchase) + Opening Nov) / 25500, 0) * 25500."
F_COVER = "Inventory / Average Cons * 26. Blank when there is no consumption."
F_STATUS = 'SLOB when Average Cons <= 0, otherwise ACTIVE.'
F_FLAG = "CRITICAL when cover days <= 29 (a blank cover counts as 0), otherwise OKAY."
F_TGT = "CRITICAL when the planner target >= cover days, otherwise OKAY. A blank target counts as 0."
F_TRANSIT = "Open purchase-order quantity still to arrive. The workbook cell is typed by hand; here it is the BC figure."
F_PLUS = "Inventory + in transit."
F_COVER_T = "(Inventory + in transit) / Average Cons * 26."
F_OCT = "October order (planner input) + inventory + in transit."
F_COVER_O = "(Inventory + in transit + October order) / Average Cons * 26."

SHEETS = [
    ("Planning", "MRP", "MRP"),
    ("Planning", "Resin Material", "Resin Material"),
    ("Planning", "Masterbatch", "Masterbatch"),
    ("Planning", "Local & Others", "Local & Others"),
    ("Planning", "RISKS", "RISKS"),
    ("Planning", "Milk and crates", "Milk and crates"),
    ("Planning", "Others", "Others"),
    ("Planning", "All Purchase Plans", "All Purchase Plans"),
    ("Planning", "Local Purchase", "Local Purchase"),
    ("Stock", "Stock", "Stock"),
    ("Consumption", "Consmptn Dump", "Consmptn Dump"),
    ("Consumption", "Pivot Tables", "Pivot Tables"),
    ("Consumption", "Sheet6", "Sheet6"),
    ("Forecast and recipe", "RM MATRIX", "RM MATRIX"),
    ("Forecast and recipe", "Projected Sales", "Projected Sales "),
    ("Forecast and recipe", "Sales FY26", "Sales FY26"),
    ("Forecast and recipe", "Sales rev. forecast", "Sales rev. forecast mar."),
    ("Reference", "Alternatives", "Alternatives"),
    ("Reference", "KEY", "KEY"),
    ("Reference", "Intransit", "Intransit"),
    ("Reference", "my sheet", "my sheet"),
    ("Reference", "Sheet1", "Sheet1"),
]

_CACHE: dict = {}
_PACK = None


def _pack():
    global _PACK
    if _PACK is None:
        _PACK = json.loads(gzip.open(PACK_PATH, "rb").read().decode("utf-8"))
    return _PACK


def _clean(v):
    if v is None or v == "":
        return None
    if isinstance(v, (float, np.floating)):
        if math.isnan(v) or math.isinf(v):
            return None
        v = float(v)
        if abs(v) >= 100:
            return round(v, 1)
        if abs(v) >= 1:
            return round(v, 2)
        return round(v, 4)
    if isinstance(v, (np.integer,)):
        return int(v)
    if isinstance(v, (int,)):
        return v
    return v


def _num(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
            return 0.0
        return float(v)
    return 0.0


def _roundup(value, step):
    """Excel ROUNDUP to a multiple of step, for positive quantities."""
    if value <= 0 or step <= 0:
        return 0.0
    units = round(value / step, 8)
    return math.ceil(units - 1e-9) * step


def _norm(text):
    return re.sub(r"\s+", " ", str(text or "")).strip()


def _slug(text):
    s = re.sub(r"[^A-Za-z0-9]+", "_", str(text or "x")).strip("_")
    return s or "x"


def _col(field, label, kind, formula=None):
    out = {"field": field, "label": label, "kind": kind}
    if formula:
        out["formula"] = formula
    return out


def _rows(records, columns):
    fields = [c["field"] for c in columns]
    return [{f: _clean(r.get(f)) for f in fields} for r in records]


def _table(title, records, columns, subtitle=None):
    out = {"title": title, "total_rows": len(records), "drill": "item_no", "bars": False,
           "columns": columns, "rows": _rows(records, columns)}
    if subtitle:
        out["subtitle"] = subtitle
    return out


def _monthly(model):
    """Absolute quantity by item and by description, for the dump's entry types."""
    from .model import load_history
    h = load_history()
    empty = pd.DataFrame(columns=["item_no", "month", "abs_qty"])
    if h.empty:
        return empty, empty, []
    h = h.rename(columns={"Item_No": "item_no", "Entry_Type": "entry_type", "Quantity": "qty",
                          "Posting_Date": "posting_date", "Item_Description": "description"})
    h["month"] = h.posting_date.astype(str).str[:7]
    h["abs_qty"] = pd.to_numeric(h.qty, errors="coerce").fillna(0).abs()
    dump = h[h.entry_type.isin(DUMP_TYPES)]
    by_item = dump.groupby(["item_no", "month"], as_index=False).abs_qty.sum()
    sheet6 = h[h.entry_type.isin(SHEET6_TYPES)]
    by_desc = sheet6.groupby(["description", "month"], as_index=False).abs_qty.sum()
    months = sorted(set(by_item.month.tolist()) | set(by_desc.month.tolist()))
    return by_item, by_desc, months


def _stats(by_item):
    """Average = mean of months with quantity > 0. Max = largest month. Same as the pivot formulas."""
    if by_item.empty:
        return {}, {}
    pos = by_item[by_item.abs_qty > 0]
    avg = pos.groupby("item_no").abs_qty.mean().to_dict()
    mx = pos.groupby("item_no").abs_qty.max().to_dict()
    return avg, mx


def _mrp_rows(model, pack, avg, mx, sum_rm, incoming):
    required = pack.get("matrix_required") or {}
    alt = {_norm(a["item"]): a["alt_code"] for a in pack.get("alternatives") or [] if a.get("alt_code")}
    rows = []
    for src in pack.get("mrp") or []:
        item = src["item"]
        average = float(avg.get(item, 0) or 0)
        maximum = float(mx.get(item, 0) or 0)
        inventory = float(sum_rm.get(item, 0) or 0)
        transit = float(incoming.get(item, 0) or 0)
        minimum = _roundup((average / 26 * 30) / 25, 25) if average else None
        maximum_45 = minimum * 1.5 if minimum is not None else None
        opening_dec = inventory - average * 2.5 + transit
        projected = float(required.get(item, 0) or 0)
        purchase = opening_dec - projected
        opening_nov = (inventory / 26) * 15 if purchase <= 0 else abs(purchase)
        order_qty = _roundup(abs(purchase) + opening_nov, 25500)
        cover = inventory / average * 26 if average else None
        cover_n = cover if cover is not None else 0.0
        status = "SLOB" if average <= 0 else "ACTIVE"
        flag = "CRITICAL" if cover_n <= 29 else "OKAY"
        target = src.get("target") if isinstance(src.get("target"), (int, float)) else None
        target_cover = "CRITICAL" if _num(target) >= cover_n else "OKAY"
        rows.append({
            "item_no": item, "application": src.get("application") or None,
            "description": src.get("description") or None, "lc_im": src.get("lc_im") or None,
            "average": average, "max_cons": maximum, "inventory": inventory, "minimum": minimum,
            "maximum": maximum_45, "opening_dec": opening_dec, "projected_jan": projected,
            "purchase": purchase, "in_transit": transit, "comments": src.get("comments") or None,
            "opening_nov": opening_nov, "order_qty": order_qty, "cover_days": cover, "status": status,
            "target": target, "cover_flag": flag, "target_cover": target_cover,
            "remarks": src.get("remarks") or None, "alt_code": alt.get(_norm(item), item),
        })
    return rows


def _index(rows):
    by_item = {r["item_no"]: r for r in rows}
    by_desc = {}
    for r in rows:
        key = _norm(r.get("description"))
        if key and key not in by_desc:
            by_desc[key] = r
    return by_item, by_desc


def _lookup(by_item, by_desc, item, description):
    if item in by_item:
        return by_item[item]
    return by_desc.get(_norm(description))


def _board(spec_rows, by_item, by_desc, extra):
    """Group sheets look up MRP (the workbook uses VLOOKUP on the description)."""
    out = []
    for src in spec_rows:
        base = _lookup(by_item, by_desc, src["item"], src.get("description")) or {}
        average = base.get("average") or 0
        inventory = base.get("inventory") or 0
        transit = base.get("in_transit") or 0
        cover = inventory / average * 26 if average else None
        plus = inventory + transit
        cover_t = plus / average * 26 if average else None
        rec = {
            "item_no": src["item"], "application": src.get("application") or None,
            "description": src.get("description") or base.get("description"),
            "lc_im": src.get("lc_im") or None, "average": average, "max_cons": base.get("max_cons") or 0,
            "inventory": inventory, "in_transit": transit, "cover_days": cover,
            "on_hand_plus": plus, "cover_with_transit": cover_t,
        }
        rec.update(extra(src, rec))
        out.append(rec)
    return out


def _resin_extra(src, rec):
    order = _num(src.get("october_order"))
    plus = (rec.get("on_hand_plus") or 0) + order
    average = rec.get("average") or 0
    return {"october_order": order or None, "with_october": plus,
            "cover_with_october": plus / average * 26 if average else None}


def _stock_map(model):
    sl = model["stock_loc"]
    if sl.empty:
        return {}
    loc = sl.location.fillna("").astype(str).str.strip()
    qty = sl.stock_qty.fillna(0)
    frame = pd.DataFrame({"item_no": sl.item_no.to_numpy(), "location": loc.to_numpy(), "qty": qty.to_numpy()})
    piv = frame.pivot_table(index="item_no", columns="location", values="qty", aggfunc="sum", fill_value=0)
    return {item: row.to_dict() for item, row in piv.iterrows()}


def _sum_locs(stock, item, names):
    row = stock.get(item) or {}
    return float(sum(row.get(n, 0) or 0 for n in names))


def _resolve_formula(formula, key_text, by_desc):
    """XLOOKUP/VLOOKUP into MRP columns, keyed by the description in column A or C."""
    if not isinstance(formula, str) or "MRP!" not in formula.upper():
        return None
    match = re.search(r"MRP!(\$?[A-Z]{1,2}):", formula.upper())
    if not match or "LOOKUP" not in formula.upper():
        return None
    field = {"E": "average", "F": "max_cons", "G": "inventory", "M": "in_transit",
             "P": "order_qty", "Q": "cover_days"}.get(match.group(1).replace("$", ""))
    rec = by_desc.get(_norm(key_text))
    if not field or not rec:
        return None
    return rec.get(field)


def _cell_value(value, key_text, by_desc):
    if isinstance(value, dict):
        live = _resolve_formula(value.get("f"), key_text, by_desc)
        if live is not None:
            return live
        return value.get("v")
    return value


def _grid_table(title, grid, by_desc=None, subtitle=None):
    grid = grid or []
    if not grid:
        return _table(title, [], [_col("note", "Note", "text")], subtitle)
    header_i = 0
    for i, rec in enumerate(grid[:5]):
        strings = sum(1 for v in rec["c"].values() if isinstance(v, str) and not isinstance(v, dict))
        if strings >= 3:
            header_i = i
            break
    letters = []
    for rec in grid:
        for letter in rec["c"]:
            if letter not in letters:
                letters.append(letter)
    letters.sort(key=lambda s: (len(s), s))
    labels = {}
    head = grid[header_i]["c"]
    for letter in letters:
        raw = head.get(letter)
        text = raw if isinstance(raw, str) else ""
        labels[letter] = (text.strip() or letter)[:48]
    seen = {}
    columns = [_col("row", "Row", "qty")]
    for letter in letters:
        label = labels[letter]
        n = seen.get(label, 0) + 1
        seen[label] = n
        field = _slug(label) if n == 1 else _slug(f"{label}_{letter}")
        columns.append(_col(field, label, "text"))
    fields = [c["field"] for c in columns[1:]]
    records = []
    for rec in grid:
        if rec is grid[header_i]:
            continue
        cells = rec["c"]
        key = cells.get("A")
        key = key.get("v") if isinstance(key, dict) else key
        row = {"row": rec["r"]}
        blank = True
        for letter, field in zip(letters, fields):
            value = _cell_value(cells.get(letter), key, by_desc)
            row[field] = value
            if value not in (None, ""):
                blank = False
        if not blank:
            records.append(row)
    # Numeric-looking columns get a qty kind so they align and sort as numbers.
    for col, letter in zip(columns[1:], letters):
        nums = [r.get(col["field"]) for r in records if isinstance(r.get(col["field"]), (int, float))]
        if nums and len(nums) >= max(1, int(0.5 * len(records))):
            col["kind"] = "qty"
    return _table(title, records, columns, subtitle)


def compute(model):
    key = (model.get("notes") or {}).get("extracted_at")
    if key in _CACHE:
        return _CACHE[key]
    pack = _pack()
    by_item_m, by_desc_m, months = _monthly(model)
    avg, mx = _stats(by_item_m)
    stock = _stock_map(model)
    sum_rm = {item: _sum_locs(stock, item, SUM_RM) for item in stock}
    item = model["item"]
    incoming = item.set_index("item_no").incoming_qty.fillna(0).to_dict() if len(item) else {}
    # Items on the planning lists with no open stock still need a zero SUM RM.
    for src in pack.get("mrp") or []:
        sum_rm.setdefault(src["item"], 0.0)
    mrp = _mrp_rows(model, pack, avg, mx, sum_rm, incoming)
    by_item, by_desc = _index(mrp)
    names = model["items"].set_index("item_no").description.to_dict()
    classes = model["items"].set_index("item_no").item_class.to_dict()
    costs = {}
    if "unit_cost" in model["items"].columns:
        costs = model["items"].set_index("item_no").unit_cost.fillna(0).astype(float).to_dict()
    sl = model["stock_loc"]
    at_site, at_loc = {}, {}
    if len(sl):
        for item_no, site, loc in zip(sl.item_no, sl.site.fillna(""), sl.location.fillna("")):
            at_site.setdefault(item_no, set()).add(str(site))
            at_loc.setdefault(item_no, set()).add(str(loc).strip())
    built = {"pack": pack, "mrp": mrp, "by_item": by_item, "by_desc": by_desc, "months": months,
             "by_item_m": by_item_m, "by_desc_m": by_desc_m, "stock": stock, "names": names,
             "classes": classes, "at_site": at_site, "at_loc": at_loc, "costs": costs}
    _CACHE.clear()
    _CACHE[key] = built
    return built


def _mrp_table(built):
    cols = [
        _col("item_no", "Item Code", "text"), _col("application", "Application", "text"),
        _col("description", "Description", "text"), _col("lc_im", "LC/IM", "text"),
        _col("average", "Average Cons", "qty", F_AVG), _col("max_cons", "Max Cons", "qty", F_MAX),
        _col("inventory", "Inventory", "qty", F_INV), _col("minimum", "Minimum (30 days)", "qty", F_MIN),
        _col("maximum", "Maximum (45 days)", "qty", F_MAX45),
        _col("opening_dec", "Anticipated opening Dec", "qty", F_OPEN),
        _col("projected_jan", "Projected cons Jan", "qty", F_PROJ),
        _col("purchase", "Purchase", "qty", F_PUR),
        _col("in_transit", "IN TRANSIT QTY", "qty", F_TRANSIT),
        _col("comments", "Comments", "text"),
        _col("opening_nov", "Opening Nov", "qty", F_NOV),
        _col("order_qty", "Order Qty", "qty", F_ORDER),
        _col("cover_days", "Cover days", "days", F_COVER),
        _col("status", "Status", "status", F_STATUS),
        _col("target", "Target", "qty"),
        _col("cover_flag", "Cover flag", "status", F_FLAG),
        _col("target_cover", "Target cover", "status", F_TGT),
        _col("remarks", "Remarks", "text"),
    ]
    return _table("MRP", built["mrp"], cols, "The planning list. Group sheets look up this one.")


def _group_table(built, key, title, transit_label, resin=False):
    rows = _board(built["pack"].get(key) or [], built["by_item"], built["by_desc"],
                  _resin_extra if resin else (lambda src, rec: {}))
    cols = [
        _col("item_no", "Item Code", "text"), _col("application", "Application", "text"),
        _col("description", "Description", "text"), _col("lc_im", "LC/IM", "text"),
        _col("average", "Average Cons", "qty", "VLOOKUP of the description into MRP Average Cons."),
        _col("max_cons", "Max Cons", "qty", "VLOOKUP into MRP Max Cons."),
        _col("inventory", "Inventory", "qty", "VLOOKUP into MRP Inventory (SUM RM)."),
        _col("in_transit", transit_label, "qty", F_TRANSIT),
        _col("cover_days", "Cover days", "days", F_COVER),
        _col("on_hand_plus", "Inventory + in transit", "qty", F_PLUS),
        _col("cover_with_transit", "Cover days + in transit", "days", F_COVER_T),
    ]
    if resin:
        cols += [_col("october_order", "October order", "qty", "Planner input from the workbook."),
                 _col("with_october", "Inventory + in transit + October order", "qty", F_OCT),
                 _col("cover_with_october", "Cover days with October order", "days", F_COVER_O)]
    return _table(title, rows, cols, "Same rows as the workbook. Figures are the MRP row for that item.")


def _others(built):
    rows = []
    for src in built["pack"].get("others") or []:
        base = _lookup(built["by_item"], built["by_desc"], src["item"], src.get("description")) or {}
        average = base.get("average") or 0
        inventory = base.get("inventory") or 0
        rows.append({
            "item_no": src["item"], "application": src.get("application") or None,
            "description": src.get("description") or base.get("description"),
            "lc_im": src.get("lc_im") or None, "average": average, "max_cons": base.get("max_cons") or 0,
            "inventory": inventory, "in_transit": base.get("in_transit") or 0,
            "order_qty": src.get("order_qty") if isinstance(src.get("order_qty"), (int, float)) else base.get("order_qty"),
            "cover_days": inventory / average * 26 if average else None,
            "status": "SLOB" if average <= 0 else "ACTIVE",
        })
    cols = [
        _col("item_no", "Item Code", "text"), _col("application", "Application", "text"),
        _col("description", "Description", "text"), _col("lc_im", "LC/IM", "text"),
        _col("average", "Average Cons", "qty", F_AVG), _col("max_cons", "Max Cons", "qty", F_MAX),
        _col("inventory", "Inventory", "qty", F_INV), _col("in_transit", "IN TRANSIT QTY", "qty", F_TRANSIT),
        _col("order_qty", "Order Qty", "qty", "Planner figure from the sheet where it was typed; otherwise the MRP order quantity."),
        _col("cover_days", "Cover days", "days", F_COVER), _col("status", "Status", "status", F_STATUS),
    ]
    return _table("Others", rows, cols)


def _stock_table(built):
    planning = set(built["by_item"])
    alt = {_norm(a["item"]): a["alt_code"] for a in built["pack"].get("alternatives") or [] if a.get("alt_code")}
    for key in ("resin", "masterbatch", "local", "risks", "milk", "others"):
        planning.update(r["item"] for r in built["pack"].get(key) or [])
    known = set(LOCATIONS)
    records = []
    items = set(built["stock"]) | planning
    names = built["names"]
    for item in items:
        row = built["stock"].get(item) or {}
        if not row and item not in planning:
            continue
        rec = {"item_no": item, "description": names.get(item),
               "code": alt.get(_norm(item), item)}
        for loc in LOCATIONS:
            qty = row.get(loc, 0) or 0
            rec["l_" + _slug(loc or "UNSPECIFIED")] = qty or None
        other = sum(v for k, v in row.items() if k not in known)
        rec["other"] = other or None
        rec["sum_rm"] = sum(row.get(n, 0) or 0 for n in SUM_RM)
        rec["total"] = sum(row.get(n, 0) or 0 for n in TOTAL_LOCS)
        rec["ritah"] = sum(row.get(n, 0) or 0 for n in RITAH)
        if rec["total"] or rec["sum_rm"] or other or item in planning:
            records.append(rec)
    records.sort(key=lambda r: r["item_no"])
    cols = [_col("item_no", "No.", "text"),
            _col("code", "Code", "text", "VLOOKUP into Alternatives; the item number when it has no alternative."),
            _col("description", "Description", "text")]
    for loc in LOCATIONS:
        label = LOCATION_LABEL.get(loc, loc)
        cols.append(_col("l_" + _slug(loc or "UNSPECIFIED"), label, "qty"))
    cols += [
        _col("sum_rm", "SUM RM", "qty", F_INV),
        _col("total", "Total", "qty", "Sum of UNSPECIFIED through UMA- MIU. UMAFOS is outside this sum, as in the workbook."),
        _col("ritah", "Ritah", "qty", "Sum of the locations in the workbook's Ritah formula."),
        _col("other", "Other locations", "qty", "Stock at locations that the workbook's matrix does not list."),
    ]
    return _table("Stock", records, cols, "On-hand quantity from open item-ledger entries, one column per location on the workbook's stock sheet.")


def _pivot(built, by, key_field, key_label, title, subtitle, types_note):
    frame = built[by]
    months = built["months"]
    if frame.empty:
        return _table(title, [], [_col(key_field, key_label, "text")], subtitle)
    piv = frame.pivot_table(index=key_field if key_field != "item_no" else "item_no",
                            columns="month", values="abs_qty", aggfunc="sum", fill_value=0)
    records = []
    names = built["names"]
    for key, row in piv.iterrows():
        vals = [float(row.get(m, 0) or 0) for m in months]
        used = [v for v in vals if v > 0]
        rec = {key_field: key, "average": (sum(used) / len(used)) if used else 0,
               "maximum": max(vals) if vals else 0, "total": sum(vals)}
        if key_field == "item_no":
            rec["description"] = names.get(key)
        for m, v in zip(months, vals):
            rec["m_" + m] = v
        records.append(rec)
    records.sort(key=lambda r: r.get("total") or 0, reverse=True)
    cols = [_col(key_field, key_label, "text")]
    if key_field == "item_no":
        cols.append(_col("description", "Description", "text"))
    for m in months:
        cols.append(_col("m_" + m, m, "qty", "Sum of absolute quantity in the month."))
    cols += [_col("average", "Average", "qty", F_AVG + " " + types_note),
             _col("maximum", "Maximum", "qty", F_MAX),
             _col("total", "Total", "qty")]
    return _table(title, records, cols, subtitle)


def _sheet1(built):
    rows = [{"item_no": r["item"], "description": r.get("name")} for r in built["pack"].get("sheet1") or []]
    return _table("Sheet1", rows, [_col("item_no", "Item", "text"), _col("description", "Description", "text")],
                  "Item lookup used by the consumption dump's Code column.")


def _alternatives(built):
    rows = [{"item_no": r["item"], "alt_description": r.get("alt_description"), "alt_code": r.get("alt_code"),
             "description": r.get("description"), "other_code": r.get("other_code")}
            for r in built["pack"].get("alternatives") or []]
    cols = [_col("item_no", "Item Code", "text"), _col("alt_description", "Alternative grade", "text"),
            _col("alt_code", "Alternative code", "text", "What Stock's Code column and the dump's Alternative column return."),
            _col("description", "Description", "text"), _col("other_code", "Other code", "text")]
    return _table("Alternatives", rows, cols, "Planner cross-reference. Not held in Business Central.")


def _key(built):
    rows = [{"bc_name": r["bc_name"], "old_name": r["old_name"]} for r in built["pack"].get("key") or []]
    return _table("KEY", rows, [_col("bc_name", "BC name", "text"), _col("old_name", "Old name", "text")],
                  "Names the planner maps onto Business Central descriptions.")


def sheet_table(built, sheet):
    pack = built["pack"]
    if sheet == "MRP":
        return _mrp_table(built)
    if sheet == "Resin Material":
        return _group_table(built, "resin", "Resin Material", "IN TRANSIT QTY", resin=True)
    if sheet == "Masterbatch":
        return _group_table(built, "masterbatch", "Masterbatch", "IN TRANSIT QTY")
    if sheet == "Local & Others":
        return _group_table(built, "local", "Local & Others", "Pending Orders")
    if sheet == "RISKS":
        return _group_table(built, "risks", "RISKS", "IN TRANSIT QTY")
    if sheet == "Milk and crates":
        return _group_table(built, "milk", "Milk and crates", "IN TRANSIT QTY")
    if sheet == "Others":
        return _others(built)
    if sheet == "Stock":
        return _stock_table(built)
    if sheet == "Consmptn Dump":
        return _pivot(built, "by_item_m", "item_no", "Item No.", "Consmptn Dump",
                      "Item by month absolute quantity for Consumption, Negative Adjmt. and Positive Adjmt. "
                      "This is the pivot MRP reads. The 70,000-row ledger is not listed line by line.",
                      "Entry types: Consumption, Negative Adjmt., Positive Adjmt.")
    if sheet == "Pivot Tables":
        table = _pivot(built, "by_item_m", "item_no", "Item No.", "Pivot Tables",
                       "Average = SUM of months with quantity / COUNT of those months. Maximum = MAX of the months.",
                       "")
        keep = {"item_no", "description", "average", "maximum"}
        table["columns"] = [c for c in table["columns"] if c["field"] in keep]
        table["rows"] = [{f: r.get(f) for f in keep} for r in table["rows"]]
        return table
    if sheet == "Sheet6":
        return _pivot(built, "by_desc_m", "description", "Description", "Sheet6",
                      "Consumption entries by description and month: the pivot on Sheet6 (Sum of ABS Quantity).",
                      "Entry type: Consumption.")
    if sheet == "All Purchase Plans":
        return _grid_table("All Purchase Plans", pack.get("purchase_plans"), built["by_desc"],
                           "Planner working sheet. Cells that look up MRP now show the current MRP figure.")
    if sheet == "Local Purchase":
        return _grid_table("Local Purchase", pack.get("local_purchase"), built["by_desc"],
                           "Descriptions looked up into MRP for inventory and cover. Order quantities are planner inputs.")
    if sheet == "RM MATRIX":
        return _grid_table("RM MATRIX", pack.get("rm_matrix"), None,
                           "Resin recipe. Row 134 is the projected requirement MRP reads. Weights, projections and factors are planner inputs.")
    if sheet == "Projected Sales ":
        return _grid_table("Projected Sales", pack.get("projected_sales"), None,
                           "Planner sales forecast. Not a Business Central figure.")
    if sheet == "Sales FY26":
        return _grid_table("Sales FY26", pack.get("sales_fy26"), None,
                           "Planner FY26 sales, including lookups into Projected Sales. Not a Business Central figure.")
    if sheet == "Sales rev. forecast mar.":
        return _grid_table("Sales rev. forecast", pack.get("sales_rev"), None,
                           "Planner revenue forecast. Not a Business Central figure.")
    if sheet == "Intransit":
        table = _grid_table("Intransit", pack.get("intransit"), None,
                            "Shipment tracker kept from the workbook. Business Central does not hold BL, vessel or ETA. "
                            "The second quantity column repeats the kgs column.")
        for row in table["rows"]:
            if row.get("QUANTITY_KGS_U") in (None, "") and row.get("QUANTITY_KGS") not in (None, ""):
                row["QUANTITY_KGS_U"] = row["QUANTITY_KGS"]
        return table
    if sheet == "my sheet":
        return _grid_table("my sheet", pack.get("my_sheet"), None, "Lookup behind the in-transit tracker's BC material name.")
    if sheet == "Alternatives":
        return _alternatives(built)
    if sheet == "KEY":
        return _key(built)
    if sheet == "Sheet1":
        return _sheet1(built)
    return _table(sheet, [], [_col("note", "Note", "text")])


def _valued(qty, item, costs):
    return float(qty or 0) * float(costs.get(item) or 0)


def _headline(built, rows, sheet):
    from . import views as V
    costs = built.get("costs") or {}
    planning = [r for r in rows if r.get("item_no") and ("inventory" in r or "order_qty" in r or "cover_flag" in r)]
    if planning:
        return _planning_headline(built, planning, costs, V)
    stock_rows = [r for r in rows if r.get("item_no") and "sum_rm" in r]
    if stock_rows:
        return _stock_headline(built, stock_rows, costs, V, sheet)
    return _grid_headline(rows, sheet, V)


def _planning_headline(built, rows, costs, V):
    critical_rows = [r for r in rows if r.get("cover_flag") == "CRITICAL" and r.get("status") == "ACTIVE"]
    slob_rows = [r for r in rows if r.get("status") == "SLOB"]
    okay = sum(1 for r in rows if r.get("status") == "ACTIVE" and r.get("cover_flag") == "OKAY")
    stock_qty = sum(r.get("inventory") or 0 for r in rows)
    stock_val = sum(_valued(r.get("inventory"), r.get("item_no"), costs) for r in rows)
    transit_qty = sum(r.get("in_transit") or 0 for r in rows)
    transit_val = sum(_valued(r.get("in_transit"), r.get("item_no"), costs) for r in rows)
    order_qty = sum(r.get("order_qty") or 0 for r in rows)
    order_val = sum(_valued(r.get("order_qty"), r.get("item_no"), costs) for r in rows)
    ids = {r["item_no"] for r in rows}
    frame = built["by_item_m"]
    months = built["months"]
    monthly = []
    monthly_val = []
    if not frame.empty and months and ids:
        sub = frame[frame.item_no.isin(ids)]
        totals = sub.groupby("month").abs_qty.sum()
        monthly = [float(totals.get(m, 0) or 0) for m in months]
        valued = sub.assign(_v=sub.abs_qty * sub.item_no.map(lambda i: float(costs.get(i) or 0)))
        by_m = valued.groupby("month")._v.sum()
        monthly_val = [float(by_m.get(m, 0) or 0) for m in months]
    labels = months or ["—"]
    loc_qty = {name: 0.0 for name in SUM_RM}
    loc_val = {name: 0.0 for name in SUM_RM}
    for item in ids:
        row = built["stock"].get(item) or {}
        cost = float(costs.get(item) or 0)
        for name in SUM_RM:
            qty = float(row.get(name, 0) or 0)
            loc_qty[name] += qty
            loc_val[name] += qty * cost
    cover_labels = ["Critical (≤ 29 days)", "Okay", "SLOB"]
    cover_qty = [len(critical_rows), okay, len(slob_rows)]
    cover_val = [
        sum(_valued(r.get("inventory"), r.get("item_no"), costs) for r in critical_rows),
        sum(_valued(r.get("inventory"), r.get("item_no"), costs) for r in rows if r.get("status") == "ACTIVE" and r.get("cover_flag") == "OKAY"),
        sum(_valued(r.get("inventory"), r.get("item_no"), costs) for r in slob_rows),
    ]
    top = sorted(rows, key=lambda r: r.get("order_qty") or 0, reverse=True)[:12]
    top = [r for r in top if (r.get("order_qty") or 0) > 0]
    kpis = [
        V.kpi("RM stock (SUM RM)", stock_val, "money", f"{len(rows)} items in this filter",
              spark=monthly_val[-6:] or monthly[-6:] or None, qty=stock_qty),
        V.kpi("Cover ≤ 29 days", len(critical_rows), "qty", f"of {len(rows)} items", status="critical",
              pct=V.pct(len(critical_rows), len(rows))),
        V.kpi("SLOB", len(slob_rows), "qty", "Average consumption is zero", status="serious",
              pct=V.pct(len(slob_rows), len(rows))),
        V.kpi("Open PO quantity", transit_val, "money", "IN TRANSIT QTY, from open purchase orders", qty=transit_qty),
        V.kpi("Suggested order qty", order_val, "money", "Rounded up to 25,500, the workbook's order multiple", qty=order_qty),
    ]
    charts = [
        V.chart("area", "Monthly consumption of MRP items", labels,
                [("Value", "seq", monthly_val or [0])], "UGX", span=2,
                extra={"quantity": monthly or [0]}),
        V.chart("doughnut", "MRP cover position", cover_labels,
                [("Value", "seq", cover_val)], "UGX", colors=["crit", "good", "warning"],
                extra={"quantity": cover_qty}),
        V.chart("hbar", "Largest suggested orders",
                [f"{r['item_no']} · {str(r.get('description') or '')[:28]}" for r in top] or ["—"],
                [("Value", "serious", [_valued(r.get("order_qty"), r.get("item_no"), costs) for r in top] or [0])],
                "UGX", span=2, height=360,
                extra={"quantity": [r.get("order_qty") or 0 for r in top] or [0],
                       "item_nos": [r["item_no"] for r in top]}),
        V.chart("bar", "Stock on the SUM RM locations", list(loc_qty),
                [("Value", "RM", list(loc_val.values()))], "UGX",
                extra={"quantity": list(loc_qty.values())}),
    ]
    return kpis, charts


def _stock_headline(built, rows, costs, V, sheet):
    qty = sum(r.get("sum_rm") or 0 for r in rows)
    val = sum(_valued(r.get("sum_rm"), r.get("item_no"), costs) for r in rows)
    top = sorted(rows, key=lambda r: r.get("sum_rm") or 0, reverse=True)[:12]
    kpis = [
        V.kpi("RM stock (SUM RM)", val, "money", f"{len(rows)} items on {sheet}", qty=qty),
        V.kpi("Cover ≤ 29 days", 0, "qty", "Cover is calculated on the MRP sheet"),
        V.kpi("SLOB", 0, "qty", "Cover is calculated on the MRP sheet"),
        V.kpi("Open PO quantity", 0, "money", "Open PO quantity is on the MRP sheet", qty=0),
        V.kpi("Suggested order qty", 0, "money", "Order quantity is on the MRP sheet", qty=0),
    ]
    charts = [
        V.chart("hbar", f"Largest SUM RM on {sheet}",
                [f"{r['item_no']} · {str(r.get('description') or '')[:28]}" for r in top] or ["—"],
                [("Value", "RM", [_valued(r.get("sum_rm"), r.get("item_no"), costs) for r in top] or [0])],
                "UGX", span=2, height=360,
                extra={"quantity": [r.get("sum_rm") or 0 for r in top] or [0],
                       "item_nos": [r["item_no"] for r in top]}),
    ]
    return kpis, charts


def _grid_headline(rows, sheet, V):
    numeric = []
    sample = rows[:40]
    for row in sample:
        for key, val in row.items():
            if key in numeric or key == "item_no":
                continue
            if isinstance(val, (int, float)):
                numeric.append(key)
    sums = []
    for key in numeric[:8]:
        total = sum(float(r.get(key) or 0) for r in rows)
        if total:
            sums.append((key, total))
    kpis = [V.kpi("RM stock (SUM RM)", len(rows), "qty", f"{len(rows)} rows on {sheet}")]
    for key, total in sums[:4]:
        kpis.append(V.kpi(key, total, "qty", sheet))
    while len(kpis) < 5:
        kpis.append(V.kpi("Rows", len(rows), "qty", sheet))
    labels = [k for k, _ in sums] or ["Rows"]
    data = [v for _, v in sums] or [len(rows)]
    charts = [V.chart("bar", sheet, labels, [("Quantity", "seq", data)], "qty", span=2)]
    return kpis[:5], charts


def _filtered_rows(built, rows, filters):
    cls = filters.get("item_class")
    site = filters.get("site")
    loc = (filters.get("location") or "").strip()
    q = (filters.get("q") or "").strip().lower()
    if not (cls or site or loc or q):
        return rows
    kept = []
    for r in rows:
        item = r.get("item_no")
        if cls or site or loc:
            if not item:
                continue
            if cls and built["classes"].get(item) != cls:
                continue
            if site and site not in built["at_site"].get(item, ()):
                continue
            if loc and loc not in built["at_loc"].get(item, ()):
                continue
        if q and not any(q in str(v).lower() for v in r.values() if v not in (None, "")):
            continue
        kept.append(r)
    return kept


def render(model, filters=None):
    filters = filters or {}
    built = compute(model)
    sheet = filters.get("sheet") or "MRP"
    known = {sid for _g, _title, sid in SHEETS}
    if sheet not in known:
        sheet = "MRP"
    table = sheet_table(built, sheet)
    rows = _filtered_rows(built, table["rows"], filters)
    kpis, charts = _headline(built, rows, sheet)
    table = dict(table)
    table["rows"] = rows
    table["total_rows"] = len(rows)
    return {
        "id": "rm_planning", "no": "12", "title": "Raw Material Planning",
        "question": "",
        "kpis": kpis, "insights": [], "charts": charts,
        "tables": [table], "sheet": sheet,
        "rm_sheets": [{"group": g, "id": sid, "title": title} for g, title, sid in SHEETS],
    }
