"""Reporting: standard report catalogue, custom report builder, Excel export, history and schedules.

Every report is built from the same governed model as the dashboards, so exported figures match the screens.
"""
import io
import json
import re
import uuid
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd

from . import config as C
from . import views as V
from .bc_client import DATA_DIR

TZ = ZoneInfo("Africa/Kampala")
HISTORY = DATA_DIR / "report_history.json"
FILES = DATA_DIR / "report_files"
CUSTOM = DATA_DIR / "custom_reports.json"
SCHEDULES = DATA_DIR / "schedules.json"


def now():
    return datetime.now(TZ)


def _read(p, default):
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError):
        return default


def _write(p, obj):
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(obj, indent=2, default=str), encoding="utf-8")


# ---------------------------------------------------------------- catalogue
CATALOGUE = [
    # id, name, category, description, view (or None for custom builders)
    ("stock_status", "Stock Status Report", "Inventory", "Stock on hand and value by item for RM, WIP and FG.", "inventory_position"),
    ("stock_by_location", "Stock by Location Report", "Inventory", "Stock value by site and location, with item detail.", "inventory_by_location"),
    ("stock_valuation", "Stock Valuation Summary", "Inventory", "Stock value by site, class and posting group.", None),
    ("rm_cover", "Raw Material Consumption & Cover", "Cover", "RM consumption rates and production days of cover.", "rm_consumption_cover"),
    ("wip_cover", "WIP Cover Report", "Cover", "WIP stock and complete parent units it supports.", "wip_cover"),
    ("fg_cover", "Finished Goods Cover Report", "Cover", "FG stock against expected daily demand.", "fg_cover"),
    ("open_po", "Open & Overdue Purchase Orders", "Supply", "Outstanding purchase lines, expected dates and lateness.", "in_transit_overdue"),
    ("vendor_performance", "Vendor Delivery Performance", "Supply", "Overdue lines, value and days late by vendor.", None),
    ("stockout_exposure", "Stockout Exposure Report", "Supply", "Items projected to run out before replenishment.", "projection_stockout"),
    ("inventory_risk", "Inventory Risk Register", "Risk", "Items needing attention by cover and incoming supply.", "inventory_risk"),
    ("rebalancing", "Location Rebalancing Report", "Risk", "Site shortages and suggested transfers from surplus sites.", "location_risk"),
    ("production_readiness", "Production Readiness Report", "Production", "What can be produced and the constraining components.", "bom_readiness"),
    ("exceptions", "Supply Chain Exceptions Report", "Management", "Ranked list of everything needing management attention.", "exceptions"),
    ("kpi_summary", "Management KPI Summary", "Management", "Headline KPIs from every view on one page.", "overview"),
    ("movement_trends", "Weekly Movement Trends", "Management", "Weekly consumption, output, sales and receipts by class (value).", None),
    ("data_quality", "Data Quality & Reconciliation", "Governance", "Reconciliation to BC, stale POs and data gaps found in profiling.", None),
]
CAT_IDS = {c[0] for c in CATALOGUE}


def catalogue():
    out = [{"id": i, "name": n, "category": c, "description": d, "custom": False} for i, n, c, d, _ in CATALOGUE]
    for r in _read(CUSTOM, []):
        out.append({"id": "custom:" + r["id"], "name": r["name"], "category": "My reports",
                    "description": r.get("description") or f"Custom report on {DATASETS[r['spec']['dataset']]['label']}",
                    "custom": True, "spec": r["spec"], "created_by": r.get("created_by")})
    return out


def _filters(params):
    return {"item_class": params.get("item_class") or None, "site": params.get("site") or None,
            "location": None, "q": params.get("q") or None}


def run_report(m, report_id, params=None):
    params = params or {}
    tok = V.ROW_LIMIT.set(1_000_000)
    try:
        if report_id.startswith("custom:"):
            r = next((x for x in _read(CUSTOM, []) if "custom:" + x["id"] == report_id), None)
            if not r:
                raise KeyError("Unknown custom report")
            spec = {**r["spec"], "filters": {**r["spec"].get("filters", {}), **{k: v for k, v in params.items() if v}}}
            body = run_custom(m, spec)
            meta = {"name": r["name"], "category": "My reports", "description": r.get("description", "")}
        else:
            row = next((c for c in CATALOGUE if c[0] == report_id), None)
            if not row:
                raise KeyError("Unknown report")
            _, name, cat, desc, view = row
            meta = {"name": name, "category": cat, "description": desc}
            if view:
                out = V.render(m, view, _filters(params))
                body = {"summary": out["kpis"], "charts": out["charts"], "sections": out["tables"]}
            else:
                body = globals()["_r_" + report_id](m, _filters(params))
    finally:
        V.ROW_LIMIT.reset(tok)
    n = m["notes"]
    return {"id": report_id, **meta, **body, "params": {k: v for k, v in params.items() if v},
            "company": n["company"], "environment": n["environment"], "reporting_date": n["reporting_date"],
            "data_loaded": n["extracted_at"], "generated_at": now().strftime("%Y-%m-%d %H:%M")}


def _r_stock_valuation(m, f):
    sl = V.apply_filters(m["stock_loc"], f)
    sl = sl[(sl.item_class != "OTHER") & ~sl.in_transit]
    site = sl.pivot_table(index="site", columns="item_class", values="stock_value", aggfunc="sum",
                          fill_value=0).reindex(columns=V.CLS, fill_value=0)
    site["total"] = site.sum(axis=1)
    site = site.sort_values("total", ascending=False).reset_index()
    pg = sl.groupby(["item_class", "posting_group"]).agg(items=("item_no", "nunique"), qty=("stock_qty", "sum"),
                                                         value=("stock_value", "sum")).reset_index()
    pg = pg.sort_values("value", ascending=False)
    tot = sl.groupby("item_class").stock_value.sum().reindex(V.CLS).fillna(0)
    return {"summary": [V.kpi("Total value", tot.sum(), "money")] +
                       [V.kpi(C.CLASS_LABELS[c], tot[c], "money") for c in V.CLS],
            "charts": [V.chart("bar", "Stock value by site", site.site.tolist(),
                               [(C.CLASS_LABELS[c], c, site[c].tolist()) for c in V.CLS], "UGX", stacked=True)],
            "sections": [V.table("Value by site", site, [("site", "Site", "text"), ("RM", "RM (UGX)", "money"),
                                                         ("WIP", "WIP (UGX)", "money"), ("FG", "FG (UGX)", "money"),
                                                         ("total", "Total (UGX)", "money")], drill=None),
                         V.table("Value by posting group", pg, [("item_class", "Class", "text"),
                                                                ("posting_group", "Posting group", "text"),
                                                                ("items", "Items", "qty"), ("qty", "Quantity", "qty"),
                                                                ("value", "Value (UGX)", "money")], drill=None)]}


def _r_vendor_performance(m, f):
    pl = V.apply_filters(m["purchase_lines"], f)
    g = pl.groupby(["vendor_no", "vendor_name"])
    v = pd.DataFrame({
        "open_lines": g.size(), "open_value": g.outstanding_value.sum(),
        "overdue_lines": g.apply(lambda d: int((d.overdue & ~d.stale).sum())),
        "overdue_value": g.apply(lambda d: float(d[d.overdue & ~d.stale].outstanding_value.sum())),
        "stale_lines": g.apply(lambda d: int(d.stale.sum())),
        "avg_days_late": g.apply(lambda d: float(d[d.overdue].days_overdue.mean()) if d.overdue.any() else 0.0),
        "max_days_late": g.days_overdue.max().clip(lower=0),
    }).reset_index()
    v["on_time_pct"] = np.where(v.open_lines > 0, 100 * (1 - (v.overdue_lines + v.stale_lines) / v.open_lines), np.nan)
    v = v.sort_values("overdue_value", ascending=False)
    top = v.head(12)
    return {"summary": [V.kpi("Vendors with open orders", len(v), "qty"),
                        V.kpi("Vendors with overdue lines", int((v.overdue_lines > 0).sum()), "qty", status="serious"),
                        V.kpi("Overdue value (≤90d)", v.overdue_value.sum(), "money", status="serious"),
                        V.kpi("Stale lines (>90d)", int(v.stale_lines.sum()), "qty", status="warning")],
            "charts": [V.chart("hbar", "Overdue value by vendor (top 12)", top.vendor_name.str.strip().tolist(),
                               [("Overdue value", "serious", top.overdue_value.tolist())], "UGX")],
            "sections": [V.table("Vendor delivery performance", v, [
                ("vendor_no", "Vendor no.", "text"), ("vendor_name", "Vendor", "text"), ("open_lines", "Open lines", "qty"),
                ("open_value", "Open value (UGX)", "money"), ("overdue_lines", "Overdue lines", "qty"),
                ("overdue_value", "Overdue value (UGX)", "money"), ("stale_lines", "Stale lines", "qty"),
                ("avg_days_late", "Avg days late", "days"), ("max_days_late", "Max days late", "qty"),
                ("on_time_pct", "Not overdue %", "days")], drill=None)]}


def _r_movement_trends(m, f):
    w = m["weekly"]
    if f.get("item_class"):
        w = w[w.item_class == f["item_class"]]
    t = w.groupby("week")[["consumed_value", "output_value", "sold_value", "received_value"]].sum().reset_index()
    labels = [str(x) for x in t.week]
    return {"summary": [V.kpi("Weeks covered", len(t), "qty"),
                        V.kpi("Avg weekly consumption", t.consumed_value.mean(), "money"),
                        V.kpi("Avg weekly sales (cost)", t.sold_value.mean(), "money"),
                        V.kpi("Avg weekly receipts", t.received_value.mean(), "money")],
            "charts": [V.chart("line", "Weekly flow at cost (UGX)", labels, [
                ("Consumption", "seq", t.consumed_value.tolist()), ("Output", "good", t.output_value.tolist()),
                ("Sales", "FG", t.sold_value.tolist()), ("Purchase receipts", "serious", t.received_value.tolist())], "UGX")],
            "sections": [V.table("Weekly flow (value at cost)", t.assign(week=t.week.astype(str)), [
                ("week", "Week starting", "text"), ("consumed_value", "Consumption (UGX)", "money"),
                ("output_value", "Output (UGX)", "money"), ("sold_value", "Sales at cost (UGX)", "money"),
                ("received_value", "Purchase receipts (UGX)", "money")], drill=None),
                V.table("Weekly flow by class", w.assign(week=w.week.astype(str)).sort_values(["week", "item_class"]), [
                    ("week", "Week starting", "text"), ("item_class", "Class", "text"),
                    ("consumed_value", "Consumption (UGX)", "money"), ("output_value", "Output (UGX)", "money"),
                    ("sold_value", "Sales at cost (UGX)", "money"), ("received_value", "Receipts (UGX)", "money")],
                    drill=None)]}


def _r_data_quality(m, f):
    n = m["notes"]
    r = n.get("reconciliation") or {}
    pl = m["purchase_lines"]
    stale = pl[pl.stale].sort_values("days_overdue", ascending=False)
    mism = pd.DataFrame(r.get("mismatches") or [], columns=["item_no", "bc", "pulse"])
    if len(mism):
        mism["difference"] = mism.pulse - mism.bc
    pct = 100 * r["items_matching"] / max(r["items_compared"], 1) if r else None
    checks = pd.DataFrame([
        ("Stock reconciles to BC item inventory", f"{r.get('items_matching', 0):,} of {r.get('items_compared', 0):,} items",
         "OK" if pct and pct >= 99 else "Review"),
        ("Outstanding PO lines past expected date", f"{n['purchase_lines_overdue']} of {n['purchase_lines_outstanding']}",
         "Review" if n["purchase_lines_overdue"] else "OK"),
        ("Stale PO lines (>90 days overdue)", str(n["purchase_lines_stale"]), "Review" if n["purchase_lines_stale"] else "OK"),
        ("Certified BOM lines with Quantity per = 0", f"{n['bom_lines_zero_qty']:,} of {n['bom_lines_certified']:,}",
         "Review" if n["bom_lines_zero_qty"] else "OK"),
        ("Items with stock outside RM/WIP/FG", f"{n['items_unclassified_with_stock']:,}", "Info"),
    ], columns=["check", "result", "status"])
    return {"summary": [V.kpi("Reconciliation match", pct, "days", "% of items matching BC exactly"),
                        V.kpi("Stale PO lines", n["purchase_lines_stale"], "qty", status="warning"),
                        V.kpi("BOM lines with qty 0", n["bom_lines_zero_qty"], "qty", status="warning"),
                        V.kpi("Open ledger entries loaded", n["open_ledger_entries"], "qty")],
            "charts": [],
            "sections": [V.table("Data checks", checks, [("check", "Check", "text"), ("result", "Result", "text"),
                                                         ("status", "Status", "status")], drill=None),
                         V.table("Reconciliation differences", mism, [("item_no", "Item", "text"),
                                 ("bc", "BC inventory", "qty"), ("pulse", "Pulse stock", "qty"),
                                 ("difference", "Difference", "qty")]),
                         V.table("Stale purchase lines", stale, [
                             ("po_no", "PO", "text"), ("vendor_name", "Vendor", "text"), ("item_no", "Item", "text"),
                             ("description", "Description", "text"), ("outstanding_qty", "Outstanding", "qty"),
                             ("expected_date", "Expected", "date"), ("days_overdue", "Days late", "qty"),
                             ("outstanding_value", "Value (UGX)", "money")])]}


# ---------------------------------------------------------------- custom builder
def _ds_items(m):
    return m["item"]


DATASETS = {
    "items": {"label": "Item position, cover & risk", "get": lambda m: m["item"], "columns": [
        "item_no", "description", "item_class", "abc", "posting_group", "uom", "stock_qty", "stock_value", "daily_usage",
        "cover_days", "incoming_qty", "overdue_qty", "cover_status", "supply_status", "risk_level", "stockout_date",
        "unit_cost", "consumed", "sold", "output", "received"]},
    "stock_by_location": {"label": "Stock by item and location", "get": lambda m: m["stock_loc"], "columns": [
        "item_no", "description", "item_class", "posting_group", "site", "location", "location_name", "uom",
        "stock_qty", "stock_value"]},
    "purchase_lines": {"label": "Open purchase order lines", "get": lambda m: m["purchase_lines"], "columns": [
        "po_no", "vendor_no", "vendor_name", "item_no", "description", "item_class", "site", "location",
        "ordered_qty", "received_qty", "outstanding_qty", "expected_date", "days_overdue", "receipt_status",
        "outstanding_value"]},
    "exceptions": {"label": "Exceptions", "get": lambda m: m["exceptions"], "columns": [
        "severity", "exception_type", "item_no", "description", "item_class", "location", "detail", "value_at_stake"]},
    "readiness": {"label": "Production readiness",
                  "get": lambda m: m["readiness"].rename(columns={"parent_description": "description",
                                                                  "parent_class": "item_class"}),
                  "columns": ["item_no", "description", "item_class", "daily_output", "can_make_now",
                              "production_days_supported", "can_make_with_incoming", "limiting_component",
                              "limiting_description", "readiness"]},
    "location_risk": {"label": "Location shortages & transfers", "get": lambda m: m["location_risk"], "columns": [
        "item_no", "description", "item_class", "site", "stock_qty", "daily_usage", "cover_days", "need_qty",
        "donor_site", "suggested_transfer", "rebalance"]},
    "weekly_flow": {"label": "Weekly movement flow (value)", "get": lambda m: m["weekly"].assign(week=m["weekly"].week.astype(str)),
                    "columns": ["week", "item_class", "consumed_value", "output_value", "sold_value", "received_value"]},
}
AGGS = {"sum": "Sum", "avg": "Average", "min": "Min", "max": "Max", "count": "Count"}


def _kind(col, s):
    if re.search(r"value|cost|amount", col):
        return "money"
    if re.search(r"cover|days", col) and pd.api.types.is_numeric_dtype(s):
        return "days"
    if pd.api.types.is_numeric_dtype(s) and not pd.api.types.is_bool_dtype(s):
        return "qty"
    if re.search(r"date", col):
        return "date"
    if re.search(r"status|risk_level|severity|readiness|rebalance", col):
        return "status"
    return "text"


def _label(col):
    return col.replace("_", " ").replace("qty", "qty").capitalize().replace("Po no", "PO").replace("Uom", "UoM")


def dataset_meta(m):
    out = {}
    for k, d in DATASETS.items():
        df = d["get"](m)
        out[k] = {"label": d["label"], "columns": [{"field": c, "label": _label(c), "kind": _kind(c, df[c])}
                                                   for c in d["columns"] if c in df]}
    return out


def run_custom(m, spec):
    ds = DATASETS.get(spec.get("dataset"))
    if not ds:
        raise KeyError("Unknown dataset")
    df = V.apply_filters(ds["get"](m), spec.get("filters") or {})
    allowed = [c for c in ds["columns"] if c in df]
    for flt in spec.get("conditions") or []:          # simple column conditions
        c, op, val = flt.get("col"), flt.get("op"), flt.get("value")
        if c not in allowed or val in (None, ""):
            continue
        s = df[c]
        if pd.api.types.is_numeric_dtype(s):
            v = float(val)
            df = df[{"<": s < v, "<=": s <= v, ">": s > v, ">=": s >= v, "=": s == v, "!=": s != v}.get(op, s == s)]
        else:
            ss = s.astype(str).str.lower()
            df = df[ss.str.contains(str(val).lower(), regex=False) if op == "contains" else
                    (ss == str(val).lower()) if op == "=" else (ss != str(val).lower())]
    group = spec.get("group_by") if spec.get("group_by") in allowed else None
    charts = []
    if group:
        metrics = [x for x in (spec.get("metrics") or []) if x.get("col") in allowed and x.get("agg") in AGGS]
        agg = {"rows": (group, "size")}
        for x in metrics:
            agg[f"{x['agg']}_{x['col']}"] = (x["col"], "mean" if x["agg"] == "avg" else x["agg"])
        out = df.groupby(group, dropna=False).agg(**agg).reset_index()
        cols = [(group, _label(group), "text"), ("rows", "Rows", "qty")] + [
            (f"{x['agg']}_{x['col']}", f"{AGGS[x['agg']]} of {_label(x['col']).lower()}",
             "qty" if x["agg"] == "count" else _kind(x["col"], df[x["col"]])) for x in metrics]
        first = cols[2][0] if len(cols) > 2 else "rows"
        out = out.sort_values(first, ascending=False)
        top = out.head(15)
        charts.append(V.chart("hbar", f"{cols[2][1] if len(cols) > 2 else 'Rows'} by {_label(group).lower()}",
                              top[group].astype(str).tolist(), [(cols[2][1] if len(cols) > 2 else "Rows", "seq",
                                                                  top[first].tolist())],
                              "UGX" if len(cols) > 2 and cols[2][2] == "money" else None))
        drill = None
    else:
        sel = [c for c in (spec.get("columns") or allowed[:8]) if c in allowed] or allowed[:8]
        out = df[sel]
        cols = [(c, _label(c), _kind(c, df[c])) for c in sel]
        drill = "item_no" if "item_no" in sel else None
    srt = spec.get("sort") or {}
    if srt.get("col") in out:
        out = out.sort_values(srt["col"], ascending=srt.get("dir") == "asc")
    if spec.get("limit"):
        out = out.head(int(spec["limit"]))
    t = V.table(spec.get("title") or ds["label"], out, cols, drill=drill)
    num = [c for c, _, k in cols if k in ("qty", "money")]
    summary = [V.kpi("Groups" if group else "Rows", len(out), "qty")] + [
        V.kpi("Rows in groups" if c == "rows" else f"Total {lbl.lower()}", out[c].sum(), k)
        for c, lbl, k in cols if c in num and not (k == "days")][:3]
    return {"summary": summary, "charts": charts, "sections": [t]}


def save_custom(name, spec, description="", user=None, rid=None):
    allr = _read(CUSTOM, [])
    if rid:
        allr = [r for r in allr if r["id"] != rid]
    r = {"id": rid or uuid.uuid4().hex[:8], "name": name, "description": description, "spec": spec,
         "created_by": user, "created": now().isoformat(timespec="seconds")}
    _write(CUSTOM, allr + [r])
    return r


def delete_custom(rid):
    _write(CUSTOM, [r for r in _read(CUSTOM, []) if "custom:" + r["id"] != rid and r["id"] != rid])


# ---------------------------------------------------------------- export
FMT = {"money": "#,##0", "qty": "#,##0.##", "days": "#,##0.0"}


def to_xlsx(rep):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    navy, white = PatternFill("solid", fgColor="10243E"), Font(color="FFFFFF", bold=True)
    ws["A1"] = rep["name"]
    ws["A1"].font = Font(size=16, bold=True, color="10243E")
    meta = [("Company", rep["company"]), ("Environment", rep["environment"]),
            ("Reporting date", rep["reporting_date"]), ("Data loaded", rep["data_loaded"]),
            ("Generated", rep["generated_at"] + " (EAT)"),
            ("Filters", ", ".join(f"{k}={v}" for k, v in rep["params"].items()) or "None"),
            ("Description", rep.get("description", ""))]
    for i, (k, v) in enumerate(meta, start=3):
        ws.cell(i, 1, k).font = Font(bold=True)
        ws.cell(i, 2, v)
    r0 = len(meta) + 4
    ws.cell(r0, 1, "Measure").fill = navy
    ws.cell(r0, 1).font = white
    ws.cell(r0, 2, "Value").fill = navy
    ws.cell(r0, 2).font = white
    for i, k in enumerate(rep.get("summary", []), start=r0 + 1):
        ws.cell(i, 1, k["label"])
        c = ws.cell(i, 2, k["value"])
        if k["kind"] in FMT:
            c.number_format = FMT[k["kind"]]
        if k.get("note"):
            ws.cell(i, 3, k["note"]).font = Font(italic=True, color="7A7974")
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 42
    ws.column_dimensions["C"].width = 40
    ws.cell(r0 + len(rep.get("summary", [])) + 2, 1,
            "Source: Microsoft Dynamics 365 Business Central via Pulse (read-only). Figures match the Pulse dashboards."
            ).font = Font(italic=True, color="7A7974")

    used = set()
    for sec in rep.get("sections", []):
        name = re.sub(r"[\[\]\*\?/\\:]", "", sec["title"])[:31] or "Data"
        base, n = name, 2
        while name in used:
            name = f"{base[:28]} {n}"
            n += 1
        used.add(name)
        sh = wb.create_sheet(name)
        for j, col in enumerate(sec["columns"], start=1):
            c = sh.cell(1, j, col["label"])
            c.fill, c.font = navy, white
            c.alignment = Alignment(vertical="center", wrap_text=True)
        for i, row in enumerate(sec["rows"], start=2):
            for j, col in enumerate(sec["columns"], start=1):
                c = sh.cell(i, j, row.get(col["field"]))
                if col["kind"] in FMT:
                    c.number_format = FMT[col["kind"]]
        for j, col in enumerate(sec["columns"], start=1):
            width = max([len(str(col["label"]))] + [len(str(r.get(col["field"]) or "")) for r in sec["rows"][:300]])
            sh.column_dimensions[get_column_letter(j)].width = min(max(10, width + 2), 60)
        sh.freeze_panes = "A2"
        if sec["rows"]:
            sh.auto_filter.ref = f"A1:{get_column_letter(len(sec['columns']))}{len(sec['rows']) + 1}"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def to_csv(rep, section=0):
    sec = rep["sections"][section]
    df = pd.DataFrame(sec["rows"], columns=[c["field"] for c in sec["columns"]])
    df.columns = [c["label"] for c in sec["columns"]]
    return df.to_csv(index=False).encode("utf-8-sig")


# ---------------------------------------------------------------- history
def generate(m, report_id, params=None, user=None, trigger="manual"):
    rep = run_report(m, report_id, params)
    hid = uuid.uuid4().hex[:10]
    FILES.mkdir(parents=True, exist_ok=True)
    (FILES / f"{hid}.xlsx").write_bytes(to_xlsx(rep))
    entry = {"id": hid, "report_id": report_id, "name": rep["name"], "params": rep["params"],
             "generated_at": rep["generated_at"], "by": user or "system", "trigger": trigger,
             "rows": sum(len(s["rows"]) for s in rep["sections"]), "data_loaded": rep["data_loaded"]}
    hist = _read(HISTORY, [])
    _write(HISTORY, ([entry] + hist)[:500])
    return entry


def history():
    return _read(HISTORY, [])


def history_file(hid):
    e = next((h for h in history() if h["id"] == hid), None)
    if not e:
        return None, None
    ext = (e.get("ext") or "xlsx").lower().lstrip(".")
    if ext not in ("xlsx", "xlsm", "docx", "pptx", "csv"):
        ext = "xlsx"
    p = FILES / f"{hid}.{ext}"
    if not p.exists():
        alt = FILES / f"{hid}.xlsx"
        p = alt if alt.exists() else p
    return (e, p) if p.exists() else (None, None)


# ---------------------------------------------------------------- schedules
DEFAULT_SCHED = {"refresh": {"enabled": True, "time": "06:00", "days": [0, 1, 2, 3, 4, 5], "last_run": None},
                 "reports": []}


def schedules():
    s = _read(SCHEDULES, None)
    if s is None:   # first start: don't fire today's refresh immediately on deploy
        s = json.loads(json.dumps(DEFAULT_SCHED))
        s["refresh"]["last_run"] = now().isoformat(timespec="minutes")
        _write(SCHEDULES, s)
    s.setdefault("refresh", DEFAULT_SCHED["refresh"])
    s.setdefault("reports", [])
    return s


def save_schedules(s):
    _write(SCHEDULES, s)


def next_run(sch, after=None):
    after = after or now()
    hh, mm = (int(x) for x in (sch.get("time") or "07:00").split(":")[:2])
    freq = sch.get("frequency", "daily")
    if freq == "custom" and (sch.get("unit") or "days") == "hours":
        every = max(1, min(int(sch.get("every") or 1), 24 * 30))
        return after + timedelta(hours=every)
    step = 1
    if freq == "custom":
        step = max(1, min(int(sch.get("every") or 1), 365))
    cand = after.replace(hour=hh, minute=mm, second=0, microsecond=0)
    if cand <= after:
        cand += timedelta(days=step)
    for _ in range(400):
        if cand > after and (freq in ("daily", "custom")
                             or (freq == "weekdays" and cand.weekday() < 5)
                             or (freq == "weekly" and cand.weekday() == int(sch.get("weekday", 0)))
                             or (freq == "monthly" and cand.day == int(sch.get("monthday", 1)))):
            return cand
        cand += timedelta(days=1 if freq != "custom" else step)
    return None


def add_schedule(report_id, params, frequency, time_, weekday=0, monthday=1, user=None, recipients=None, every=None, unit=None):
    from . import mail
    s = schedules()
    sch = {"id": uuid.uuid4().hex[:8], "report_id": report_id, "params": params or {}, "frequency": frequency,
           "time": time_, "weekday": int(weekday), "monthday": int(monthday), "enabled": True,
           "created_by": user, "last_run": None, "recipients": mail.clean_recipients(recipients),
           "every": int(every) if frequency == "custom" and every else None,
           "unit": unit if frequency == "custom" and unit in ("hours", "days") else None}
    nxt = next_run(sch)
    sch["next_run"] = nxt.isoformat(timespec="minutes") if nxt else None
    s["reports"].append(sch)
    save_schedules(s)
    return sch
