"""The eleven supply chain views. Each returns {kpis, insights, charts, tables} built only from the governed model.

Chart specs are renderer-neutral: bar / hbar / line / area / doughnut / gauge / scatter / heatmap / treemap / sankey.
"""
import math
from contextvars import ContextVar

import numpy as np
import pandas as pd

from . import config as C

VIEWS = [
    ("01", "inventory_position", "Inventory Position", "What RM, WIP and FG do we physically have?"),
    ("02", "inventory_by_location", "Inventory by Location", "Where is the inventory physically located?"),
    ("03", "rm_consumption_cover", "Raw Material Consumption & Cover",
     "How fast are raw materials consumed and how many production days will current RM support?"),
    ("04", "wip_cover", "WIP Cover", "How much complete production can current WIP support?"),
    ("05", "fg_cover", "Finished Goods Cover", "How many days of expected demand can current FG support?"),
    ("06", "in_transit_overdue", "In-Transit & Overdue Inventory",
     "What is coming, how much is outstanding, when is it expected, and what is overdue?"),
    ("07", "projection_stockout", "Projected Inventory & Stockout Exposure",
     "What will inventory look like over time, and what could run out before replenishment arrives?"),
    ("08", "inventory_risk", "Inventory Risk", "Which items require attention based on cover and incoming supply?"),
    ("09", "location_risk", "Location Supply Risk & Rebalancing",
     "Which locations are at risk, and can stock elsewhere address the shortage?"),
    ("10", "bom_readiness", "BOM / Production Readiness",
     "What can be produced from available/incoming RM and WIP, and which components constrain production?"),
    ("11", "exceptions", "Supply Chain Exceptions",
     "What needs management attention across stockouts, low cover, overdue receipts, locations and production?"),
]
VIEW_IDS = {v[1] for v in VIEWS}

DEFINITIONS = {
    "Stock on hand": "Sum of Remaining Quantity on open Item Ledger Entries at the reporting date, by item and "
                     "location. Stock in in-transit locations is shown separately.",
    "Item class": "Raw Material / WIP / Finished Goods from the item's Inventory Posting Group (see config). "
                  "Spares, consumables and stationery are 'Other' and excluded from cover views.",
    "Daily usage": "RM and WIP: Consumption + Assembly Consumption over the window ÷ calendar days. "
                   "FG: Sales + Consumption over the window ÷ calendar days.",
    "RM production-day cover": "RM stock ÷ (consumption ÷ number of days with production output in the window).",
    "Cover days": "Stock on hand ÷ daily usage.",
    "ABC class": "Within each item class, items ranked by usage value in the window: A = top 80% of usage value, "
                 "B = next 15%, C = the rest (including items with no usage).",
    "WIP cover": "Complete parent units the WIP stock supports = WIP stock ÷ quantity per (incl. scrap) on the "
                 "certified production BOM of its highest-output parent.",
    "FG cover": "FG stock ÷ expected daily demand (recent sales + consumption).",
    "Made to order": "FG with no stock where output in the window covers at least 90% of sales: produced and "
                     "shipped straight away, so not treated as a stockout.",
    "In-transit quantity": "Outstanding quantity on open purchase order lines, with expected receipt date.",
    "Overdue receipt": "Open purchase line whose expected receipt date is before the reporting date. Lines more "
                       "than 90 days overdue are 'stale' and excluded from projections.",
    "Projected inventory": "Stock on hand − daily usage × days + expected receipts (overdue lines assumed to land "
                           "7 days after the reporting date), weekly for 12 weeks.",
    "Stockout exposure": "Item whose projected stock reaches zero before its next receipt arrives (or with no "
                         "receipt due).",
    "Inventory risk": "High: stockout/critical cover (<7d) with weak supply. Medium: stockout/critical with "
                      "supply on time, low cover (<14d) with weak supply, or stockout exposure. Low: low cover.",
    "Location supply risk": "Site (e.g. SHIP 4 = SHIP4RM+SHIP4PR+SHIP4FG) where cover < 14 days. Surplus = stock "
                            "above 60 days of that site's usage; transfer suggested up to 30 days' cover.",
    "Production readiness": "For each actively produced item: units that can be made = min over BOM components "
                            "of stock ÷ quantity per; the component giving the minimum is the constraint. "
                            "Components shared by several products are counted fully for each product.",
    "Exceptions": "One ranked list (High → Low severity, then value at stake) of stockouts, stockout exposure, "
                  "low cover, overdue receipts, location risks and production constraints.",
}

CLS = ["RM", "WIP", "FG"]
ITEM_COLS = [("item_no", "Item", "text"), ("description", "Description", "text"), ("item_class", "Class", "text")]
COVER_ORDER = ["Stockout", "Critical", "Low", "Healthy", "Excess", "Made to order", "No movement"]
COVER_KEY = {"Stockout": "crit", "Critical": "serious", "Low": "warning", "Healthy": "good", "Excess": "violet",
             "Made to order": "seq", "No movement": "muted"}


# ---------------------------------------------------------------- helpers
def _clean(v):
    if v is None:
        return None
    if isinstance(v, (float, np.floating)):
        return None if (math.isnan(v) or math.isinf(v)) else round(float(v), 2)
    if isinstance(v, (np.integer,)):
        return int(v)
    if isinstance(v, (np.bool_,)):
        return bool(v)
    if hasattr(v, "isoformat"):
        return v.isoformat()
    return v


ROW_LIMIT = ContextVar("ROW_LIMIT", default=500)   # reports raise this to export every row


def table(title, df, cols, limit=None, drill="item_no", bars=True):
    """cols: list of (field, label, kind) where kind in text|qty|money|days|date|status."""
    d = df.head(min(limit, ROW_LIMIT.get()) if limit else ROW_LIMIT.get())
    return {"title": title, "total_rows": int(len(df)), "drill": drill, "bars": bars,
            "columns": [{"field": f, "label": l, "kind": k} for f, l, k in cols],
            "rows": [{f: _clean(r.get(f)) for f, _, _ in cols} for r in d.to_dict("records")]}


def kpi(label, value, kind="qty", note=None, status=None, spark=None, pct=None, view=None):
    out = {"label": label, "value": _clean(value), "kind": kind, "note": note, "status": status}
    if spark is not None:
        out["spark"] = [_clean(x) for x in spark]
    if pct is not None and not (isinstance(pct, float) and math.isnan(pct)):
        out["pct"] = round(float(max(0, min(100, pct))), 1)
    if view:
        out["view"] = view
    return out


def chart(kind, title, labels, series, unit=None, stacked=False, subtitle=None, span=1, height=None, extra=None,
          colors=None):
    out = {"type": kind, "title": title, "labels": [_clean(x) for x in labels], "unit": unit, "stacked": stacked,
           "series": [{"name": n, "key": k, "data": [_clean(x) for x in data]} for n, k, data in series],
           "subtitle": subtitle, "span": span}
    if height:
        out["height"] = height
    if extra:
        out["extra"] = extra
    if colors:
        out["colors"] = colors
    return out


def insight(tone, text):
    return {"tone": tone, "text": text}


def money(v):
    v = float(v or 0)
    if abs(v) >= 1e9:
        return f"{v / 1e9:,.2f} B UGX"
    if abs(v) >= 1e6:
        return f"{v / 1e6:,.1f} M UGX"
    return f"{v:,.0f} UGX"


def pct(a, b):
    return 100.0 * a / b if b else 0.0


def apply_filters(df, f):
    if df is None or not len(df):
        return df
    if f.get("item_class") and "item_class" in df:
        df = df[df.item_class == f["item_class"]]
    if f.get("site") and "site" in df:
        df = df[df.site == f["site"]]
    if f.get("location") and "location" in df:
        df = df[df.location == f["location"]]
    if f.get("q"):
        q = f["q"].lower()
        cols = [c for c in ("item_no", "description", "vendor_name", "po_no", "posting_group") if c in df]
        mask = np.zeros(len(df), bool)
        for c in cols:
            mask |= df[c].astype(str).str.lower().str.contains(q, regex=False).to_numpy()
        df = df[mask]
    return df


def _weekly(m, cls=None, field="consumed_value", f=None):
    w = m["weekly"]
    if cls:
        w = w[w.item_class == cls]
    elif f and f.get("item_class"):
        w = w[w.item_class == f["item_class"]]
    s = w.groupby("week")[field].sum().sort_index()
    return [str(x) for x in s.index], s.tolist()


def _full_weeks(labels, values):
    """Drop a partial first and last week so trend lines aren't distorted."""
    if len(values) > 3:
        return labels[1:-1], values[1:-1]
    return labels, values


def _cover_bins(it, col="cover_days"):
    c = it[col]
    bins = [("Out of stock", it.cover_status.eq("Stockout"), "crit"),
            ("< 7 days", (c < 7) & ~it.cover_status.eq("Stockout"), "serious"),
            ("7–14", (c >= 7) & (c < 14), "warning"), ("14–30", (c >= 14) & (c < 30), "good"),
            ("30–90", (c >= 30) & (c < 90), "good"), ("90–180", (c >= 90) & (c < 180), "seq"),
            ("> 180", c >= 180, "violet")]
    return [b for b, _, _ in bins], [int(mk.sum()) for _, mk, _ in bins], [k for _, _, k in bins]


def _scatter(it, color_by, keys, title, subtitle, size_col="stock_value", ycol="cover_days", ylabel="Cover days"):
    d = it[(it.daily_usage_value > 0)].copy()
    d["y"] = d[ycol].fillna(0).clip(lower=0.1, upper=365)
    mx = d[size_col].max() or 1
    pts = []
    for grp, key in keys.items():
        g = d[d[color_by] == grp].nlargest(150, "daily_usage_value")
        pts.append({"name": grp, "key": key, "data": [
            {"x": round(float(r.daily_usage_value), 0), "y": round(float(r.y), 1),
             "r": round(3 + 14 * math.sqrt(max(float(getattr(r, size_col) or 0), 0) / mx), 1),
             "label": f"{r.item_no} {str(r.description)[:40]}"} for r in g.itertuples()]})
    return {"type": "scatter", "title": title, "subtitle": subtitle, "labels": [], "unit": None, "stacked": False,
            "span": 2, "series": pts,
            "extra": {"xLabel": "Usage value per day (UGX, log scale)", "yLabel": ylabel + " (log scale, capped at 365)",
                      "xLog": True, "yLog": True, "yRef": [7, 14]}}


def _heatmap(title, rows, cols, values, unit="items", key="seq", subtitle=None, span=1):
    return {"type": "heatmap", "title": title, "subtitle": subtitle, "labels": rows, "unit": unit, "stacked": False,
            "span": span, "series": [], "extra": {"cols": cols, "values": [[_clean(v) for v in r] for r in values],
                                                  "key": key}}


def _treemap(title, items, unit="UGX", subtitle=None, span=1):
    return {"type": "treemap", "title": title, "subtitle": subtitle, "labels": [], "unit": unit, "stacked": False,
            "span": span, "series": [], "extra": {"items": [{"label": str(a), "value": _clean(v), "group": g}
                                                            for a, v, g in items if v and v > 0]}}


def _sankey(title, flows, unit="UGX", subtitle=None, span=1):
    return {"type": "sankey", "title": title, "subtitle": subtitle, "labels": [], "unit": unit, "stacked": False,
            "span": span, "series": [], "extra": {"flows": [{"from": a, "to": b, "flow": _clean(v)} for a, b, v in flows
                                                            if v and v > 0]}}


def _gauge(title, value, subtitle=None, key="good"):
    return {"type": "gauge", "title": title, "subtitle": subtitle, "labels": [], "unit": "%", "stacked": False,
            "span": 1, "series": [], "extra": {"value": round(float(value), 1), "key": key}}


# ---------------------------------------------------------------- 01
def inventory_position(m, f):
    sl = apply_filters(m["stock_loc"], f)
    sl = sl[(sl.item_class != "OTHER") & ~sl.in_transit]
    by = sl.groupby("item_class").agg(value=("stock_value", "sum"), qty=("stock_qty", "sum"),
                                      items=("item_no", "nunique")).reindex(CLS).fillna(0)
    tot = by.value.sum()
    it = apply_filters(m["item"], f)
    it = it[it.stock_qty > 0].sort_values("stock_value", ascending=False)
    pg = sl.groupby(["posting_group", "item_class"]).stock_value.sum().reset_index().sort_values("stock_value",
                                                                                               ascending=False)
    top = it.head(15)
    health = it.groupby("cover_status").stock_value.sum().reindex(COVER_ORDER).fillna(0)
    hc = it.pivot_table(index="item_class", columns="cover_status", values="stock_value", aggfunc="sum",
                        fill_value=0).reindex(index=CLS, columns=COVER_ORDER, fill_value=0)
    allit = apply_filters(m["item"], f)
    abc = allit.groupby("abc").agg(items=("item_no", "size"), uv=("usage_value", "sum"), sv=("stock_value", "sum"))
    abc = abc.reindex(["A", "B", "C"]).fillna(0)
    excess = health.get("Excess", 0) + health.get("No movement", 0)
    top10 = it.head(10).stock_value.sum()
    ins = [insight("info", f"Raw material holds {pct(by.loc['RM', 'value'], tot):.0f}% of stock value "
                           f"({money(by.loc['RM', 'value'])}); finished goods {pct(by.loc['FG', 'value'], tot):.0f}%."),
           insight("info", f"The 10 most valuable items hold {pct(top10, tot):.0f}% of all stock value "
                           f"— led by {top.item_no.iloc[0]} {str(top.description.iloc[0])[:40]} ({money(top.stock_value.iloc[0])}).")
           if len(top) else None,
           insight("serious", f"{money(excess)} ({pct(excess, tot):.0f}%) sits in items with over 180 days of cover or "
                              f"no movement in {m['notes']['window_days']} days — candidates to slow purchasing."),
           insight("good", f"A-class items: {int(abc.loc['A', 'items'])} items drive 80% of usage value and hold "
                           f"{money(abc.loc['A', 'sv'])} of stock.")]
    return {
        "kpis": [kpi("Total RM / WIP / FG value", tot, "money", f"{int(by['items'].sum()):,} items in stock",
                     spark=_full_weeks(*_weekly(m, f=f, field="received_value"))[1])]
                + [kpi(C.CLASS_LABELS[c], by.loc[c, "value"], "money", f"{int(by.loc[c, 'items']):,} items · "
                       f"{pct(by.loc[c, 'value'], tot):.0f}% of value", pct=pct(by.loc[c, "value"], tot)) for c in CLS]
                + [kpi("Excess / not moving", excess, "money", "cover > 180 days or no movement", status="warning",
                       pct=pct(excess, tot))],
        "insights": [i for i in ins if i],
        "charts": [chart("doughnut", "Stock value by class", [C.CLASS_LABELS[c] for c in CLS],
                         [("Value", "class", by.value.tolist())], "UGX", subtitle="Share of on-hand value",
                         extra={"quantity": [_clean(x) for x in by.qty.tolist()]}),
                   _treemap("Stock value by posting group", [(r.posting_group, r.stock_value, r.item_class)
                                                              for r in pg.itertuples()],
                            subtitle="Box size = value; colour = class", span=2),
                   chart("bar", "Stock health by value", [C.CLASS_LABELS[c] for c in CLS],
                         [(s, COVER_KEY[s], hc[s].tolist()) for s in COVER_ORDER if hc[s].sum() > 0], "UGX",
                         stacked=True, subtitle="How much of each class is healthy, short or excess"),
                   chart("hbar", "Top 15 items by stock value", [f"{r.item_no} · {str(r.description)[:28]}" for r in top.itertuples()],
                         [("Value", "class:" + ",".join(top.item_class.tolist()), top.stock_value.tolist())], "UGX",
                         span=2, height=380),
                   chart("bar", "ABC analysis", ["A", "B", "C"],
                         [("% of items", "seq", [pct(x, abc["items"].sum()) for x in abc["items"]]),
                          ("% of usage value", "FG", [pct(x, abc.uv.sum()) for x in abc.uv]),
                          ("% of stock value", "WIP", [pct(x, abc.sv.sum()) for x in abc.sv])], "%",
                         subtitle="Few A items drive most usage — focus control there",
                         extra={"items": [_clean(x) for x in abc["items"].tolist()],
                                "usage": [_clean(x) for x in abc.uv.tolist()],
                                "stock": [_clean(x) for x in abc.sv.tolist()]})],
        "tables": [table("Stock on hand by item", it, ITEM_COLS + [
            ("abc", "ABC", "text"), ("posting_group", "Posting group", "text"), ("stock_qty", "Stock", "qty"),
            ("uom", "UoM", "text"), ("stock_value", "Value (UGX)", "money"), ("cover_days", "Cover days", "days"),
            ("cover_status", "Cover status", "status"), ("locations", "Locations", "qty")])],
    }


# ---------------------------------------------------------------- 02
def inventory_by_location(m, f):
    sl = apply_filters(m["stock_loc"], f)
    sl = sl[sl.item_class != "OTHER"]
    loc = sl.pivot_table(index=["site", "location", "location_name"], columns="item_class", values="stock_value",
                         aggfunc="sum", fill_value=0).reindex(columns=CLS, fill_value=0)
    loc["total"] = loc.sum(axis=1)
    loc["items"] = sl.groupby(["site", "location", "location_name"]).item_no.nunique()
    loc = loc.sort_values("total", ascending=False).reset_index()
    site = loc.groupby("site")[CLS + ["total"]].sum().sort_values("total", ascending=False)
    site_items = sl.groupby("site").item_no.nunique().reindex(site.index).fillna(0)
    tot = site.total.sum()
    top = loc.head(15)
    detail = sl.sort_values("stock_value", ascending=False)
    top3 = loc.head(3).total.sum()
    ins = [insight("info", f"{site.index[0]} holds {pct(site.total.iloc[0], tot):.0f}% of stock value "
                           f"({money(site.total.iloc[0])}) across {int(site_items.iloc[0])} items.") if len(site) else None,
           insight("info", f"The three largest locations ({', '.join(loc.head(3).location)}) hold "
                           f"{pct(top3, tot):.0f}% of value — a concentration risk for stock counts and security.")
           if len(loc) >= 3 else None,
           insight("serious", f"{money(sl[sl.in_transit].stock_value.sum())} is sitting in in-transit locations "
                              f"and is not yet usable.") if sl.in_transit.any() else None]
    return {
        "kpis": [kpi("Locations holding stock", int((loc.total > 0).sum()), "qty", f"{len(site)} sites"),
                 kpi("Largest site", site.index[0] if len(site) else "-", "text",
                     f"{money(site.total.iloc[0])} · {pct(site.total.iloc[0], tot):.0f}%" if len(site) else None,
                     pct=pct(site.total.iloc[0], tot) if len(site) else None),
                 kpi("Top 3 locations share", pct(top3, tot), "days", "of total value (%)", pct=pct(top3, tot)),
                 kpi("Stock in transit locations", sl[sl.in_transit].stock_value.sum(), "money", status="warning")],
        "insights": [i for i in ins if i],
        "charts": [chart("bar", "Stock value by site", site.index.tolist(),
                         [(C.CLASS_LABELS[c], c, site[c].tolist()) for c in CLS], "UGX", stacked=True, span=2),
                   _heatmap("Value heatmap: site × class", site.index.tolist(), [C.CLASS_LABELS[c] for c in CLS],
                            site[CLS].values.tolist(), "UGX", subtitle="Darker = more value"),
                   _treemap("Locations sized by value", [(f"{r.location}", r.total, r.site) for r in loc.itertuples()],
                            subtitle="Grouped by site", span=2),
                   chart("hbar", "Top 15 locations by stock value", top.location.tolist(),
                         [(C.CLASS_LABELS[c], c, top[c].tolist()) for c in CLS], "UGX", stacked=True, height=380),
                   chart("hbar", "Distinct items held per site", site_items.index.tolist(),
                         [("Items", "seq", site_items.tolist())], "items")],
        "tables": [table("Stock value by location", loc, [
            ("site", "Site", "text"), ("location", "Location", "text"), ("location_name", "Name", "text"),
            ("RM", "RM (UGX)", "money"), ("WIP", "WIP (UGX)", "money"), ("FG", "FG (UGX)", "money"),
            ("total", "Total (UGX)", "money"), ("items", "Items", "qty")], drill="location"),
            table("Item stock by location", detail, ITEM_COLS + [
                ("location", "Location", "text"), ("site", "Site", "text"), ("stock_qty", "Stock", "qty"),
                ("uom", "UoM", "text"), ("stock_value", "Value (UGX)", "money")])],
    }


# ---------------------------------------------------------------- 03
def rm_consumption_cover(m, f):
    n = m["notes"]
    allrm = apply_filters(m["item"], {**f, "item_class": "RM"})
    it = allrm[allrm.daily_usage > 0].sort_values("daily_usage_value", ascending=False)
    labels, counts, keys = _cover_bins(it, "cover_prod_days")
    mv = it.head(15)
    w = it.stock_value.sum() / (it.daily_usage_value.sum() or 1)
    wl, wv = _full_weeks(*_weekly(m, "RM"))
    wpg = m["weekly_pg"]
    wpg = wpg[wpg.item_class == "RM"]
    topg = wpg.groupby("posting_group").consumed_value.sum().sort_values(ascending=False)
    piv = wpg[wpg.posting_group.isin(topg.index[:5])].pivot_table(index="week", columns="posting_group",
                                                                  values="consumed_value", aggfunc="sum", fill_value=0)
    piv = piv.iloc[1:-1] if len(piv) > 3 else piv
    crit = it[it.cover_prod_days < 7]
    crit_nosup = crit[crit.supply_status != "On time"]
    cons_val = (it.usage_qty * it.unit_cost.fillna(0)).sum()
    ins = [insight("info", f"RM consumption runs at {money(cons_val / n['production_days'])} per production day; "
                           f"{', '.join(topg.index[:2])} account for {pct(topg.iloc[:2].sum(), topg.sum()):.0f}% of it."),
           insight("critical", f"{len(crit)} consumed raw materials have under 7 production days of cover; "
                               f"{len(crit_nosup)} of them have no on-time supply.") if len(crit) else None,
           insight("info", f"Top consumer: {mv.item_no.iloc[0]} {str(mv.description.iloc[0])[:40]} at "
                           f"{money(mv.daily_usage_value.iloc[0])}/day with {mv.cover_days.iloc[0]:,.0f} days of cover.")
           if len(mv) else None,
           insight("good", f"Value-weighted RM cover is {w:,.0f} days — "
                           f"{'comfortable overall; risk is concentrated in specific items' if w > 30 else 'tight overall'}.")]
    return {
        "kpis": [kpi("RM consumption / production day", cons_val / n["production_days"], "money",
                     f"{n['production_days']} production days in window", spark=wv),
                 kpi("Value-weighted RM cover", w, "days", "calendar days"),
                 kpi("Under 7 production days", len(crit), "qty", f"of {len(it)} consumed RM items", status="critical",
                     pct=pct(len(crit), len(it))),
                 kpi("Under 14 production days", int((it.cover_prod_days < 14).sum()), "qty", status="serious",
                     pct=pct((it.cover_prod_days < 14).sum(), len(it))),
                 kpi("RM stock value", allrm.stock_value.sum(), "money")],
        "insights": [i for i in ins if i],
        "charts": [chart("area", "Weekly RM consumption by posting group (UGX)", [str(x) for x in piv.index],
                         [(c, "cat", piv[c].tolist()) for c in piv.columns], "UGX", stacked=True, span=2,
                         subtitle="Top 5 groups, full weeks only"),
                   chart("bar", "Items by production days of cover", labels, [("Items", "seq", counts)], "items",
                         colors=keys, subtitle="Red = act now"),
                   _scatter(it, "cover_status", {"Stockout": "crit", "Critical": "serious", "Low": "warning",
                                                 "Healthy": "good", "Excess": "violet"},
                            "RM risk map: usage vs cover",
                            "Bottom-right = high usage, low cover (act first). Bubble = stock value."),
                   chart("doughnut", "Consumption value by posting group", topg.index[:8].tolist(),
                         [("Value", "cat", topg.iloc[:8].tolist())], "UGX"),
                   chart("hbar", "Top 15 raw materials by daily usage value",
                         [f"{r.item_no} · {str(r.description)[:26]}" for r in mv.itertuples()],
                         [("UGX / day", "RM", mv.daily_usage_value.tolist())], "UGX", span=2, height=400)],
        "tables": [table("Raw material consumption & cover", it, ITEM_COLS[:2] + [
            ("abc", "ABC", "text"), ("posting_group", "Group", "text"), ("stock_qty", "Stock", "qty"), ("uom", "UoM", "text"),
            ("usage_qty", "Consumed in window", "qty"), ("daily_usage_prod", "Per production day", "qty"),
            ("daily_usage_value", "Usage value / day", "money"),
            ("cover_prod_days", "Production days cover", "days"), ("incoming_qty", "Incoming (PO)", "qty"),
            ("supply_status", "Supply", "status"), ("cover_status", "Status", "status")])],
    }


# ---------------------------------------------------------------- 04
def wip_cover(m, f):
    w = apply_filters(m["wip"], f)
    w = w[(w.stock_qty > 0) | (w.daily_usage > 0)].sort_values("stock_value", ascending=False)
    fw = apply_filters(m["fg_from_wip"], f)
    fw = fw[fw.daily_output > 0].sort_values("days_of_output")
    labels, counts, keys = _cover_bins(w)
    top = fw.head(15)
    wl, made = _full_weeks(*_weekly(m, "WIP", field="output_value"))
    _, used = _full_weeks(*_weekly(m, "WIP", field="consumed_value"))
    cat = m["items"].set_index("item_no").category
    wc = w.assign(cat=w.item_no.map(cat).fillna("Other")).groupby("cat").stock_value.sum().sort_values(ascending=False)
    short = fw[fw.days_of_output < 7]
    ins = [insight("info", f"WIP worth {money(w.stock_value.sum())} across {len(w)} items; "
                           f"median cover {w.cover_days.median():,.0f} days."),
           insight("critical", f"{len(short)} products have under 7 days of output from current WIP — first: "
                               f"{short.item_no.iloc[0]} {str(short.parent_description.iloc[0])[:40]} "
                               f"({short.days_of_output.iloc[0]:,.1f} days, limited by {short.limiting_wip.iloc[0]}).")
           if len(short) else None,
           insight("info", f"Over the last full weeks WIP output averaged {money(np.mean(made) if made else 0)} and "
                           f"consumption {money(np.mean(used) if used else 0)} per week "
                           f"({'building' if (np.mean(made) if made else 0) > (np.mean(used) if used else 0) else 'drawing down'} WIP).")]
    return {
        "kpis": [kpi("WIP value", w.stock_value.sum(), "money", f"{len(w):,} WIP items", spark=made),
                 kpi("WIP linked to a BOM parent", int(w.main_parent.notna().sum()), "qty",
                     pct=pct(w.main_parent.notna().sum(), len(w))),
                 kpi("Products < 7 days from WIP", len(short), "qty", f"of {len(fw)} products", status="serious",
                     pct=pct(len(short), len(fw))),
                 kpi("Median WIP cover", w.cover_days.median(), "days")],
        "insights": [i for i in ins if i],
        "charts": [chart("line", "Weekly WIP produced vs consumed (UGX)", wl,
                         [("Produced", "good", made), ("Consumed", "serious", used)], "UGX", span=2),
                   chart("bar", "WIP items by days of cover", labels, [("Items", "seq", counts)], "items", colors=keys),
                   chart("hbar", "Products with least output supported by WIP (days)",
                         [f"{r.item_no} · {str(r.parent_description)[:24]}" for r in top.itertuples()],
                         [("Days of output", "WIP", top.days_of_output.tolist())], "days", span=2, height=400),
                   _treemap("WIP value by category", [(k, v, "WIP") for k, v in wc.items()])],
        "tables": [table("WIP cover", w, ITEM_COLS[:2] + [
            ("stock_qty", "WIP stock", "qty"), ("uom", "UoM", "text"), ("stock_value", "Value (UGX)", "money"),
            ("daily_usage", "Used / day", "qty"), ("cover_days", "Cover days", "days"),
            ("main_parent", "Main parent", "text"), ("main_parent_description", "Parent description", "text"),
            ("parent_units_supported", "Complete parent units", "qty"), ("parent_count", "Parents", "qty")]),
            table("Complete production supportable from WIP", fw, [
                ("item_no", "Product", "text"), ("parent_description", "Description", "text"),
                ("units_from_wip", "Units WIP supports", "qty"), ("limiting_wip", "Limiting WIP", "text"),
                ("daily_output", "Output / day", "qty"), ("days_of_output", "Days of output", "days")])],
    }


# ---------------------------------------------------------------- 05
def fg_cover(m, f):
    it = apply_filters(m["item"], {**f, "item_class": "FG"})
    it = it[(it.daily_usage > 0) | (it.stock_qty > 0)].sort_values("daily_usage_value", ascending=False)
    labels, counts, keys = _cover_bins(it[it.cover_status != "Made to order"])
    st = it.cover_status.value_counts().reindex(COVER_ORDER).fillna(0)
    wl, sales = _full_weeks(*_weekly(m, "FG", field="sold_value"))
    _, made = _full_weeks(*_weekly(m, "FG", field="output_value"))
    top = it.head(15)
    excess_v = it[it.cover_status == "Excess"].stock_value.sum()
    so = it[it.cover_status == "Stockout"].sort_values("daily_usage_value", ascending=False)
    wcover = it.stock_value.sum() / (it.daily_usage_value.sum() or 1)
    ins = [insight("info", f"FG demand runs at {money(it.daily_usage_value.sum())} per day (at cost); stock covers "
                           f"{wcover:,.0f} days on a value-weighted basis."),
           insight("critical", f"{len(so)} FG items with demand are out of stock — biggest: {so.item_no.iloc[0]} "
                               f"{str(so.description.iloc[0])[:40]} ({money(so.daily_usage_value.iloc[0])}/day).")
           if len(so) else None,
           insight("info", f"{int(st['Made to order'])} items behave as made-to-order (produced and shipped "
                           f"straight away) and are not counted as stockouts."),
           insight("warning", f"{money(excess_v)} of FG has more than 180 days of cover.") if excess_v else None]
    return {
        "kpis": [kpi("FG value", it.stock_value.sum(), "money"),
                 kpi("FG demand / day (cost)", it.daily_usage_value.sum(), "money", spark=sales and [s / 7 for s in sales]),
                 kpi("Value-weighted FG cover", wcover, "days"),
                 kpi("FG stockouts", int(st["Stockout"]), "qty", "excl. made to order", status="critical",
                     pct=pct(st["Stockout"], len(it))),
                 kpi("Made to order", int(st["Made to order"]), "qty"),
                 kpi("Excess FG (> 180 days)", excess_v, "money", status="warning")],
        "insights": [i for i in ins if i],
        "charts": [chart("line", "Weekly FG sales vs production (UGX at cost)", wl,
                         [("Sales", "FG", sales), ("Production output", "seq", made)], "UGX", span=2),
                   chart("doughnut", "FG items by cover status", [s for s in COVER_ORDER if st[s] > 0],
                         [("Items", "status", [st[s] for s in COVER_ORDER if st[s] > 0])], "items",
                         colors=[COVER_KEY[s] for s in COVER_ORDER if st[s] > 0]),
                   chart("bar", "FG items by days of demand cover", labels, [("Items", "seq", counts)], "items",
                         colors=keys),
                   _scatter(it, "cover_status", {"Stockout": "crit", "Critical": "serious", "Low": "warning",
                                                 "Healthy": "good", "Excess": "violet"},
                            "FG risk map: demand vs cover", "Bottom-right = high demand, low cover. Bubble = stock value."),
                   chart("hbar", "Top 15 FG by daily demand value",
                         [f"{r.item_no} · {str(r.description)[:26]}" for r in top.itertuples()],
                         [("UGX / day", "FG", top.daily_usage_value.tolist())], "UGX", span=2, height=400)],
        "tables": [table("Finished goods cover", it, ITEM_COLS[:2] + [
            ("abc", "ABC", "text"), ("stock_qty", "Stock", "qty"), ("uom", "UoM", "text"), ("sold", "Sold in window", "qty"),
            ("daily_usage", "Demand / day", "qty"), ("daily_usage_value", "Demand value / day", "money"),
            ("cover_days", "Cover days", "days"), ("output", "Produced in window", "qty"),
            ("cover_status", "Status", "status")])],
    }


# ---------------------------------------------------------------- 06
def in_transit_overdue(m, f):
    pl = apply_filters(m["purchase_lines"], f)
    pl2 = pl.copy()
    pl2["bucket"] = pd.cut(pl2.days_overdue, [-10_000, 0, 7, 30, 90, 10_000],
                           labels=["Not yet due", "1–7 days late", "8–30 days", "31–90 days", "> 90 (stale)"])
    age = pl2.groupby("bucket", observed=False).outstanding_value.sum()
    age_n = pl2.groupby("bucket", observed=False).size()
    od = pl[pl.overdue & ~pl.stale]
    ven = od.groupby("vendor_name").outstanding_value.sum().sort_values(ascending=False).head(12)
    tr = m["stock_loc"]
    tr = apply_filters(tr[tr.in_transit], f)
    pl2["month"] = pd.to_datetime(pl2.expected_date).dt.to_period("M").astype(str)
    tl = pl2.pivot_table(index="month", columns="receipt_status", values="outstanding_value", aggfunc="sum",
                         fill_value=0).sort_index()
    byc = pl.groupby(pl.item_class.fillna("OTHER")).outstanding_value.sum()
    vtop = pl.groupby("vendor_name").outstanding_value.sum().sort_values(ascending=False).head(8).index
    flows = pl[pl.vendor_name.isin(vtop)].groupby(["vendor_name", "item_class"]).outstanding_value.sum()
    ins = [insight("critical", f"All {int(pl.overdue.sum())} of {len(pl)} open lines are past their expected date; "
                               f"{int(pl.stale.sum())} are more than 90 days late and look stale.")
           if pl.overdue.all() and len(pl) else
           insight("serious", f"{int(pl.overdue.sum())} of {len(pl)} open lines are overdue."),
           insight("serious", f"Recently overdue (≤90 days): {len(od)} lines worth {money(od.outstanding_value.sum())}; "
                              f"{ven.index[0].strip()} alone accounts for {money(ven.iloc[0])}.") if len(ven) else None,
           insight("info", f"Average lateness of overdue lines: {pl[pl.overdue].days_overdue.mean():,.0f} days; "
                           f"oldest {int(pl.days_overdue.max())} days.") if pl.overdue.any() else None,
           insight("info", "Close or re-date stale lines in Business Central so projections and buyer workload are accurate.")]
    return {
        "kpis": [kpi("Open PO lines", len(pl), "qty", f"{pl.po_no.nunique():,} orders · {pl.vendor_no.nunique()} vendors"),
                 kpi("Outstanding value", pl.outstanding_value.sum(), "money"),
                 kpi("Overdue ≤ 90 days", od.outstanding_value.sum(), "money", f"{len(od)} lines", status="serious",
                     pct=pct(od.outstanding_value.sum(), pl.outstanding_value.sum())),
                 kpi("Stale > 90 days", pl[pl.stale].outstanding_value.sum(), "money", f"{int(pl.stale.sum())} lines",
                     status="warning", pct=pct(pl[pl.stale].outstanding_value.sum(), pl.outstanding_value.sum())),
                 kpi("Average days late", pl[pl.overdue].days_overdue.mean() if pl.overdue.any() else 0, "days"),
                 kpi("Stock in transit locations", tr.stock_value.sum(), "money")],
        "insights": [i for i in ins if i],
        "charts": [chart("bar", "Outstanding value by lateness", [str(x) for x in age.index],
                         [("Value", "seq", age.tolist())], "UGX",
                         colors=["good", "warning", "serious", "crit", "muted"],
                         subtitle=" · ".join(f"{k}: {v} lines" for k, v in age_n.items())),
                   chart("bar", "Expected receipts by month (UGX)", tl.index.tolist(),
                         [(c, {"Expected": "good", "Overdue": "serious"}.get(c, "muted"), tl[c].tolist())
                          for c in tl.columns], "UGX", stacked=True, span=2,
                         subtitle="When open lines were due — grey = stale"),
                   chart("hbar", "Overdue value by vendor (≤ 90 days)", [v.strip()[:30] for v in ven.index],
                         [("Overdue value", "serious", ven.tolist())], "UGX", span=2, height=380),
                   chart("doughnut", "Open value by item class", [C.CLASS_LABELS.get(c, c) for c in byc.index],
                         [("Value", "class", byc.tolist())], "UGX"),
                   _sankey("Open order value: vendor → class", [(a.strip()[:24], C.CLASS_LABELS.get(b, b), v)
                                                                 for (a, b), v in flows.items()],
                           subtitle="Top 8 vendors", span=3)],
        "tables": [table("Open purchase order lines", pl, [
            ("po_no", "PO", "text"), ("vendor_name", "Vendor", "text"), ("item_no", "Item", "text"),
            ("description", "Description", "text"), ("item_class", "Class", "text"), ("location", "Location", "text"),
            ("ordered_qty", "Ordered", "qty"), ("received_qty", "Received", "qty"),
            ("outstanding_qty", "Outstanding", "qty"), ("expected_date", "Expected", "date"),
            ("days_overdue", "Days late", "qty"), ("receipt_status", "Status", "status"),
            ("outstanding_value", "Outstanding (UGX)", "money")])],
    }


# ---------------------------------------------------------------- 07
def projection_stockout(m, f):
    it = apply_filters(m["item"], f)
    it = it[it.daily_usage > 0]
    pr = m["projection"]
    pr = pr[pr.item_no.isin(it.item_no)].merge(it[["item_no", "unit_cost", "item_class"]], on="item_no")
    pr["value"] = pr.projected_qty.clip(lower=0) * pr.unit_cost.fillna(0)
    curve = pr.pivot_table(index="week_date", columns="item_class", values="value", aggfunc="sum").reindex(
        columns=CLS, fill_value=0).fillna(0)
    zero = pr[pr.projected_qty <= 0].pivot_table(index="week_date", columns="item_class", values="item_no",
                                                 aggfunc="nunique").reindex(index=curve.index, columns=CLS).fillna(0)
    ex = it[it.stockout_exposed].sort_values("stockout_day")
    out_now = ex.cover_status.eq("Stockout")
    fut = ex[~out_now].copy()
    fut["wk"] = (fut.stockout_day // 7).astype(int).clip(upper=12)
    tl = fut.pivot_table(index="wk", columns="item_class", values="item_no", aggfunc="count").reindex(
        index=range(13), columns=CLS).fillna(0)
    topx = fut.sort_values("daily_usage_value", ascending=False).head(15)
    ins = [insight("critical", f"{int(out_now.sum())} items in use are already out of stock and "
                               f"{int(((ex.stockout_day <= 7) & ~out_now).sum())} more run out within 7 days before any receipt."),
           insight("serious", f"Highest-impact upcoming stockout: {topx.item_no.iloc[0]} {str(topx.description.iloc[0])[:40]} "
                              f"on {topx.stockout_date.iloc[0]} ({money(topx.daily_usage_value.iloc[0])}/day).")
           if len(topx) else None,
           insight("info", f"Without new orders, projected stock value falls from {money(curve.iloc[0].sum())} today to "
                           f"{money(curve.iloc[-1].sum())} in 12 weeks.") if len(curve) else None]
    return {
        "kpis": [kpi("Items exposed to stockout", len(ex), "qty", "within 12 weeks, before replenishment",
                     status="critical"),
                 kpi("Already out of stock", int(out_now.sum()), "qty", status="critical"),
                 kpi("Run out within 7 days", int(((ex.stockout_day <= 7) & ~out_now).sum()), "qty", status="critical"),
                 kpi("Run out in 8–30 days", int(((ex.stockout_day > 7) & (ex.stockout_day <= 30)).sum()), "qty",
                     status="serious"),
                 kpi("Projected value in 12 weeks", curve.iloc[-1].sum() if len(curve) else 0, "money",
                     spark=curve.sum(axis=1).tolist(),
                     pct=pct(curve.iloc[-1].sum(), curve.iloc[0].sum()) if len(curve) else None)],
        "insights": [i for i in ins if i],
        "charts": [chart("area", "Projected stock value by week (UGX)", [d.isoformat() for d in curve.index],
                         [(C.CLASS_LABELS[c], c, curve[c].tolist()) for c in CLS], "UGX", stacked=True, span=2,
                         subtitle="Usage continues at the recent rate; open POs land on their expected dates"),
                   chart("bar", "Items at zero by week", [d.isoformat() for d in zero.index],
                         [(C.CLASS_LABELS[c], c, zero[c].tolist()) for c in CLS], "items", stacked=True),
                   chart("bar", "When items run out (weeks from today)", [f"Wk {i}" if i < 12 else "Wk 12+" for i in tl.index],
                         [(C.CLASS_LABELS[c], c, tl[c].tolist()) for c in CLS], "items", stacked=True, span=2,
                         subtitle="Items not yet out that hit zero before replenishment"),
                   chart("hbar", "Biggest upcoming stockouts: days left",
                         [f"{r.item_no} · {str(r.description)[:24]}" for r in topx.itertuples()],
                         [("Days to stockout", "class:" + ",".join(topx.item_class.tolist()), topx.stockout_day.tolist())],
                         "days", height=400)],
        "tables": [table("Stockout exposure", ex, ITEM_COLS + [
            ("abc", "ABC", "text"), ("stock_qty", "Stock", "qty"), ("daily_usage", "Usage / day", "qty"),
            ("daily_usage_value", "Usage value / day", "money"), ("cover_days", "Cover days", "days"),
            ("stockout_date", "Runs out", "date"), ("incoming_qty", "Incoming", "qty"),
            ("next_arrival_day", "Next receipt in (days)", "qty"), ("supply_status", "Supply", "status")])],
    }


# ---------------------------------------------------------------- 08
def inventory_risk(m, f):
    allit = apply_filters(m["item"], f)
    it = allit[allit.risk_level != "None"].sort_values("risk_score", ascending=False)
    cov_rows = [s for s in ["Stockout", "Critical", "Low", "Healthy", "Excess"] if s in set(it.cover_status)]
    sup_cols = [s for s in ["No open supply", "Only stale supply", "Supply arrives too late", "Overdue supply", "On time"]
                if s in set(it.supply_status)]
    mat = it.pivot_table(index="cover_status", columns="supply_status", values="item_no", aggfunc="count",
                         fill_value=0).reindex(index=cov_rows, columns=sup_cols, fill_value=0)
    lv = it.groupby(["risk_level", "item_class"]).size().unstack(fill_value=0).reindex(
        index=["High", "Medium", "Low"], columns=CLS, fill_value=0)
    hi = it[it.risk_level == "High"]
    top = hi.sort_values("daily_usage_value", ascending=False).head(15)
    ins = [insight("critical", f"{len(hi)} high-risk items (out of stock or < 7 days cover with weak supply) "
                               f"put {money(hi.daily_usage_value.sum())} of daily usage at risk."),
           insight("serious", f"Most common pattern: '{mat.stack().idxmax()[0]}' cover with '{mat.stack().idxmax()[1]}' "
                              f"({int(mat.stack().max())} items).") if mat.size else None,
           insight("info", f"By class: {', '.join(f'{C.CLASS_LABELS[c]} {int(lv[c].sum())}' for c in CLS)} items need attention.")]
    return {
        "kpis": [kpi("High risk items", len(hi), "qty", status="critical", pct=pct(len(hi), len(allit[allit.daily_usage > 0]))),
                 kpi("Medium risk items", int((it.risk_level == "Medium").sum()), "qty", status="serious"),
                 kpi("Low risk items", int((it.risk_level == "Low").sum()), "qty", status="warning"),
                 kpi("Usage value at risk / day", hi.daily_usage_value.sum(), "money", "high-risk items")],
        "insights": [i for i in ins if i],
        "charts": [_heatmap("Risk matrix: cover × supply", cov_rows, sup_cols, mat.values.tolist(), "items", "crit",
                            subtitle="Number of items in each combination", span=2),
                   chart("doughnut", "Items by risk level", ["High", "Medium", "Low"],
                         [("Items", "status", [int(lv.loc[x].sum()) for x in ["High", "Medium", "Low"]])], "items",
                         colors=["crit", "serious", "warning"]),
                   _scatter(it, "risk_level", {"High": "crit", "Medium": "serious", "Low": "warning"},
                            "Risk map: usage vs cover", "Each bubble an item at risk; size = stock value"),
                   chart("bar", "Items by risk level and class", lv.index.tolist(),
                         [(C.CLASS_LABELS[c], c, lv[c].tolist()) for c in CLS], "items", stacked=True),
                   chart("hbar", "Top 15 high-risk items by usage value",
                         [f"{r.item_no} · {str(r.description)[:24]}" for r in top.itertuples()],
                         [("UGX / day", "class:" + ",".join(top.item_class.tolist()), top.daily_usage_value.tolist())],
                         "UGX", span=3, height=400)],
        "tables": [table("Items requiring attention", it, ITEM_COLS + [
            ("abc", "ABC", "text"), ("risk_level", "Risk", "status"), ("cover_status", "Cover", "status"), ("cover_days", "Cover days", "days"),
            ("stock_qty", "Stock", "qty"), ("daily_usage", "Usage / day", "qty"), ("incoming_qty", "Incoming", "qty"),
            ("overdue_qty", "Overdue", "qty"), ("supply_status", "Supply", "status"),
            ("daily_usage_value", "Usage value / day", "money")])],
    }


# ---------------------------------------------------------------- 09
def location_risk(m, f):
    lr = apply_filters(m["location_risk"], f)
    covered_lbl = "No surplus elsewhere – purchase/produce"
    by = lr.groupby("site").agg(items=("item_no", "size"),
                                full=("rebalance", lambda s: (s == "Fully covered by transfer").sum()),
                                part=("rebalance", lambda s: (s == "Partly covered by transfer").sum()))
    by["none"] = by["items"] - by.full - by.part
    by = by.sort_values("items", ascending=False).head(15)
    tr = lr[lr.suggested_transfer > 0].copy()
    tr["value"] = tr.suggested_transfer * tr.unit_cost.fillna(0)
    flows = tr.groupby(["donor_site", "site"]).value.sum().sort_values(ascending=False)
    lanes = flows.head(14)
    rest = flows.iloc[14:]
    if len(rest):        # keep the diagram readable: smaller lanes grouped into "Other"
        extra = rest.groupby(level=1).sum()
        extra.index = pd.MultiIndex.from_tuples([("Other sites", b) for b in extra.index])
        lanes = pd.concat([lanes, extra]).groupby(level=[0, 1]).sum().sort_values(ascending=False)
    hm = lr.pivot_table(index="site", columns="item_class", values="item_no", aggfunc="count", fill_value=0).reindex(
        columns=CLS, fill_value=0)
    hm = hm.loc[hm.sum(axis=1).sort_values(ascending=False).index]
    topt = tr.sort_values("value", ascending=False).head(15)
    ins = [insight("serious", f"{len(lr)} item-site shortages across {lr.site.nunique()} sites; "
                              f"{by.index[0]} has the most ({int(by['items'].iloc[0])}).") if len(by) else None,
           insight("good", f"{int((lr.rebalance == 'Fully covered by transfer').sum())} shortages can be fully solved and "
                           f"{int((lr.rebalance == 'Partly covered by transfer').sum())} partly solved by moving "
                           f"{money(tr.value.sum())} of surplus between sites — no purchase needed."),
           insight("info", f"Largest transfer lane: {flows.idxmax()[0]} → {flows.idxmax()[1]} ({money(flows.max())}).")
           if len(flows) else None,
           insight("critical", f"{int((lr.rebalance == covered_lbl).sum())} shortages have no surplus anywhere and need "
                               f"purchase or production.")]
    return {
        "kpis": [kpi("Item-site shortages", len(lr), "qty", "cover < 14 days at site", status="serious"),
                 kpi("Sites affected", lr.site.nunique(), "qty"),
                 kpi("Fully solvable by transfer", int((lr.rebalance == "Fully covered by transfer").sum()), "qty",
                     status="good", pct=pct((lr.rebalance == "Fully covered by transfer").sum(), len(lr))),
                 kpi("Partly solvable", int((lr.rebalance == "Partly covered by transfer").sum()), "qty"),
                 kpi("Transferable value", tr.value.sum(), "money")],
        "insights": [i for i in ins if i],
        "charts": [_sankey("Suggested transfer flows (value)", [(f"{a} ", b, v) for (a, b), v in lanes.items()],
                           subtitle="From surplus site (left) to short site (right) · 14 largest lanes", span=2),
                   chart("bar", "Shortages by site and outcome", by.index.tolist(),
                         [("Fully by transfer", "good", by.full.tolist()), ("Partly by transfer", "warning", by.part.tolist()),
                          ("Needs purchase / production", "crit", by["none"].tolist())], "items", stacked=True),
                   _heatmap("Shortages: site × class", hm.index.tolist(), [C.CLASS_LABELS[c] for c in CLS],
                            hm.values.tolist(), "items", "serious"),
                   chart("hbar", "Top 15 suggested transfers by value",
                         [f"{r.item_no} {r.donor_site}→{r.site}" for r in topt.itertuples()],
                         [("Value", "class:" + ",".join(topt.item_class.tolist()), topt.value.tolist())], "UGX",
                         span=2, height=400)],
        "tables": [table("Location supply risk & rebalancing", lr, ITEM_COLS + [
            ("site", "Site at risk", "text"), ("stock_qty", "Site stock", "qty"),
            ("daily_usage", "Site usage / day", "qty"), ("cover_days", "Site cover days", "days"),
            ("need_qty", "Need for 30d", "qty"), ("donor_site", "Best donor site", "text"),
            ("donor_surplus", "Donor surplus", "qty"), ("suggested_transfer", "Suggested transfer", "qty"),
            ("rebalance", "Outcome", "status")])],
    }


# ---------------------------------------------------------------- 10
def bom_readiness(m, f):
    r = apply_filters(m["readiness"].rename(columns={"parent_description": "description", "parent_class": "item_class"}),
                      f)
    cons = r[r.readiness != "Ready"].groupby(["limiting_component", "limiting_description"]).size().sort_values(
        ascending=False).head(12).reset_index(name="products")
    st = r.readiness.value_counts().reindex(["Constrained", "At risk", "Ready"]).fillna(0)
    d = r.production_days_supported
    bins = [("0 (can't make)", d <= 0, "crit"), ("< 2 days", (d > 0) & (d < 2), "crit"), ("2–7", (d >= 2) & (d < 7), "serious"),
            ("7–14", (d >= 7) & (d < 14), "warning"), ("14–30", (d >= 14) & (d < 30), "good"), ("30+", d >= 30, "seq")]
    top = r[r.readiness != "Ready"].sort_values("daily_output", ascending=False).head(15)
    fix = top[top.production_days_with_incoming > top.production_days_supported]
    ins = [insight("critical", f"{int(st['Constrained'])} of {len(r)} actively produced items have under 7 days of "
                               f"components; {int((d <= 0).sum())} cannot be made at all from current stock."),
           insight("serious", f"Most constraining component: {cons.limiting_component.iloc[0]} "
                              f"{str(cons.limiting_description.iloc[0])[:40]} limits {int(cons.products.iloc[0])} products.")
           if len(cons) else None,
           insight("good", f"Open purchase orders would lift {len(fix)} of the 15 highest-volume constrained products.")
           if len(top) else None]
    return {
        "kpis": [kpi("Active products with BOM", len(r), "qty"),
                 kpi("Constrained (< 7 days)", int(st["Constrained"]), "qty", status="critical", pct=pct(st["Constrained"], len(r))),
                 kpi("At risk (7–14 days)", int(st["At risk"]), "qty", status="serious", pct=pct(st["At risk"], len(r))),
                 kpi("Ready (≥ 14 days)", int(st["Ready"]), "qty", status="good", pct=pct(st["Ready"], len(r)))],
        "insights": [i for i in ins if i],
        "charts": [_gauge("Production readiness", pct(st["Ready"], len(r)), "% of active products with ≥ 14 days of components",
                          "good"),
                   chart("bar", "Products by days of production supported", [b for b, _, _ in bins],
                         [("Products", "seq", [int(mk.sum()) for _, mk, _ in bins])], "products",
                         colors=[k for _, _, k in bins], span=2),
                   chart("hbar", "Components constraining the most products",
                         [f"{a} · {str(b)[:24]}" for a, b in zip(cons.limiting_component, cons.limiting_description)],
                         [("Products constrained", "crit", cons.products.tolist())], "products", span=2, height=380),
                   chart("doughnut", "Readiness", st.index.tolist(), [("Products", "status", st.tolist())], "products",
                         colors=["crit", "serious", "good"]),
                   chart("hbar", "Highest-volume constrained products: days now vs with incoming POs",
                         [f"{r_.item_no} · {str(r_.description)[:22]}" for r_ in top.itertuples()],
                         [("Days supported now", "serious", top.production_days_supported.tolist()),
                          ("Days incl. incoming", "good", top.production_days_with_incoming.clip(upper=60).tolist())],
                         "days", span=3, height=430)],
        "tables": [table("Production readiness", r, [
            ("item_no", "Product", "text"), ("description", "Description", "text"), ("item_class", "Class", "text"),
            ("daily_output", "Output / day", "qty"), ("can_make_now", "Can make now", "qty"),
            ("production_days_supported", "Days supported", "days"),
            ("can_make_with_incoming", "Can make incl. incoming", "qty"),
            ("production_days_with_incoming", "Days incl. incoming", "days"),
            ("limiting_component", "Limiting component", "text"), ("limiting_description", "Component", "text"),
            ("components", "Components", "qty"), ("readiness", "Readiness", "status")])],
    }


# ---------------------------------------------------------------- 11
def exceptions(m, f):
    e = apply_filters(m["exceptions"], f)
    sev = ["High", "Medium", "Low"]
    t = e.pivot_table(index="exception_type", columns="severity", values="item_no", aggfunc="count",
                      fill_value=0).reindex(columns=sev, fill_value=0)
    t = t.loc[t.sum(axis=1).sort_values(ascending=False).index]
    bc = e.pivot_table(index=e.item_class.fillna("OTHER"), columns="severity", values="item_no", aggfunc="count",
                       fill_value=0).reindex(columns=sev, fill_value=0)
    val = e.groupby("exception_type").value_at_stake.sum().sort_values(ascending=False)
    ins = [insight("critical", f"{int((e.severity == 'High').sum())} high-severity exceptions need action today; "
                               f"the largest group is '{t.High.idxmax()}' ({int(t.High.max())}).") if len(t) else None,
           insight("serious", f"Top item: {e.item_no.iloc[0]} {str(e.description.iloc[0])[:40]} — {e.detail.iloc[0]}")
           if len(e) else None,
           insight("info", f"Exceptions by type: " + ", ".join(f"{k} {int(v)}" for k, v in t.sum(axis=1).items()) + ".")]
    return {
        "kpis": [kpi("Exceptions", len(e), "qty"),
                 kpi("High", int((e.severity == "High").sum()), "qty", status="critical", pct=pct((e.severity == "High").sum(), len(e))),
                 kpi("Medium", int((e.severity == "Medium").sum()), "qty", status="serious", pct=pct((e.severity == "Medium").sum(), len(e))),
                 kpi("Low", int((e.severity == "Low").sum()), "qty", status="warning", pct=pct((e.severity == "Low").sum(), len(e)))],
        "insights": [i for i in ins if i],
        "charts": [chart("hbar", "Exceptions by type and severity", t.index.tolist(),
                         [("High", "crit", t.High.tolist()), ("Medium", "serious", t.Medium.tolist()),
                          ("Low", "warning", t.Low.tolist())], "exceptions", stacked=True, span=2),
                   chart("doughnut", "By severity", sev, [("Exceptions", "status", [int(t[s].sum()) for s in sev])],
                         "exceptions", colors=["crit", "serious", "warning"]),
                   chart("bar", "By item class and severity", [C.CLASS_LABELS.get(c, c) for c in bc.index],
                         [("High", "crit", bc.High.tolist()), ("Medium", "serious", bc.Medium.tolist()),
                          ("Low", "warning", bc.Low.tolist())], "exceptions", stacked=True),
                   chart("hbar", "Value at stake by exception type", val.index.tolist(),
                         [("UGX", "seq", val.tolist())], "UGX", span=2,
                         subtitle="Daily usage value for stock issues; outstanding value for overdue receipts")],
        "tables": [table("Ranked exceptions", e, [
            ("severity", "Severity", "status"), ("exception_type", "Type", "text"), ("item_no", "Item", "text"),
            ("description", "Description", "text"), ("item_class", "Class", "text"), ("location", "Location", "text"),
            ("detail", "Detail", "text"), ("value_at_stake", "Value at stake (UGX)", "money")])],
    }


# ---------------------------------------------------------------- overview
def overview(m, f):
    """One-page summary pulling the headline of every view together."""
    it = apply_filters(m["item"], f)
    sl = apply_filters(m["stock_loc"], f)
    sl = sl[(sl.item_class != "OTHER") & ~sl.in_transit]
    pl = apply_filters(m["purchase_lines"], f)
    ex = apply_filters(m["exceptions"], f)
    rd = apply_filters(m["readiness"].rename(columns={"parent_class": "item_class"}), f)
    lr = apply_filters(m["location_risk"], f)
    soon = it.stockout_exposed & (it.cover_status != "Stockout") & (it.stockout_day <= 7)
    by = sl.groupby("item_class").stock_value.sum().reindex(CLS).fillna(0)
    t = ex.groupby(["exception_type", "severity"]).size().unstack(fill_value=0).reindex(
        columns=["High", "Medium", "Low"], fill_value=0)
    t = t.loc[t.sum(axis=1).sort_values(ascending=False).index]
    used = it[it.daily_usage > 0]
    inv_health = pct(used.cover_status.isin(["Healthy", "Excess", "Made to order"]).sum(), len(used))
    od_v = pl[pl.overdue & ~pl.stale].outstanding_value.sum()
    supply_health = 100 - pct(pl[pl.overdue].outstanding_value.sum(), pl.outstanding_value.sum())
    prod_health = pct((rd.readiness == "Ready").sum(), len(rd))
    wl, cons = _full_weeks(*_weekly(m, field="consumed_value", f=f))
    _, sales = _full_weeks(*_weekly(m, field="sold_value", f=f))
    _, rec = _full_weeks(*_weekly(m, field="received_value", f=f))
    _, outp = _full_weeks(*_weekly(m, field="output_value", f=f))
    ins = [insight("critical", f"{int((it.cover_status == 'Stockout').sum())} items in use are out of stock and "
                               f"{int(soon.sum())} more run out within 7 days."),
           insight("serious", f"{int((pl.overdue & ~pl.stale).sum())} purchase lines worth {money(od_v)} are overdue "
                              f"(≤ 90 days); {int(pl.stale.sum())} more are stale."),
           insight("serious", f"{int((rd.readiness == 'Constrained').sum())} produced items have under 7 days of components."),
           insight("good", f"{int((lr.suggested_transfer > 0).sum())} site shortages can be eased by inter-site transfers.")]
    return {
        "kpis": [kpi("Stock value (RM/WIP/FG)", by.sum(), "money", "01 Inventory Position", view="inventory_position",
                     spark=cons),
                 kpi("Out of stock (in use)", int((it.cover_status == "Stockout").sum()), "qty",
                     "07 Projected Inventory", status="critical", view="projection_stockout"),
                 kpi("Run out within 7 days", int(soon.sum()), "qty", "07 Projected Inventory", status="critical",
                     view="projection_stockout"),
                 kpi("High-risk items", int((it.risk_level == "High").sum()), "qty", "08 Inventory Risk",
                     status="serious", view="inventory_risk"),
                 kpi("Overdue PO value (≤ 90d)", od_v, "money", f"{int((pl.overdue & ~pl.stale).sum())} lines · 06 In-Transit",
                     status="serious", view="in_transit_overdue"),
                 kpi("Constrained products", int((rd.readiness == "Constrained").sum()), "qty",
                     "10 Production Readiness", status="serious", view="bom_readiness"),
                 kpi("Site shortages", len(lr), "qty", f"{int((lr.suggested_transfer > 0).sum())} solvable by transfer",
                     status="warning", view="location_risk"),
                 kpi("High exceptions", int((ex.severity == "High").sum()), "qty", "11 Exceptions", status="critical",
                     view="exceptions")],
        "insights": ins,
        "charts": [_gauge("Inventory health", inv_health, "% of items in use with adequate cover",
                          "good" if inv_health >= 70 else "serious"),
                   _gauge("Supply health", supply_health, "% of open order value not overdue",
                          "good" if supply_health >= 70 else "crit"),
                   _gauge("Production readiness", prod_health, "% of produced items with ≥ 14 days of components",
                          "good" if prod_health >= 70 else "serious"),
                   chart("line", "Weekly flow at cost (UGX)", wl,
                         [("Consumption", "seq", cons), ("Production output", "good", outp), ("Sales", "FG", sales),
                          ("Purchase receipts", "serious", rec)], "UGX", span=2),
                   chart("doughnut", "Stock value by class", [C.CLASS_LABELS[c] for c in CLS],
                         [("Value", "class", by.tolist())], "UGX"),
                   chart("hbar", "Exceptions by type and severity", t.index.tolist(),
                         [("High", "crit", t.High.tolist()), ("Medium", "serious", t.Medium.tolist()),
                          ("Low", "warning", t.Low.tolist())], "exceptions", stacked=True, span=3)],
        "tables": [table("Top exceptions", ex.head(15), [
            ("severity", "Severity", "status"), ("exception_type", "Type", "text"), ("item_no", "Item", "text"),
            ("description", "Description", "text"), ("detail", "Detail", "text")], limit=15)],
    }


def render(m, view_id, filters):
    fn = globals()[view_id]
    out = fn(m, filters or {})
    out.setdefault("insights", [])
    if view_id == "overview":
        out.update({"id": "overview", "no": "", "title": "Overview",
                    "question": "The headline of every supply chain view on one page."})
        return out
    meta = next(v for v in VIEWS if v[1] == view_id)
    out.update({"id": view_id, "no": meta[0], "title": meta[2], "question": meta[3]})
    return out


# ---------------------------------------------------------------- item detail
def item_detail(m, item_no):
    it = m["item"].set_index("item_no")
    base = m["items"].set_index("item_no")
    if item_no not in base.index:
        return None
    row = it.loc[item_no].to_dict() if item_no in it.index else base.loc[item_no].to_dict()
    cls = row.get("item_class", "OTHER")
    sl = m["stock_loc"]
    pr = m["projection"]
    bom = m["bom"]
    pl = m["purchase_lines"]
    mine_sl = sl[sl.item_no == item_no].sort_values("stock_qty", ascending=False)
    mine_pl = pl[pl.item_no == item_no].sort_values("expected_date")
    iw = m["item_week"]
    iw = iw[iw.item_no == item_no].sort_values("week")
    weeks = sorted(m["weekly"].week.unique())
    iw = iw.set_index("week").reindex(weeks).fillna(0)
    used = (iw.consumed + (iw.sold if cls == "FG" else 0)).tolist()
    p = pr[pr.item_no == item_no]
    comps = bom[bom.bom_no == item_no].sort_values("can_make_now")
    parents = bom[bom.component_no == item_no]
    rd = m["readiness"]
    rdr = rd[rd.item_no == item_no]
    ins = []
    if row.get("cover_status") in ("Stockout", "Critical", "Low"):
        ins.append(insight("critical" if row.get("cover_status") != "Low" else "serious",
                           f"{row.get('cover_status')} cover ({_clean(row.get('cover_days')) or 0:,.1f} days) and supply is "
                           f"'{row.get('supply_status', '-')}'."))
    if row.get("stockout_date"):
        ins.append(insight("serious", f"Projected to run out on {row['stockout_date']}."))
    if len(mine_pl):
        ins.append(insight("info", f"{len(mine_pl)} open purchase line(s), {mine_pl.outstanding_qty.sum():,.0f} "
                                   f"{row.get('uom', '')} outstanding; {int(mine_pl.overdue.sum())} overdue."))
    if len(rdr):
        r0 = rdr.iloc[0]
        ins.append(insight("info", f"Can make {r0.can_make_now:,.0f} units now ({r0.production_days_supported:,.1f} days of "
                                   f"output); limited by {r0.limiting_component} {str(r0.limiting_description)[:40]}."))
    if len(parents):
        ins.append(insight("info", f"Used in {parents.bom_no.nunique()} certified BOM(s)."))
    charts = []
    if sum(used) or iw.output.sum() or iw.received.sum():
        charts.append(chart("bar", "Weekly movement (qty)", [str(w) for w in weeks],
                            [("Used / sold", "serious", used), ("Produced", "good", iw.output.tolist()),
                             ("Received", "seq", iw.received.tolist())], row.get("uom") or "qty"))
    if len(p):
        receipts = [0.0] * len(p)
        for r in mine_pl[~mine_pl.stale].itertuples():
            d = 7 if r.overdue else max(-r.days_overdue, 0)
            receipts[min(int(d // 7), len(p) - 1)] += float(r.outstanding_qty)
        charts.append(chart("line", "Projected stock, next 12 weeks", [str(x) for x in p.week_date],
                            [("Projected stock", cls if cls in CLS else "seq", p.projected_qty.tolist()),
                             ("Expected receipts", "good", receipts)], row.get("uom") or "qty",
                            extra={"zeroLine": True, "seriesTypes": ["line", "bar"]}))
    if len(mine_sl):
        charts.append(chart("hbar", "Stock by location", mine_sl.location.tolist(),
                            [("Stock", cls if cls in CLS else "seq", mine_sl.stock_qty.tolist())], row.get("uom") or "qty"))
    if len(comps):
        cc = comps.head(12)
        charts.append(chart("hbar", "Units each component supports", cc.component_no.tolist(),
                            [("Parent units", "class:" + ",".join(cc.component_class.tolist()), cc.can_make_now.tolist())],
                            "units"))
    return {
        "item_no": item_no, "summary": {k: _clean(v) for k, v in row.items()},
        "insights": ins, "charts": charts, "spark": used,
        "stock_by_location": table("Stock by location", mine_sl,
                                   [("location", "Location", "text"), ("location_name", "Name", "text"),
                                    ("site", "Site", "text"), ("stock_qty", "Stock", "qty"),
                                    ("stock_value", "Value", "money")], drill=None),
        "purchase_lines": table("Open purchase lines", mine_pl,
                                [("po_no", "PO", "text"), ("vendor_name", "Vendor", "text"),
                                 ("outstanding_qty", "Outstanding", "qty"), ("expected_date", "Expected", "date"),
                                 ("days_overdue", "Days late", "qty"), ("receipt_status", "Status", "status")], drill=None),
        "projection": chart("line", "Projected stock (qty)", p.week_date.tolist(),
                            [("Projected", cls, p.projected_qty.tolist())]),
        "bom_components": table("BOM components", comps,
                                [("component_no", "Component", "text"), ("component_description", "Description", "text"),
                                 ("qty_per_eff", "Qty per (incl. scrap)", "qty"), ("component_stock", "Stock", "qty"),
                                 ("can_make_now", "Parents it supports", "qty")], drill=None),
        "used_in": table("Used in BOMs", parents,
                         [("bom_no", "Parent", "text"), ("parent_description", "Description", "text"),
                          ("qty_per_eff", "Qty per", "qty")], drill=None),
    }
