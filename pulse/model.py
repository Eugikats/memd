"""Supply chain model: builds the governed views and measures from the staged Business Central data.

Every dashboard and every command-centre answer reads these views, so figures stay consistent.
Usage:  python -m pulse.model      (rebuilds data/model.pkl and prints reconciliation totals)
"""
import gzip
import json
import pickle
from datetime import date, timedelta

import numpy as np
import pandas as pd

from . import config as C
from .bc_client import DATA_DIR, ROOT

RAW = DATA_DIR / "raw"
MODEL_PATH = DATA_DIR / "model.pkl"

STALE_AFTER_DAYS = 90        # overdue PO lines older than this are treated as stale (not expected to arrive)
OVERDUE_ARRIVAL_DAYS = 7     # recently overdue lines are assumed to arrive this many days after the reporting date


def _load(name):
    with gzip.open(RAW / f"{name}.json.gz", "rt", encoding="utf-8") as f:
        return pd.DataFrame(json.load(f))


def _cover(stock, rate):
    return np.where(rate > 0, stock / np.where(rate > 0, rate, 1), np.nan)


def build():
    log = json.loads((RAW / "run_log.json").read_text())
    rep = date.fromisoformat(log["finished"][:10])
    window = int(log["window_days"])

    # ---------- dimensions ----------
    items = _load("items").rename(columns={
        "number": "item_no", "displayName": "description", "itemCategoryCode": "category",
        "inventoryPostingGroupCode": "posting_group", "baseUnitOfMeasureCode": "uom", "unitCost": "unit_cost"})
    items["item_class"] = items.posting_group.map(C.CLASS_BY_POSTING_GROUP).fillna("OTHER")
    bc_inventory = items.set_index("item_no").get("inventory")
    items = items[["item_no", "description", "item_class", "posting_group", "category", "uom", "unit_cost", "blocked"]]
    imap = items.set_index("item_no")

    locs = _load("locations").rename(columns={"Code": "location", "Name": "location_name"})
    locs["site"] = locs.location.map(C.site_of)
    locs["in_transit"] = locs.Use_As_In_Transit | locs.location.isin(C.IN_TRANSIT_LOCATIONS)
    locs = locs[["location", "location_name", "site", "in_transit", "Production_Floor"]].rename(
        columns={"Production_Floor": "production_floor"})

    def enrich(df):
        df = df.merge(items[["item_no", "description", "item_class", "posting_group", "uom", "unit_cost"]],
                      on="item_no", how="left")
        df["item_class"] = df.item_class.fillna("OTHER")
        return df

    # ---------- stock on hand (sum of remaining quantity on open ledger entries) ----------
    op = _load("ile_open").rename(columns={"Item_No": "item_no", "Location_Code": "location"})
    stock_loc = op.groupby(["item_no", "location"], as_index=False).Remaining_Quantity.sum().rename(
        columns={"Remaining_Quantity": "stock_qty"})
    stock_loc = enrich(stock_loc).merge(locs, on="location", how="left")
    stock_loc["site"] = stock_loc.site.fillna(stock_loc.location.map(C.site_of))
    stock_loc["in_transit"] = stock_loc.in_transit.fillna(False).astype(bool)
    stock_loc["stock_value"] = stock_loc.stock_qty * stock_loc.unit_cost.fillna(0)

    # ---------- movements over the reporting window ----------
    mv = _load("ile_movements").rename(columns={"Item_No": "item_no", "Location_Code": "location",
                                               "Posting_Date": "posting_date", "Entry_Type": "entry_type",
                                               "Quantity": "qty"})
    mv["site"] = mv.location.map(C.site_of)
    prod_days = max(mv.loc[mv.entry_type.isin({"Output", "Assembly Output"}), "posting_date"].nunique(), 1)
    mv["consumed"] = np.where(mv.entry_type.isin(C.CONSUMPTION_TYPES), -mv.qty, 0.0)
    mv["sold"] = np.where(mv.entry_type.isin(C.DEMAND_TYPES), -mv.qty, 0.0)
    mv["output"] = np.where(mv.entry_type.isin({"Output", "Assembly Output"}), mv.qty, 0.0)
    mv["received"] = np.where(mv.entry_type == "Purchase", mv.qty, 0.0)

    usage = mv.groupby("item_no")[["consumed", "sold", "output", "received"]].sum()
    usage_site = mv.groupby(["item_no", "site"])[["consumed", "sold"]].sum().reset_index()

    # Daily usage: RM/WIP by consumption, FG by outward sales + consumption (e.g. packed into sets).
    def usage_qty(df, cls):
        return np.where(cls == "FG", df.sold + df.consumed, df.consumed)

    # ---------- incoming supply (open purchase lines) ----------
    pl = _load("purchase_lines").rename(columns={
        "Document_No": "po_no", "Line_No": "line_no", "Buy_from_Vendor_No": "vendor_no",
        "Buy_from_Vendor_Name": "vendor_name", "No": "item_no", "Location_Code": "location",
        "Quantity": "ordered_qty", "Quantity_Received": "received_qty", "Outstanding_Quantity": "outstanding_qty",
        "Expected_Receipt_Date": "expected_date", "Status": "po_status",
        "Outstanding_Amount_LCY": "outstanding_value"})
    pl = pl.drop(columns=["Description"]).pipe(enrich)
    pl["expected_date"] = pd.to_datetime(pl.expected_date).dt.date
    pl["days_overdue"] = pl.expected_date.map(lambda d: (rep - d).days)
    pl["overdue"] = pl.days_overdue > 0
    pl["stale"] = pl.days_overdue > STALE_AFTER_DAYS
    pl["receipt_status"] = np.select([pl.stale, pl.overdue], ["Stale (>%dd overdue)" % STALE_AFTER_DAYS, "Overdue"],
                                     "Expected")
    pl["site"] = pl.location.map(C.site_of)
    # Effective arrival used in projections (days from reporting date)
    pl["arrival_day"] = np.where(pl.stale, np.nan,
                                 np.where(pl.overdue, OVERDUE_ARRIVAL_DAYS, -pl.days_overdue))
    pl = pl.sort_values(["overdue", "days_overdue"], ascending=[False, False])

    live = pl[~pl.stale]
    incoming = live.groupby("item_no").agg(incoming_qty=("outstanding_qty", "sum"),
                                           next_arrival_day=("arrival_day", "min"))
    overdue_in = pl[pl.overdue].groupby("item_no").outstanding_qty.sum().rename("overdue_qty")

    # ---------- item position (company level) ----------
    pos = stock_loc[~stock_loc.in_transit].groupby("item_no").agg(
        stock_qty=("stock_qty", "sum"), stock_value=("stock_value", "sum"), locations=("location", "nunique"))
    transit_stock = stock_loc[stock_loc.in_transit].groupby("item_no").stock_qty.sum().rename("in_transit_stock")
    keys = pos.index.union(usage.index).union(incoming.index)
    item = pd.DataFrame(index=keys).join([pos, transit_stock, usage, incoming, overdue_in]).fillna(
        {"stock_qty": 0, "stock_value": 0, "in_transit_stock": 0, "consumed": 0, "sold": 0, "output": 0,
         "received": 0, "incoming_qty": 0, "overdue_qty": 0, "locations": 0})
    item.index.name = "item_no"
    item = enrich(item.reset_index())
    item = item[item.item_class != "OTHER"].copy()
    item["usage_qty"] = usage_qty(item, item.item_class)
    item["daily_usage"] = item.usage_qty / window
    item["daily_usage_prod"] = item.usage_qty / prod_days
    item["cover_days"] = _cover(item.stock_qty, item.daily_usage)
    item["cover_prod_days"] = _cover(item.stock_qty, item.daily_usage_prod)
    item["daily_usage_value"] = item.daily_usage * item.unit_cost.fillna(0)

    # ---------- projection & stockout exposure ----------
    horizon = C.PROJECTION_WEEKS * 7
    arrivals = live.dropna(subset=["arrival_day"]).groupby(["item_no", "arrival_day"]).outstanding_qty.sum()
    proj_rows, so_day = [], {}
    days = np.arange(0, horizon + 1)
    for r in item[(item.daily_usage > 0)].itertuples():
        bal = r.stock_qty - r.daily_usage * days
        if r.item_no in arrivals.index.get_level_values(0):
            for d, q in arrivals.loc[r.item_no].items():
                bal[days >= d] += q
        hit = np.flatnonzero(bal <= 0)
        so_day[r.item_no] = int(hit[0]) if hit.size else np.nan
        for w in range(C.PROJECTION_WEEKS + 1):
            proj_rows.append((r.item_no, w, rep + timedelta(days=7 * w), float(bal[7 * w])))
    projection = pd.DataFrame(proj_rows, columns=["item_no", "week", "week_date", "projected_qty"])
    item["stockout_day"] = item.item_no.map(so_day)
    item["stockout_date"] = item.stockout_day.map(lambda d: rep + timedelta(days=int(d)) if pd.notna(d) else None)
    item["stockout_exposed"] = item.stockout_day.notna() & (
        item.next_arrival_day.isna() | (item.next_arrival_day > item.stockout_day))
    item["projected_end_qty"] = item.item_no.map(
        projection[projection.week == C.PROJECTION_WEEKS].set_index("item_no").projected_qty)

    # ---------- inventory risk ----------
    cov = item.cover_days
    # FG produced and shipped straight away (output covers >= 90% of sales) is make-to-order, not a stockout.
    mto = (item.item_class == "FG") & (item.stock_qty <= 0) & (item.output >= 0.9 * item.sold) & (item.sold > 0)
    item["cover_status"] = np.select(
        [mto, (item.daily_usage > 0) & (item.stock_qty <= 0), cov < C.COVER_CRITICAL, cov < C.COVER_LOW,
         cov > C.COVER_EXCESS, (item.daily_usage <= 0) & (item.stock_qty > 0)],
        ["Made to order", "Stockout", "Critical", "Low", "Excess", "No movement"], "Healthy")
    item.loc[mto, "stockout_exposed"] = False
    item["supply_status"] = np.select(
        [(item.incoming_qty <= 0) & (item.overdue_qty > 0), item.incoming_qty <= 0,
         item.stockout_exposed & (item.incoming_qty > 0), item.overdue_qty > 0],
        ["Only stale supply", "No open supply", "Supply arrives too late", "Overdue supply"], "On time")
    weak = item.supply_status != "On time"
    item["risk_level"] = np.select(
        [item.cover_status.isin(["Stockout", "Critical"]) & weak,
         item.cover_status.isin(["Stockout", "Critical"]) | (item.cover_status.eq("Low") & weak) | item.stockout_exposed,
         item.cover_status.eq("Low")],
        ["High", "Medium", "Low"], "None")
    sev = item.risk_level.map({"High": 3, "Medium": 2, "Low": 1, "None": 0})
    item["risk_score"] = sev * 1e12 + item.daily_usage_value.fillna(0)

    # ---------- location (site) supply risk and rebalancing ----------
    ss = stock_loc[~stock_loc.in_transit].groupby(["item_no", "site"]).stock_qty.sum().reset_index()
    site = ss.merge(usage_site, on=["item_no", "site"], how="outer").fillna({"stock_qty": 0, "consumed": 0, "sold": 0})
    site = enrich(site)
    site = site[site.item_class != "OTHER"].copy()
    site["daily_usage"] = usage_qty(site, site.item_class) / window
    site["cover_days"] = _cover(site.stock_qty, site.daily_usage)
    site["surplus_qty"] = np.maximum(site.stock_qty - site.daily_usage * C.COVER_TARGET * 2, 0)
    site["need_qty"] = np.maximum(site.daily_usage * C.COVER_TARGET - site.stock_qty, 0)
    at_risk = site[(site.daily_usage > 0) & (site.cover_days.fillna(0) < C.COVER_LOW)].copy()
    donors = site[site.surplus_qty > 0].sort_values("surplus_qty", ascending=False)
    best = donors.groupby("item_no").head(1).set_index("item_no")
    tot_surplus = donors.groupby("item_no").surplus_qty.sum()
    at_risk["donor_site"] = at_risk.item_no.map(best.site)
    at_risk["donor_surplus"] = at_risk.item_no.map(best.surplus_qty).fillna(0)
    at_risk.loc[at_risk.donor_site == at_risk.site, ["donor_site", "donor_surplus"]] = [None, 0]
    at_risk["suggested_transfer"] = np.minimum(at_risk.need_qty, at_risk.donor_surplus)
    at_risk["rebalance"] = np.where(at_risk.suggested_transfer >= at_risk.need_qty * 0.999, "Fully covered by transfer",
                                    np.where(at_risk.suggested_transfer > 0, "Partly covered by transfer",
                                             "No surplus elsewhere – purchase/produce"))
    at_risk["total_surplus_elsewhere"] = at_risk.item_no.map(tot_surplus).fillna(0)
    location_risk = at_risk.sort_values(["need_qty"], ascending=False)

    # ---------- BOM ----------
    bh = _load("bom_headers").rename(columns={"No": "bom_no", "Status": "bom_status"})
    bl = _load("bom_lines").rename(columns={"Production_BOM_No": "bom_no", "No": "component_no",
                                            "Quantity_per": "qty_per", "Scrap_Percent": "scrap_pct"})
    bl = bl[(bl.Type == "Item") & bl.bom_no.isin(bh[bh.bom_status == "Certified"].bom_no)]
    bom_zero_qty = int((bl.qty_per <= 0).sum())
    bom = bl[bl.qty_per > 0].copy()
    bom["qty_per_eff"] = bom.qty_per * (1 + bom.scrap_pct.fillna(0) / 100)
    bom["parent_class"] = bom.bom_no.map(imap.item_class)
    bom["component_class"] = bom.component_no.map(imap.item_class).fillna("OTHER")
    bom["parent_description"] = bom.bom_no.map(imap.description)
    bom["component_description"] = bom.component_no.map(imap.description)
    co = item.set_index("item_no")
    stock_all = stock_loc[~stock_loc.in_transit].groupby("item_no").stock_qty.sum()
    bom["component_stock"] = bom.component_no.map(stock_all).fillna(0).clip(lower=0)
    bom["component_incoming"] = bom.component_no.map(co.incoming_qty).fillna(0)
    bom["can_make_now"] = bom.component_stock / bom.qty_per_eff
    bom["can_make_with_incoming"] = (bom.component_stock + bom.component_incoming) / bom.qty_per_eff
    bom = bom[bom.component_class.isin(["RM", "WIP", "FG"])]

    out_rate = (usage.output / window).rename("daily_output")
    active = bom[bom.bom_no.isin(out_rate[out_rate > 0].index) & bom.parent_class.isin(["FG", "WIP"])]
    g = active.sort_values("can_make_now").groupby("bom_no")
    readiness = g.first()[["parent_description", "parent_class", "component_no", "component_description",
                           "component_class", "can_make_now"]].rename(columns={
        "component_no": "limiting_component", "component_description": "limiting_description",
        "component_class": "limiting_class"})
    readiness["can_make_with_incoming"] = g.can_make_with_incoming.min()
    readiness["components"] = g.size()
    readiness["daily_output"] = out_rate.reindex(readiness.index)
    readiness["production_days_supported"] = readiness.can_make_now / readiness.daily_output
    readiness["production_days_with_incoming"] = readiness.can_make_with_incoming / readiness.daily_output
    readiness["readiness"] = np.select(
        [readiness.production_days_supported < C.COVER_CRITICAL,
         readiness.production_days_supported < C.COVER_LOW],
        ["Constrained", "At risk"], "Ready")
    readiness = readiness.reset_index().rename(columns={"bom_no": "item_no"}).sort_values("production_days_supported")

    # WIP cover: complete parent units current WIP supports
    wb = bom[bom.component_class == "WIP"]
    wip_parent = (wb.assign(parent_output=wb.bom_no.map(out_rate).fillna(0))
                  .sort_values("parent_output", ascending=False))
    wip = item[item.item_class == "WIP"].copy()
    main_parent = wip_parent.groupby("component_no").head(1).set_index("component_no")
    wip["main_parent"] = wip.item_no.map(main_parent.bom_no)
    wip["main_parent_description"] = wip.item_no.map(main_parent.parent_description)
    wip["qty_per_parent"] = wip.item_no.map(main_parent.qty_per_eff)
    wip["parent_units_supported"] = wip.stock_qty.clip(lower=0) / wip.qty_per_parent
    wip["parent_count"] = wip.item_no.map(wb.groupby("component_no").bom_no.nunique()).fillna(0).astype(int)
    fg_from_wip = (wb.sort_values("can_make_now").groupby("bom_no")
                   .agg(units_from_wip=("can_make_now", "first"), limiting_wip=("component_no", "first"),
                        parent_description=("parent_description", "first"))
                   .join(out_rate).reset_index().rename(columns={"bom_no": "item_no"}))
    fg_from_wip["days_of_output"] = fg_from_wip.units_from_wip / fg_from_wip.daily_output

    # ---------- exceptions ----------
    ex = []
    for r in item[item.cover_status == "Stockout"].itertuples():
        ex.append(("Stockout", "High", r.item_no, r.description, r.item_class, "",
                   f"No stock; using {r.daily_usage:,.1f} {r.uom}/day. Supply: {r.supply_status}", r.daily_usage_value))
    for r in item[item.stockout_exposed & (item.cover_status != "Stockout")].itertuples():
        ex.append(("Stockout exposure", "High" if r.stockout_day <= C.COVER_CRITICAL else "Medium", r.item_no,
                   r.description, r.item_class, "",
                   f"Projected to run out {r.stockout_date} ({r.cover_days:,.0f} days cover); {r.supply_status.lower()}",
                   r.daily_usage_value))
    for r in item[item.cover_status.isin(["Critical", "Low"]) & ~item.stockout_exposed].itertuples():
        ex.append(("Low cover", "Medium" if r.cover_status == "Critical" else "Low", r.item_no, r.description,
                   r.item_class, "", f"{r.cover_days:,.1f} days cover; {r.supply_status.lower()}", r.daily_usage_value))
    for r in pl[pl.overdue & ~pl.stale].itertuples():
        ex.append(("Overdue receipt", "Medium" if r.item_class in ("RM", "WIP", "FG") else "Low", r.item_no,
                   r.description, r.item_class, r.location,
                   f"PO {r.po_no} from {r.vendor_name}: {r.outstanding_qty:,.0f} outstanding, {r.days_overdue} days late",
                   r.outstanding_value))
    for r in location_risk.head(200).itertuples():
        ex.append(("Location risk", "Medium" if r.cover_days < C.COVER_CRITICAL else "Low", r.item_no,
                   r.description, r.item_class, r.site,
                   f"{r.cover_days:,.1f} days cover at {r.site}; {r.rebalance}"
                   + (f" ({r.suggested_transfer:,.0f} from {r.donor_site})" if r.suggested_transfer > 0 else ""),
                   r.daily_usage * (r.unit_cost or 0)))
    for r in readiness[readiness.readiness == "Constrained"].itertuples():
        ex.append(("Production constraint", "High" if r.production_days_supported < 2 else "Medium", r.item_no,
                   r.parent_description, r.parent_class, "",
                   f"{r.production_days_supported:,.1f} days of output; limited by {r.limiting_component} "
                   f"({r.limiting_description})", 0.0))
    exceptions = pd.DataFrame(ex, columns=["exception_type", "severity", "item_no", "description", "item_class",
                                           "location", "detail", "value_at_stake"])
    exceptions["sev_rank"] = exceptions.severity.map({"High": 0, "Medium": 1, "Low": 2})
    exceptions = exceptions.sort_values(["sev_rank", "value_at_stake"], ascending=[True, False]).drop(columns="sev_rank")

    # ---------- data notes / reconciliation ----------
    recon = None
    if bc_inventory is not None:
        ours = stock_loc.groupby("item_no").stock_qty.sum()
        bci = bc_inventory[bc_inventory != 0]
        j = pd.concat([bci.rename("bc"), ours.rename("pulse")], axis=1).fillna(0)
        j = j[(j.bc != 0) | (j.pulse != 0)]
        recon = {"items_compared": int(len(j)), "items_matching": int(((j.bc - j.pulse).abs() < 1e-4).sum()),
                 "bc_total_qty": float(bci.sum()), "pulse_total_qty": float(ours.sum()),
                 "mismatches": j[(j.bc - j.pulse).abs() >= 1e-4].reset_index(names="item_no").head(50).to_dict("records")}
    notes = {
        "reconciliation": recon,
        "reporting_date": rep.isoformat(), "window_days": window, "production_days": int(prod_days),
        "company": log["company"], "environment": log["environment"], "extracted_at": log["finished"],
        "open_ledger_entries": int(len(op)), "movement_entries": int(len(mv)),
        "stock_qty_total": float(op.Remaining_Quantity.sum()),
        "purchase_lines_outstanding": int(len(pl)), "purchase_lines_overdue": int(pl.overdue.sum()),
        "purchase_lines_stale": int(pl.stale.sum()),
        "bom_lines_certified": int(len(bl)), "bom_lines_zero_qty": bom_zero_qty,
        "items_unclassified_with_stock": int(stock_loc[stock_loc.item_class == "OTHER"].item_no.nunique()),
        "stale_after_days": STALE_AFTER_DAYS, "overdue_arrival_days": OVERDUE_ARRIVAL_DAYS,
    }
    # Weekly flow by class (for trend reporting)
    wk = mv.merge(items[["item_no", "item_class", "unit_cost"]], on="item_no", how="left")
    wk = wk[wk.item_class.isin(["RM", "WIP", "FG"])]
    wk["week"] = pd.to_datetime(wk.posting_date).dt.to_period("W-SUN").dt.start_time.dt.date
    for c in ("consumed", "sold", "output", "received"):
        wk[c + "_value"] = wk[c] * wk.unit_cost.fillna(0)
    weekly = wk.groupby(["week", "item_class"])[["consumed_value", "sold_value", "output_value",
                                                  "received_value"]].sum().reset_index()
    wk = wk.merge(items[["item_no", "posting_group"]], on="item_no", how="left")
    weekly_pg = wk.groupby(["week", "item_class", "posting_group"])[["consumed_value", "sold_value", "output_value",
                                                                      "received_value"]].sum().reset_index()
    # Per-item weekly quantities (sparklines and item detail)
    item_week = wk.groupby(["item_no", "week"])[["consumed", "sold", "output", "received"]].sum().reset_index()

    # ABC classification on usage value within each class (A = top 80% of usage value, B = next 15%, C = rest)
    item["usage_value"] = item.usage_qty * item.unit_cost.fillna(0)
    item = item.sort_values(["item_class", "usage_value"], ascending=[True, False])
    cum = item.groupby("item_class").usage_value.cumsum() / item.groupby("item_class").usage_value.transform("sum").replace(0, np.nan)
    item["abc"] = np.select([item.usage_value <= 0, cum <= 0.80, cum <= 0.95], ["C", "A", "B"], "C")

    item_month, history = _item_month(rep)
    notes["history"] = history
    model = dict(weekly=weekly, weekly_pg=weekly_pg, item_week=item_week, item_month=item_month,
                 items=items, locations=locs, stock_loc=stock_loc, item=item, projection=projection,
                 purchase_lines=pl, site=site, location_risk=location_risk, bom=bom, readiness=readiness,
                 wip=wip, fg_from_wip=fg_from_wip, exceptions=exceptions, notes=notes)
    MODEL_PATH.parent.mkdir(parents=True, exist_ok=True)
    MODEL_PATH.write_bytes(pickle.dumps(model))
    _snapshot(model)
    return model


def load_history():
    """12-month item ledger entries (Excel ledger dumps). Empty frame when the extract predates the history pull."""
    if not (RAW / "ile_history.json.gz").exists():
        return pd.DataFrame(columns=["Entry_No", "Entry_Type", "Item_No", "Location_Code", "Posting_Date", "Quantity",
                                     "Document_No", "Document_Type", "Item_Description", "Item_Category_Code",
                                     "Remaining_Quantity", "Open", "Cost_Amount_Actual", "Sales_Amount_Actual"])
    return _load("ile_history")


def _item_month(rep):
    """Item × calendar month quantities over the 12-month history (separate from the 90-day cover window)."""
    h = load_history()
    cols = ["item_no", "month", "consumed", "sold", "received", "output", "abs_qty", "net_qty"]
    if h.empty:
        return pd.DataFrame(columns=cols), {"loaded": False, "first": None, "last": None, "rows": 0}
    h = h.rename(columns={"Item_No": "item_no", "Entry_Type": "entry_type", "Quantity": "qty",
                          "Posting_Date": "posting_date"})
    h["month"] = h.posting_date.str[:7]
    h["consumed"] = np.where(h.entry_type.isin(C.CONSUMPTION_TYPES), -h.qty, 0.0)
    h["sold"] = np.where(h.entry_type.isin(C.DEMAND_TYPES), -h.qty, 0.0)
    h["received"] = np.where(h.entry_type == "Purchase", h.qty, 0.0)
    h["output"] = np.where(h.entry_type.isin({"Output", "Assembly Output"}), h.qty, 0.0)
    h["abs_qty"] = h.qty.abs()
    h["net_qty"] = h.qty
    im = h.groupby(["item_no", "month"])[cols[2:]].sum().reset_index()
    first, last = h.posting_date.min(), h.posting_date.max()
    return im, {"loaded": True, "first": first, "last": last, "rows": int(len(h)),
                "reporting_date": rep.isoformat()}


def _snapshot(m):
    """Append one KPI row per data load, so trends build up over the proof period."""
    it, pl = m["item"], m["purchase_lines"]
    val = m["stock_loc"][~m["stock_loc"].in_transit].groupby("item_class").stock_value.sum()
    row = {"extracted_at": m["notes"]["extracted_at"], "reporting_date": m["notes"]["reporting_date"],
           "rm_value": float(val.get("RM", 0)), "wip_value": float(val.get("WIP", 0)), "fg_value": float(val.get("FG", 0)),
           "out_of_stock": int((it.cover_status == "Stockout").sum()),
           "run_out_7d": int((it.stockout_exposed & (it.cover_status != "Stockout") & (it.stockout_day <= 7)).sum()),
           "high_risk": int((it.risk_level == "High").sum()),
           "overdue_lines": int((pl.overdue & ~pl.stale).sum()),
           "overdue_value": float(pl[pl.overdue & ~pl.stale].outstanding_value.sum()),
           "constrained": int((m["readiness"].readiness == "Constrained").sum()),
           "high_exceptions": int((m["exceptions"].severity == "High").sum()),
           "recon_pct": (round(100 * m["notes"]["reconciliation"]["items_matching"] /
                               max(m["notes"]["reconciliation"]["items_compared"], 1), 2)
                         if m["notes"].get("reconciliation") else None)}
    path = DATA_DIR / "kpi_history.jsonl"
    hist = [json.loads(x) for x in path.read_text(encoding="utf-8").splitlines()] if path.exists() else []
    hist = [h for h in hist if h["extracted_at"] != row["extracted_at"]] + [row]
    path.write_text("\n".join(json.dumps(h) for h in hist) + "\n", encoding="utf-8")


def load():
    bundled = ROOT / "data" / "raw"
    if not (RAW / "run_log.json").exists() and (bundled / "run_log.json").exists() and bundled != RAW:
        import shutil
        RAW.mkdir(parents=True, exist_ok=True)
        for f in bundled.iterdir():
            shutil.copy2(f, RAW / f.name)
    elif bundled != RAW and (bundled / "ile_history.json.gz").exists() and not (RAW / "ile_history.json.gz").exists():
        # Older persistent extracts have no 12-month history yet; seed it until the next BC refresh.
        import shutil
        shutil.copy2(bundled / "ile_history.json.gz", RAW / "ile_history.json.gz")
        MODEL_PATH.unlink(missing_ok=True)
    if MODEL_PATH.exists():
        try:
            m = pickle.loads(MODEL_PATH.read_bytes())
            if "item_week" in m and "item_month" in m:   # cache from an older version is rebuilt
                return m
        except Exception:
            pass
    return build()


if __name__ == "__main__":
    m = build()
    print(json.dumps(m["notes"], indent=2))
    it = m["item"]
    print(it.groupby("item_class").agg(items=("item_no", "size"), stock_value=("stock_value", "sum")))
    print(it.risk_level.value_counts())
    print(m["exceptions"].exception_type.value_counts())
    print(m["readiness"].readiness.value_counts())
