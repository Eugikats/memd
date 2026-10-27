"""System data the rebuild may write: item measures, stock by location, monthly history, ledger rows."""
from __future__ import annotations

import math
from datetime import date

ITEM_FIELDS = {
    "stock_qty": "Stock on hand, all locations (open item ledger entries)",
    "in_transit_stock": "Stock sitting in in-transit locations",
    "incoming_qty": "Open purchase-order quantity still to arrive (in transit / on order)",
    "overdue_qty": "Open PO quantity past its expected receipt date",
    "daily_usage": "Average daily usage over the reporting window (calendar days)",
    "daily_usage_prod": "Average daily usage per production day",
    "cover_days": "Days of cover = stock / daily usage",
    "cover_prod_days": "Production days of cover = stock / usage per production day",
    "unit_cost": "Unit cost (LCY)",
    "stock_value": "Stock value = stock x unit cost",
    "consumed": "Quantity consumed in the reporting window",
    "sold": "Quantity sold in the reporting window",
    "received": "Quantity received from purchases in the reporting window",
    "output": "Quantity output from production in the reporting window",
    "usage_qty": "Usage quantity in the reporting window (RM/WIP consumption, FG sales)",
    "projected_end_qty": "Projected stock at the end of the projection horizon",
    "cover_status": "Cover status (Stockout / Critical / Low / Healthy / Excess)",
    "supply_status": "Supply status of open purchase orders",
    "risk_level": "Risk level (High / Medium / Low)",
    "abc": "ABC class on usage value",
    "avg_monthly_consumed": "Average monthly consumption over complete months of the 12-month history",
    "max_monthly_consumed": "Highest monthly consumption in the 12-month history",
    "avg_monthly_sold": "Average monthly sales quantity over complete months of the 12-month history",
    "max_monthly_sold": "Highest monthly sales quantity in the 12-month history",
    "description": "Item description (BC)",
    "uom": "Base unit of measure",
    "item_class": "Item class (RM / WIP / FG)",
    "posting_group": "Inventory posting group",
    "category": "Item category code",
}
TEXT_FIELDS = {"description", "uom", "item_class", "posting_group", "category", "cover_status",
               "supply_status", "risk_level", "abc"}
ZERO_WHEN_IDLE = {"stock_qty", "in_transit_stock", "incoming_qty", "overdue_qty", "daily_usage", "daily_usage_prod",
                  "consumed", "sold", "received", "output", "usage_qty", "stock_value",
                  "avg_monthly_consumed", "max_monthly_consumed", "avg_monthly_sold", "max_monthly_sold"}

MONTH_MEASURES = {
    "consumed": "Quantity consumed (Consumption, Assembly Consumption) in the month",
    "sold": "Quantity sold (Sale entries) in the month",
    "received": "Quantity received from purchases in the month",
    "output": "Production / assembly output in the month",
    "abs_qty": "Sum of absolute ledger quantities of all entry types in the month",
    "net_qty": "Net ledger quantity (all entry types) in the month",
}

LEDGER_FIELDS = {
    "posting_date": "Posting date (date)",
    "entry_type": "Entry type (Consumption, Sale, Purchase, Output, Positive Adjmt., ...)",
    "document_type": "Document type",
    "document_no": "Document number",
    "item_no": "Item number",
    "description": "Item description",
    "location": "Location code",
    "quantity": "Signed quantity",
    "abs_quantity": "Absolute quantity",
    "remaining_quantity": "Remaining quantity (open entries)",
    "cost_amount": "Cost amount (actual)",
    "sales_amount": "Sales amount (actual)",
    "entry_no": "Ledger entry number",
    "item_category": "Item category code",
    "open": "Entry is open (TRUE/FALSE)",
    "month_name": "Month name of posting date (January)",
    "month_year": "Month-year of posting date (Jan-26)",
    "year": "Year of posting date (2026)",
    "weekday": "Weekday of posting date (Mon)",
}

_LEDGER_SRC = {"posting_date": "Posting_Date", "entry_type": "Entry_Type", "document_type": "Document_Type",
               "document_no": "Document_No", "item_no": "Item_No", "description": "Item_Description",
               "location": "Location_Code", "quantity": "Quantity", "remaining_quantity": "Remaining_Quantity",
               "cost_amount": "Cost_Amount_Actual", "sales_amount": "Sales_Amount_Actual", "entry_no": "Entry_No",
               "item_category": "Item_Category_Code", "open": "Open"}


def _clean(v):
    if v is None:
        return None
    if isinstance(v, float) and math.isnan(v):
        return None
    try:
        import pandas as pd
        if v is pd.NaT:
            return None
    except Exception:
        pass
    if hasattr(v, "item"):
        try:
            v = v.item()
        except Exception:
            pass
    return v


class Sources:
    def __init__(self, model):
        self.model = model
        notes = model.get("notes") or {}
        self.reporting_date = date.fromisoformat(str(notes.get("reporting_date"))[:10]) \
            if notes.get("reporting_date") else date.today()
        self.history = notes.get("history") or {}
        items = model["items"]
        self.master = items.set_index("item_no").to_dict("index")
        self.measures = model["item"].set_index("item_no").to_dict("index")
        sl = model["stock_loc"]
        self.stock_loc = sl.groupby(["item_no", "location"]).stock_qty.sum().to_dict()
        self.locations = sorted(set(model["locations"].location.astype(str)) | set(sl.location.astype(str)))
        im = model.get("item_month")
        self.months = {}
        self.month_list = []
        self.monthly_stats = {}
        if im is not None and len(im):
            self.month_list = sorted(im.month.unique())
            for r in im.itertuples(index=False):
                self.months[(r.item_no, r.month)] = r
            complete = [m for m in self.month_list[1:] if m < self.reporting_date.isoformat()[:7]]
            sub = im[im.month.isin(complete)]
            if len(sub):
                g = sub.groupby("item_no")
                n = max(len(complete), 1)
                self.monthly_stats = {
                    "avg_monthly_consumed": (g.consumed.sum() / n).to_dict(),
                    "max_monthly_consumed": g.consumed.max().to_dict(),
                    "avg_monthly_sold": (g.sold.sum() / n).to_dict(),
                    "max_monthly_sold": g.sold.max().to_dict(),
                }
        self._ledger = None

    def known(self, item_no):
        return item_no in self.master or item_no in self.measures

    def item_value(self, field, item_no):
        if not self.known(item_no):
            return None
        if field in self.monthly_stats:
            v = self.monthly_stats[field].get(item_no)
            return _clean(v) if v is not None else 0.0
        row = self.measures.get(item_no)
        if row is not None and field in row:
            v = _clean(row[field])
            if v is None and field in ZERO_WHEN_IDLE:
                return 0.0
            return v
        m = self.master.get(item_no) or {}
        if field in m:
            return _clean(m[field])
        return 0.0 if field in ZERO_WHEN_IDLE else None

    def location_value(self, location, item_no):
        if not self.known(item_no):
            return None
        return float(self.stock_loc.get((item_no, location), 0.0))

    def month_value(self, measure, month, item_no):
        if not self.known(item_no):
            return None
        if month not in self.month_list:
            return None
        r = self.months.get((item_no, month))
        return float(getattr(r, measure)) if r is not None else 0.0

    def month_covered(self, month):
        return month in self.month_list

    def ledger(self, entry_types=None, prefixes=None):
        """12-month ledger rows (oldest first), filtered to the template's entry types and item prefixes."""
        if self._ledger is None:
            from pulse import model as M
            self._ledger = M.load_history()
        df = self._ledger
        if df.empty:
            return []
        if entry_types:
            df = df[df.Entry_Type.isin(entry_types)]
        if prefixes:
            df = df[df.Item_No.astype(str).str.startswith(tuple(prefixes))]
        df = df.sort_values(["Posting_Date", "Entry_No"])
        return df.to_dict("records")

    @staticmethod
    def ledger_value(field, row):
        if field == "abs_quantity":
            q = row.get("Quantity")
            return abs(q) if q is not None else None
        if field in ("month_name", "month_year", "year", "weekday"):
            try:
                d = date.fromisoformat(str(row.get("Posting_Date"))[:10])
            except ValueError:
                return None
            return {"month_name": d.strftime("%B"), "month_year": d.strftime("%b-%y"),
                    "year": d.strftime("%Y"), "weekday": d.strftime("%a")}[field]
        if field == "posting_date":
            try:
                return date.fromisoformat(str(row.get("Posting_Date"))[:10])
            except ValueError:
                return None
        src = _LEDGER_SRC.get(field)
        return _clean(row.get(src)) if src else None

    def catalog(self) -> dict:
        h = self.history or {}
        return {
            "item_fields": ITEM_FIELDS,
            "month_measures": MONTH_MEASURES,
            "ledger_fields": LEDGER_FIELDS,
            "locations": self.locations,
            "history_months": self.month_list,
            "history_first_date": h.get("first"),
            "history_last_date": h.get("last"),
            "reporting_date": self.reporting_date.isoformat(),
            "entry_types": sorted(self._entry_types()),
        }

    def _entry_types(self):
        try:
            from pulse import model as M
            if self._ledger is None:
                self._ledger = M.load_history()
            return set(self._ledger.Entry_Type.dropna().astype(str).unique()) if len(self._ledger) else set()
        except Exception:
            return set()
