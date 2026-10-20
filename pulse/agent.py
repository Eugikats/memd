"""Pulse agent: action tools, alert rules, watchlist and multi-step playbooks.

Design rule (proposal ch. 3.3 / 5.1): the agent never writes to Business Central. Everything it produces
is a *proposed action* stored in Pulse (drafted vendor message, transfer plan, briefing, alert rule) that a
person approves, dismisses or exports. Alert rules are re-evaluated after every data refresh.
"""
import json
import uuid
from datetime import datetime

import pandas as pd

from . import config as C
from .bc_client import DATA_DIR

DATA = DATA_DIR
ACTIONS = DATA / "actions.json"
ALERTS = DATA / "alerts.json"
WATCH = DATA / "watchlist.json"


# ---------- tiny JSON stores ----------
def _read(p, default):
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError):
        return default


def _write(p, obj):
    DATA.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(obj, indent=2, default=str), encoding="utf-8")


def list_actions():
    return sorted(_read(ACTIONS, []), key=lambda a: a["created"], reverse=True)


def _add_action(kind, title, summary, payload):
    acts = _read(ACTIONS, [])
    a = {"id": uuid.uuid4().hex[:10], "kind": kind, "title": title, "summary": summary, "payload": payload,
         "status": "proposed", "created": datetime.now().isoformat(timespec="seconds"), "created_by": "Pulse agent"}
    acts.append(a)
    _write(ACTIONS, acts)
    return a


def set_action_status(aid, status, user=None):
    acts = _read(ACTIONS, [])
    for a in acts:
        if a["id"] == aid:
            a["status"] = status
            a["decided"] = datetime.now().isoformat(timespec="seconds")
            a["decided_by"] = user
            if a["kind"] == "alert":
                rules = _read(ALERTS, [])
                for r in rules:
                    if r["id"] == a["payload"]["rule_id"]:
                        r["active"] = status == "approved"
                _write(ALERTS, rules)
            _write(ACTIONS, acts)
            return a
    return None


def get_action(aid):
    return next((a for a in _read(ACTIONS, []) if a["id"] == aid), None)


# ---------- action tools ----------
def draft_vendor_expedite(m, vendor=None, max_vendors=5, include_stale=False, item_nos=None):
    """Draft one expedite message per vendor with overdue open PO lines (most valuable first)."""
    pl = m["purchase_lines"]
    od = pl[pl.overdue & (include_stale | ~pl.stale)]
    if vendor:
        od = od[od.vendor_name.str.contains(vendor.strip(), case=False, na=False, regex=False) | (od.vendor_no == vendor)]
    if item_nos:
        od = od[od.item_no.isin([i.upper() for i in item_nos])]
    if od.empty:
        return {"created": [], "message": "No overdue purchase lines match."}
    by = od.groupby(["vendor_no", "vendor_name"]).outstanding_value.sum().sort_values(ascending=False)
    created = []
    for (vno, vname), val in by.head(max_vendors).items():
        lines = od[od.vendor_no == vno].sort_values("days_overdue", ascending=False)
        rows = [{"po_no": r.po_no, "item_no": r.item_no, "description": r.description,
                 "outstanding_qty": round(float(r.outstanding_qty), 2), "expected_date": str(r.expected_date),
                 "days_overdue": int(r.days_overdue)} for r in lines.itertuples()]
        body = (f"Dear {vname.strip().title()},\n\nOur records show the following purchase order lines are past "
                f"their expected receipt date:\n\n"
                + "\n".join(f"- PO {r['po_no']}: {r['item_no']} {r['description']} – {r['outstanding_qty']:,} "
                            f"outstanding, expected {r['expected_date']} ({r['days_overdue']} days late)"
                            for r in rows[:15])
                + ("\n- …and further lines (see attached)" if len(rows) > 15 else "")
                + "\n\nPlease confirm a revised delivery date for each line, or let us know if any line should be "
                  "cancelled.\n\nKind regards,\nSupply Chain – Nice House of Plastics")
        created.append(_add_action(
            "expedite", f"Expedite overdue POs – {vname.strip()}",
            f"{len(rows)} overdue line(s) worth {val:,.0f} UGX",
            {"vendor_no": vno, "vendor_name": vname, "value": float(val), "lines": rows,
             "subject": f"Overdue purchase orders – request for revised delivery dates", "body": body}))
    return {"created": created}


def propose_transfers(m, item_class=None, site=None, max_rows=40):
    """Turn location shortages that another site can cover into a proposed transfer plan."""
    lr = m["location_risk"]
    lr = lr[lr.suggested_transfer > 0]
    if item_class:
        lr = lr[lr.item_class == item_class]
    if site:
        lr = lr[lr.site == site]
    if lr.empty:
        return {"created": [], "message": "No shortages that another site can cover."}
    lr = lr.assign(val=lr.suggested_transfer * lr.unit_cost.fillna(0)).sort_values("val", ascending=False).head(max_rows)
    rows = [{"item_no": r.item_no, "description": r.description, "item_class": r.item_class,
             "from_site": r.donor_site, "to_site": r.site, "qty": round(float(r.suggested_transfer), 2),
             "uom": r.uom, "to_site_cover_days": None if pd.isna(r.cover_days) else round(float(r.cover_days), 1)}
            for r in lr.itertuples()]
    a = _add_action("transfer", f"Transfer plan – {len(rows)} lines" + (f" to {site}" if site else ""),
                    f"Rebalance stock between sites; value {lr.val.sum():,.0f} UGX", {"rows": rows})
    return {"created": [a]}


ALERT_METRICS = {
    "stockout_7d": "Items projected to run out within 7 days (before replenishment)",
    "overdue_lines": "Overdue purchase order lines (≤90 days)",
    "high_exceptions": "High-severity exceptions",
    "constrained_products": "Products with <7 days of components",
    "item_cover": "Cover days for one item (needs item_no)",
}


def metric_value(m, metric, item_no=None):
    it = m["item"]
    if metric == "stockout_7d":
        return int((it.stockout_exposed & (it.cover_status != "Stockout") & (it.stockout_day <= 7)).sum())
    if metric == "overdue_lines":
        pl = m["purchase_lines"]
        return int((pl.overdue & ~pl.stale).sum())
    if metric == "high_exceptions":
        return int((m["exceptions"].severity == "High").sum())
    if metric == "constrained_products":
        return int((m["readiness"].readiness == "Constrained").sum())
    if metric == "item_cover":
        r = it[it.item_no == (item_no or "").upper()]
        if r.empty or pd.isna(r.cover_days.iloc[0]):
            return None
        return round(float(r.cover_days.iloc[0]), 1)
    raise ValueError("unknown metric")


def create_alert(m, metric, op, threshold, item_no=None, name=None):
    if metric not in ALERT_METRICS or op not in (">", ">=", "<", "<="):
        return {"error": f"metric must be one of {list(ALERT_METRICS)} and op one of > >= < <="}
    rule = {"id": uuid.uuid4().hex[:8], "metric": metric, "op": op, "threshold": float(threshold),
            "item_no": item_no.upper() if item_no else None, "active": False,
            "name": name or f"{ALERT_METRICS[metric]}{' ' + item_no.upper() if item_no else ''} {op} {threshold:g}"}
    rules = _read(ALERTS, [])
    rules.append(rule)
    _write(ALERTS, rules)
    cur = metric_value(m, metric, item_no)
    a = _add_action("alert", f"Alert: {rule['name']}", f"Current value {cur}; checked after every refresh",
                    {"rule_id": rule["id"], "rule": rule, "current": cur})
    return {"created": [a], "current_value": cur}


def evaluate_alerts(m):
    out = []
    ops = {">": lambda a, b: a > b, ">=": lambda a, b: a >= b, "<": lambda a, b: a < b, "<=": lambda a, b: a <= b}
    rules = _read(ALERTS, [])
    for r in rules:
        v = metric_value(m, r["metric"], r.get("item_no"))
        r["last_value"] = v
        r["last_checked"] = datetime.now().isoformat(timespec="seconds")
        r["triggered"] = bool(v is not None and ops[r["op"]](v, r["threshold"]))
        if r.get("active"):
            out.append(r)
    _write(ALERTS, rules)
    return out


def list_alerts(m=None):
    if m is not None:
        evaluate_alerts(m)
    return _read(ALERTS, [])


def delete_alert(rid):
    _write(ALERTS, [r for r in _read(ALERTS, []) if r["id"] != rid])


def watch_items(m, item_nos):
    known = set(m["items"].item_no)
    w = _read(WATCH, [])
    added = [i.upper() for i in item_nos if i.upper() in known and i.upper() not in w]
    _write(WATCH, w + added)
    return {"added": added, "watchlist": w + added}


def watchlist(m):
    w = _read(WATCH, [])
    it = m["item"].set_index("item_no")
    rows = []
    for i in w:
        if i in it.index:
            r = it.loc[i]
            rows.append({"item_no": i, "description": r.description, "item_class": r.item_class,
                         "stock_qty": r.stock_qty, "cover_days": r.cover_days, "cover_status": r.cover_status,
                         "supply_status": r.supply_status, "risk_level": r.risk_level})
    return rows


def unwatch(item_no):
    _write(WATCH, [i for i in _read(WATCH, []) if i != item_no.upper()])


def generate_briefing(m, focus=None):
    """Management briefing in markdown, built only from model figures."""
    n, it, ex, pl, rd = m["notes"], m["item"], m["exceptions"], m["purchase_lines"], m["readiness"]
    soon = it[it.stockout_exposed & (it.cover_status != "Stockout") & (it.stockout_day <= 7)].sort_values(
        "daily_usage_value", ascending=False)
    out_now = it[it.cover_status == "Stockout"].sort_values("daily_usage_value", ascending=False)
    od = pl[pl.overdue & ~pl.stale]
    cons = rd[rd.readiness == "Constrained"].sort_values("daily_output", ascending=False)
    val = it.groupby("item_class").stock_value.sum()
    lr = m["location_risk"]

    def tbl(df, cols, heads):
        if df.empty:
            return "_None._\n"
        lines = ["| " + " | ".join(heads) + " |", "|" + "---|" * len(heads)]
        for r in df[cols].itertuples(index=False):
            lines.append("| " + " | ".join(f"{v:,.1f}" if isinstance(v, float) else str(v) for v in r) + " |")
        return "\n".join(lines) + "\n"

    md = [f"# Supply chain briefing – {n['company']}",
          f"_Reporting date {n['reporting_date']} · data loaded {n['extracted_at']} · usage window {n['window_days']} days_\n",
          "## Headlines",
          f"- Stock value: RM {val.get('RM', 0) / 1e9:,.2f} B, WIP {val.get('WIP', 0) / 1e6:,.0f} M, FG {val.get('FG', 0) / 1e9:,.2f} B UGX.",
          f"- **{len(out_now)}** items in use are out of stock; **{len(soon)}** more run out within 7 days before replenishment.",
          f"- **{len(od)}** purchase lines are overdue (≤90 days) worth {od.outstanding_value.sum() / 1e6:,.0f} M UGX; "
          f"{int(pl.stale.sum())} further lines are stale and should be closed or re-dated in BC.",
          f"- **{len(cons)}** actively produced items have under 7 days of components.",
          f"- **{int((lr.suggested_transfer > 0).sum())}** site shortages can be eased by transfers from other sites.",
          f"- {int((ex.severity == 'High').sum())} high-severity exceptions in total.\n",
          "## Runs out within 7 days (highest usage value)",
          tbl(soon.head(10).assign(d=soon.head(10).stockout_date.astype(str)),
              ["item_no", "description", "item_class", "cover_days", "d", "supply_status"],
              ["Item", "Description", "Class", "Cover days", "Runs out", "Supply"]),
          "## Production constraints",
          tbl(cons.head(10), ["item_no", "parent_description", "production_days_supported", "limiting_component",
                              "limiting_description"],
              ["Product", "Description", "Days supported", "Limiting component", "Component"]),
          "## Largest overdue purchase lines",
          tbl(od.sort_values("outstanding_value", ascending=False).head(10).assign(v=lambda d: d.outstanding_value / 1e6),
              ["po_no", "vendor_name", "item_no", "days_overdue", "v"],
              ["PO", "Vendor", "Item", "Days late", "Value (M UGX)"]),
          "## Recommended actions",
          "1. Expedite the overdue lines above with the vendors concerned.",
          "2. Approve the proposed inter-site transfers for shortages other sites can cover.",
          "3. Confirm purchase or production cover for items running out within 7 days.",
          "4. Close or re-date stale purchase lines and fix BOM lines with Quantity per = 0 in Business Central.",
          "\n_Generated by the Pulse agent from the governed supply chain model. No Business Central records were changed._"]
    text = "\n".join(md)
    a = _add_action("briefing", f"Management briefing – {n['reporting_date']}", "Ready to review and download",
                    {"markdown": text})
    return {"created": [a], "markdown": text}


def generate_report(m, report_id, params=None, user=None):
    """Produce a standard or saved report as an Excel file in report history, and offer it as an action."""
    from . import reports as R
    e = R.generate(m, report_id, params or {}, user=user or "Pulse agent", trigger="agent")
    a = _add_action("report", f"Report: {e['name']}", f"{e['rows']:,} rows · Excel ready to download",
                    {"history_id": e["id"], "report_id": report_id, "params": e["params"], "name": e["name"]})
    return {"created": [a], "report": e}


def match_report(text):
    """Best catalogue match for free text (offline mode)."""
    from . import reports as R
    words = set(w for w in text.lower().replace("&", " ").split() if len(w) > 2)
    best, score = "exceptions", 0
    for c in R.catalogue():
        hay = (c["name"] + " " + c["description"] + " " + c["id"].replace("_", " ")).lower()
        sc = sum(1 for w in words if w in hay)
        if sc > score:
            best, score = c["id"], sc
    return best


# ---------- offline playbooks (multi-step, deterministic) ----------
def run_playbook(m, name, args=None):
    """Returns {steps, actions, answer, ui}. Used when Claude is not configured, and as quick-start buttons."""
    args = args or {}
    steps, actions, ui = [], [], None

    def step(label, detail=""):
        steps.append({"label": label, "detail": detail})

    if name == "expedite":
        pl = m["purchase_lines"]
        od = pl[pl.overdue & ~pl.stale]
        step("Read open purchase lines", f"{len(pl)} outstanding, {len(od)} overdue within 90 days")
        r = draft_vendor_expedite(m, vendor=args.get("vendor"))
        actions += r["created"]
        step("Drafted vendor expedite messages", f"{len(r['created'])} vendor(s), most overdue value first")
        ui = {"view": "in_transit_overdue", "filters": {}}
        answer = (f"I reviewed {len(od)} overdue purchase lines and drafted **{len(r['created'])} expedite "
                  f"message(s)** for the vendors with the most overdue value. Review and approve them below — "
                  f"nothing is sent or changed in Business Central.")
    elif name == "rebalance":
        lr = m["location_risk"]
        step("Checked site cover", f"{len(lr)} item-site shortages under 14 days")
        r = propose_transfers(m, item_class=args.get("item_class"), site=args.get("site"))
        actions += r["created"]
        n_rows = len(r["created"][0]["payload"]["rows"]) if r["created"] else 0
        step("Matched shortages to surplus at other sites", f"{n_rows} transfer lines proposed")
        ui = {"view": "location_risk", "filters": {k: v for k, v in args.items() if v}}
        answer = (f"Of {len(lr)} site shortages, **{n_rows}** can be eased by moving surplus from another site. "
                  f"I've prepared a transfer plan for approval; export it to raise transfer orders in BC.")
    elif name == "briefing":
        step("Collected headline measures", "stock, stockouts, overdue receipts, production, locations")
        r = generate_briefing(m)
        actions += r["created"]
        step("Wrote management briefing", "ranked tables + recommended actions")
        answer = r["markdown"]
    elif name == "watch_critical":
        it = m["item"]
        soon = it[it.stockout_exposed & (it.cover_status != "Stockout") & (it.stockout_day <= 7)].sort_values(
            "daily_usage_value", ascending=False).head(10)
        step("Found items running out within 7 days", f"{len(soon)} highest by usage value")
        w = watch_items(m, soon.item_no.tolist())
        step("Added to watchlist", f"{len(w['added'])} new item(s)")
        r = create_alert(m, "stockout_7d", ">", 0, name="Any item projected to run out within 7 days")
        actions += r["created"]
        step("Proposed an alert rule", "re-checked after every refresh once approved")
        ui = {"view": "overview", "filters": {}}
        answer = (f"I added **{len(w['added'])}** critical items to your watchlist (shown on the Overview) and "
                  f"proposed an alert that fires whenever any item is projected to run out within 7 days "
                  f"(currently {r['current_value']}). Approve the alert to activate it.")
    elif name == "report":
        rid = args.get("report_id") or match_report(args.get("text") or "")
        step("Chose report", rid)
        r = generate_report(m, rid, {k: args.get(k) for k in ("item_class", "site") if args.get(k)})
        actions += r["created"]
        step("Generated Excel report", f"{r['report']['rows']:,} rows")
        ui = {"view": "reports", "filters": {}}
        answer = (f"I generated the **{r['report']['name']}** ({r['report']['rows']:,} rows). Download it below or "
                  f"from Reports → History.")
    elif name == "production_check":
        rd = m["readiness"]
        cons = rd[rd.readiness == "Constrained"]
        step("Checked BOM readiness", f"{len(rd)} active products, {len(cons)} constrained")
        comps = cons.limiting_component.value_counts().head(5)
        step("Ranked constraining components", ", ".join(f"{k} ({v})" for k, v in comps.items()))
        ui = {"view": "bom_readiness", "filters": {}}
        answer = (f"**{len(cons)}** products have under 7 days of components. The components constraining the "
                  f"most products are: " + ", ".join(f"**{k}** ({v} products)" for k, v in comps.items())
                  + ". I've opened BOM / Production Readiness for detail.")
    else:
        return None
    return {"steps": steps, "actions": actions, "answer": answer, "ui": ui}


PLAYBOOKS = [
    {"id": "briefing", "label": "Write today's management briefing", "icon": "doc"},
    {"id": "expedite", "label": "Chase overdue purchase orders", "icon": "mail"},
    {"id": "rebalance", "label": "Plan stock transfers between sites", "icon": "swap"},
    {"id": "watch_critical", "label": "Watch items about to run out", "icon": "bell"},
    {"id": "production_check", "label": "Check what's blocking production", "icon": "factory"},
]
