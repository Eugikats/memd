"""Supply Chain Command Centre: natural-language questions answered from the governed model.

Claude interprets the question and may only call read-only tools over the approved views, so every
figure in an answer is the same figure shown on the dashboards. Without ANTHROPIC_API_KEY, a simple
keyword router answers from the same views (offline demo mode).
"""
import json
import os
import re
from datetime import datetime

import anthropic

from . import agent as A
from . import config as C
from . import views as V
from .bc_client import DATA_DIR, ROOT

AUDIT = DATA_DIR / "audit.jsonl"
MODEL = "claude-opus-5-5"
MAX_ROWS = 25

GLOSSARY = """Mulwana / Nice House of Plastics supply chain glossary:
- RM = raw material (virgin resin, masterbatch, regrind, recycle, additives, packaging film/paper, auxiliaries, defects for reprocessing)
- WIP = semi-finished goods (e.g. battery covers, bottle bodies awaiting assembly/branding)
- FG = finished goods (crates, buckets, bottles, toothbrushes, pens, resale items)
- Sites group locations: "SHIP 4" = SHIP4RM (raw material store) + SHIP4PR (production floor) + SHIP4FG (finished goods store)
- "PR" locations are production floors, "RM" are raw material stores, "FG" are finished goods stores, "BOND" are bonded warehouses
- "Cover" means days the current stock lasts at the recent usage rate; "production days" excludes days without production
- Item numbers look like RM-00181, WP-00033, FG-00152, PM-00026"""


def _tools():
    view_ids = [v[1] for v in V.VIEWS]
    return [
        {"name": "query_view",
         "description": "Run one of the approved supply chain views and return its KPIs and the top rows of its "
                        "tables. Views: " + "; ".join(f"{v[1]} = {v[2]}: {v[3]}" for v in V.VIEWS),
         "input_schema": {"type": "object", "additionalProperties": False, "required": ["view"], "properties": {
             "view": {"type": "string", "enum": view_ids},
             "item_class": {"type": "string", "enum": ["RM", "WIP", "FG"], "description": "Filter by item class"},
             "site": {"type": "string", "description": "Filter by site, e.g. 'SHIP 4'"},
             "location": {"type": "string", "description": "Filter by location code, e.g. SHIP4RM"},
             "search": {"type": "string", "description": "Text filter on item number, description, vendor or PO"},
         }}},
        {"name": "get_item",
         "description": "Full detail for one item: position, cover, risk, stock by location, open purchase lines, "
                        "projection and BOM links.",
         "input_schema": {"type": "object", "additionalProperties": False, "required": ["item_no"],
                          "properties": {"item_no": {"type": "string"}}}},
        {"name": "find_items",
         "description": "Find item numbers whose description matches text (e.g. 'sabic resin', '10 ltr bucket').",
         "input_schema": {"type": "object", "additionalProperties": False, "required": ["text"],
                          "properties": {"text": {"type": "string"}}}},
        {"name": "open_dashboard",
         "description": "Open a dashboard for the user with filters applied (UI navigation only).",
         "input_schema": {"type": "object", "additionalProperties": False, "required": ["view"], "properties": {
             "view": {"type": "string", "enum": ["overview"] + view_ids},
             "item_class": {"type": "string", "enum": ["RM", "WIP", "FG"]},
             "site": {"type": "string"}, "search": {"type": "string"}}}},
        {"name": "draft_vendor_expedite",
         "description": "ACTION: draft expedite messages to vendors with overdue purchase lines (one per vendor, "
                        "highest overdue value first). Creates proposals for approval; nothing is sent.",
         "input_schema": {"type": "object", "additionalProperties": False, "properties": {
             "vendor": {"type": "string", "description": "Optional vendor name or number"},
             "item_nos": {"type": "array", "items": {"type": "string"}, "description": "Optional: only lines for these items"},
             "max_vendors": {"type": "integer", "minimum": 1, "maximum": 20}}}},
        {"name": "propose_transfers",
         "description": "ACTION: build a proposed inter-site transfer plan for shortages other sites can cover. "
                        "Creates a proposal for approval/export; nothing is posted to Business Central.",
         "input_schema": {"type": "object", "additionalProperties": False, "properties": {
             "item_class": {"type": "string", "enum": ["RM", "WIP", "FG"]}, "site": {"type": "string"}}}},
        {"name": "create_alert",
         "description": "ACTION: propose an alert rule re-checked after every data refresh. Metrics: "
                        + "; ".join(f"{k} = {v}" for k, v in A.ALERT_METRICS.items()),
         "input_schema": {"type": "object", "additionalProperties": False, "required": ["metric", "op", "threshold"],
                          "properties": {"metric": {"type": "string", "enum": list(A.ALERT_METRICS)},
                                         "op": {"type": "string", "enum": [">", ">=", "<", "<="]},
                                         "threshold": {"type": "number"}, "item_no": {"type": "string"},
                                         "name": {"type": "string"}}}},
        {"name": "watch_items",
         "description": "ACTION: add items to the user's watchlist shown on the Overview.",
         "input_schema": {"type": "object", "additionalProperties": False, "required": ["item_nos"],
                          "properties": {"item_nos": {"type": "array", "items": {"type": "string"}}}}},
        {"name": "generate_briefing",
         "description": "ACTION: write a management briefing (markdown) from the model and save it for review.",
         "input_schema": {"type": "object", "additionalProperties": False, "properties": {}}},
        {"name": "generate_report",
         "description": "ACTION: produce a report as a downloadable Excel file (also saved in report history). "
                        "Reports: " + "; ".join(f"{c[0]} = {c[1]}" for c in __import__("pulse.reports", fromlist=["x"]).CATALOGUE),
         "input_schema": {"type": "object", "additionalProperties": False, "required": ["report_id"], "properties": {
             "report_id": {"type": "string"},
             "item_class": {"type": "string", "enum": ["RM", "WIP", "FG"]}, "site": {"type": "string"}}}},
        {"name": "get_definitions",
         "description": "Return the documented definition of every measure and the data notes for this load.",
         "input_schema": {"type": "object", "additionalProperties": False, "properties": {}}},
    ]


def _trim_view(out):
    return {"view": out["title"], "kpis": out["kpis"],
            "tables": [{"title": t["title"], "total_rows": t["total_rows"],
                        "columns": [c["field"] for c in t["columns"]],
                        "rows": t["rows"][:MAX_ROWS]} for t in out["tables"]]}


def run_tool(m, name, args):
    if name == "query_view":
        if args.get("view") not in V.VIEW_IDS:
            return {"error": "Unknown view"}
        f = {"item_class": args.get("item_class"), "site": args.get("site"), "location": args.get("location"),
             "q": args.get("search")}
        return _trim_view(V.render(m, args["view"], f))
    if name == "get_item":
        d = V.item_detail(m, str(args.get("item_no", "")).upper().strip())
        if not d:
            return {"error": "Item not found"}
        for k in ("stock_by_location", "purchase_lines", "bom_components", "used_in"):
            d[k] = {"total_rows": d[k]["total_rows"], "rows": d[k]["rows"][:MAX_ROWS]}
        return d
    if name == "find_items":
        words = str(args.get("text", "")).lower().split()
        it = m["items"]
        mask = it.description.str.lower().fillna("")
        hit = it[mask.apply(lambda s: all(w in s for w in words))].head(30)
        return {"matches": hit[["item_no", "description", "item_class", "uom"]].to_dict("records")}
    if name == "open_dashboard":
        return {"opened": args.get("view"), "filters": {k: args.get(k) for k in ("item_class", "site", "search")
                                                        if args.get(k)}}
    if name == "draft_vendor_expedite":
        return A.draft_vendor_expedite(m, vendor=args.get("vendor"), max_vendors=int(args.get("max_vendors") or 5),
                                       item_nos=args.get("item_nos"))
    if name == "propose_transfers":
        return A.propose_transfers(m, item_class=args.get("item_class"), site=args.get("site"))
    if name == "create_alert":
        return A.create_alert(m, args["metric"], args["op"], args["threshold"], args.get("item_no"), args.get("name"))
    if name == "watch_items":
        return A.watch_items(m, args.get("item_nos") or [])
    if name == "generate_briefing":
        r = A.generate_briefing(m)
        return {"created": r["created"], "markdown": r["markdown"][:6000]}
    if name == "generate_report":
        return A.generate_report(m, args["report_id"], {k: args.get(k) for k in ("item_class", "site") if args.get(k)})
    if name == "get_definitions":
        n = dict(m["notes"])
        n.pop("reconciliation", None)
        return {"definitions": V.DEFINITIONS, "data_notes": n}
    return {"error": "Unknown tool"}


def _system(m):
    n = m["notes"]
    return (f"You are the Pulse Supply Chain Command Centre for {n['company']} (Business Central environment "
            f"{n['environment']}). Reporting date {n['reporting_date']}; usage rates use the last "
            f"{n['window_days']} days ({n['production_days']} production days). Amounts are in UGX.\n\n"
            "Answer supply chain questions using ONLY figures returned by your tools - never estimate or invent "
            "numbers. Your tools are read-only and expose only the approved views. If the question is ambiguous "
            "(e.g. an item name matching several items) or the data cannot answer it, say so and ask a short "
            "clarifying question instead of guessing. Lead with the direct answer, then the key supporting "
            "figures (a compact markdown table when listing items), then one line starting 'Basis:' naming the "
            "view(s) and measure definitions used.\n\n"
            "You are also an agent. When the user asks you to DO something (chase suppliers, plan transfers, set an "
            "alert, watch items, prepare a briefing, show a dashboard), plan the steps, gather the evidence with the "
            "read tools first, then call the ACTION tools. Action tools only create proposals inside Pulse for a "
            "person to approve - never say anything was sent, posted or changed in Business Central. After acting, "
            "summarise what you created and what the user should review. Use open_dashboard when showing a view "
            "would help.\n\n" + GLOSSARY)


def _audit(rec):
    AUDIT.parent.mkdir(parents=True, exist_ok=True)
    with AUDIT.open("a", encoding="utf-8") as f:
        f.write(json.dumps(rec, default=str) + "\n")


def provider():
    if os.environ.get("AZURE_OPENAI_ENDPOINT"):
        return "azure"
    if os.environ.get("ANTHROPIC_API_KEY"):
        return "claude"
    return "offline"


_azure_client = None


def _azure():
    """Azure OpenAI client: API key if set (local dev), otherwise the web app's managed identity."""
    global _azure_client
    if _azure_client is None:
        from openai import AzureOpenAI
        kw = {"azure_endpoint": os.environ["AZURE_OPENAI_ENDPOINT"],
              "api_version": os.environ.get("AZURE_OPENAI_API_VERSION", "2025-04-01-preview")}
        if os.environ.get("AZURE_OPENAI_API_KEY"):
            kw["api_key"] = os.environ["AZURE_OPENAI_API_KEY"]
        else:
            from azure.identity import DefaultAzureCredential, get_bearer_token_provider
            kw["azure_ad_token_provider"] = get_bearer_token_provider(
                DefaultAzureCredential(), "https://cognitiveservices.azure.com/.default")
        _azure_client = AzureOpenAI(**kw)
    return _azure_client


def _ask_azure(m, question, history):
    client = _azure()
    deployment = os.environ.get("AZURE_OPENAI_DEPLOYMENT", "gpt-5.4")
    tools = [{"type": "function", "function": {"name": t["name"], "description": t["description"],
                                               "parameters": t["input_schema"]}} for t in _tools()]
    messages = [{"role": "system", "content": _system(m)}]
    messages += [{"role": h["role"], "content": h["content"]} for h in history[-6:]]
    messages.append({"role": "user", "content": question})
    calls, last_view, actions, ui, steps = [], None, [], None, []
    for _ in range(12):
        resp = client.chat.completions.create(model=deployment, messages=messages, tools=tools,
                                              max_completion_tokens=8000)
        msg = resp.choices[0].message
        if not msg.tool_calls:
            return {"mode": "azure", "answer": (msg.content or "").strip(), "tool_calls": calls, "table": last_view,
                    "steps": steps, "actions": actions, "ui": ui}
        messages.append({"role": "assistant", "content": msg.content or "",
                         "tool_calls": [tc.model_dump() for tc in msg.tool_calls]})
        for tc in msg.tool_calls:
            try:
                args = json.loads(tc.function.arguments or "{}")
            except json.JSONDecodeError:
                args = {}
            name = tc.function.name
            try:
                out = run_tool(m, name, args)
            except Exception as e:  # tool errors go back to the model, not the user
                out = {"error": str(e)[:300]}
            calls.append({"tool": name, "input": args})
            steps.append({"label": STEP_LABELS.get(name, name), "detail": _brief(args)})
            if name == "query_view" and isinstance(out, dict) and out.get("tables"):
                last_view = out["tables"][0]
            if isinstance(out, dict) and out.get("created"):
                actions += out["created"]
            if name == "open_dashboard" and "opened" in out:
                ui = {"view": out["opened"], "filters": {"item_class": out["filters"].get("item_class"),
                                                         "site": out["filters"].get("site"),
                                                         "q": out["filters"].get("search")}}
            messages.append({"role": "tool", "tool_call_id": tc.id, "content": json.dumps(out, default=str)[:60000]})
    return {"mode": "azure", "answer": "I couldn't complete that within the allowed steps.", "tool_calls": calls,
            "table": last_view, "steps": steps, "actions": actions, "ui": ui}


def ask(m, question, user="demo", history=None):
    rec = {"at": datetime.now().isoformat(timespec="seconds"), "user": user, "question": question}
    if provider() == "azure":
        try:
            res = _ask_azure(m, question, history or [])
        except Exception as e:  # network/auth/quota: fall back so the user still gets an answer
            res = _ask_offline(m, question)
            res["answer"] = f"_Azure OpenAI unavailable ({type(e).__name__}: {str(e)[:160]}); offline answer:_\n\n" + res["answer"]
    elif os.environ.get("ANTHROPIC_API_KEY"):
        try:
            res = _ask_claude(m, question, history or [])
        except anthropic.APIError as e:
            res = _ask_offline(m, question)
            res["answer"] = f"_Claude unavailable ({type(e).__name__}); offline answer:_\n\n" + res["answer"]
    else:
        res = _ask_offline(m, question)
    res.setdefault("steps", [{"label": STEP_LABELS.get(c["tool"], c["tool"]), "detail": _brief(c.get("input"))}
                             for c in res["tool_calls"]])
    res.setdefault("actions", [])
    res.setdefault("ui", None)
    rec.update({"mode": res["mode"], "tool_calls": res["tool_calls"], "answer": res["answer"],
                "actions": [a["id"] for a in res["actions"]]})
    _audit(rec)
    return res


def _brief(inp):
    if not inp:
        return ""
    if isinstance(inp, dict):
        return ", ".join(f"{k}={v}" for k, v in inp.items() if v not in (None, "", []))[:120]
    return str(inp)[:120]


STEP_LABELS = {"query_view": "Queried view", "get_item": "Looked up item", "find_items": "Searched items",
               "get_definitions": "Read definitions", "open_dashboard": "Opened dashboard",
               "draft_vendor_expedite": "Drafted vendor expedite messages", "propose_transfers": "Proposed transfer plan",
               "create_alert": "Proposed alert rule", "watch_items": "Updated watchlist",
               "generate_briefing": "Wrote management briefing", "generate_report": "Generated Excel report"}


def _ask_claude(m, question, history):
    client = anthropic.Anthropic()
    messages = [{"role": h["role"], "content": h["content"]} for h in history[-6:]]
    messages.append({"role": "user", "content": question})
    calls, last_view, actions, ui, steps = [], None, [], None, []
    for _ in range(12):
        resp = client.beta.messages.create(
            model=MODEL, max_tokens=16000, system=_system(m), tools=_tools(), messages=messages,
            thinking={"type": "adaptive"}, output_config={"effort": "medium"},
            betas=["server-side-fallback-2026-07-01"], fallbacks="default")
        if resp.stop_reason == "refusal":
            return {"mode": "claude", "answer": "I can't help with that question.", "tool_calls": calls,
                    "table": None, "steps": steps, "actions": actions, "ui": ui}
        messages.append({"role": "assistant", "content": resp.content})
        uses = [b for b in resp.content if b.type == "tool_use"]
        if resp.stop_reason != "tool_use" or not uses:
            text = "\n".join(b.text for b in resp.content if b.type == "text").strip()
            return {"mode": "claude", "answer": text, "tool_calls": calls, "table": last_view, "steps": steps,
                    "actions": actions, "ui": ui}
        results = []
        for u in uses:
            out = run_tool(m, u.name, u.input)
            calls.append({"tool": u.name, "input": u.input})
            steps.append({"label": STEP_LABELS.get(u.name, u.name), "detail": _brief(u.input)})
            if u.name == "query_view" and "tables" in out and out["tables"]:
                last_view = out["tables"][0]
            if isinstance(out, dict) and out.get("created"):
                actions += out["created"]
            if u.name == "open_dashboard":
                ui = {"view": out["opened"], "filters": {"item_class": out["filters"].get("item_class"),
                                                         "site": out["filters"].get("site"),
                                                         "q": out["filters"].get("search")}}
            results.append({"type": "tool_result", "tool_use_id": u.id,
                            "content": json.dumps(out, default=str), "is_error": "error" in out})
        messages.append({"role": "user", "content": results})
    return {"mode": "claude", "answer": "I couldn't complete that question within the allowed steps.",
            "tool_calls": calls, "table": last_view, "steps": steps, "actions": actions, "ui": ui}


# ---------- offline keyword router (no API key) ----------
ROUTES = [
    (r"exception|attention|management|priorit", "exceptions"),
    (r"overdue|late|in.?transit|purchase|po\b|vendor|supplier|incoming|on order", "in_transit_overdue"),
    (r"run out|stock.?out|project|forecast|future|next \d+ (day|week)", "projection_stockout"),
    (r"rebalanc|transfer|location risk|site risk|other location|elsewhere", "location_risk"),
    (r"bom|produc(e|tion) readiness|can we (make|produce)|constrain|bottleneck|component", "bom_readiness"),
    (r"wip|work in progress|semi", "wip_cover"),
    (r"risk|attention", "inventory_risk"),
    (r"raw material|\brm\b|consum|resin|masterbatch", "rm_consumption_cover"),
    (r"finished|\bfg\b|demand|sales", "fg_cover"),
    (r"where|location|warehouse|store|site", "inventory_by_location"),
]


INTENTS = [
    (r"brief|write (a|the|today)", "briefing"),
    (r"\b(excel|export|spreadsheet|report)\b", "report"),
    (r"chase|expedite|follow.?up|email|write to (the )?(vendor|supplier)|remind (the )?(vendor|supplier)", "expedite"),
    (r"plan (stock )?transfer|rebalance|move stock|transfer plan", "rebalance"),
    (r"watch|alert|notify|monitor|warn me", "watch_critical"),
    (r"blocking production|production check", "production_check"),
]


def run_playbook(m, name, user="demo", args=None):
    r = A.run_playbook(m, name, args)
    if r is None:
        return None
    res = {"mode": "agent", "answer": r["answer"], "tool_calls": [{"tool": "playbook", "input": {"name": name}}],
           "table": None, "steps": r["steps"], "actions": r["actions"], "ui": r["ui"]}
    _audit({"at": datetime.now().isoformat(timespec="seconds"), "user": user, "question": f"[playbook] {name}",
            "mode": "agent", "answer": r["answer"][:2000], "actions": [a["id"] for a in r["actions"]]})
    return res


def _ask_offline(m, question):
    q = question.lower()
    for pat, pb in INTENTS:
        if re.search(pat, q):
            site = re.search(r"ship\s?(\d)", q)
            cls = "RM" if re.search(r"raw material|\brm\b", q) else "FG" if re.search(r"finished|\bfg\b", q) else None
            r = A.run_playbook(m, pb, {"site": f"SHIP {site.group(1)}" if site else None, "item_class": cls,
                                       "text": question})
            return {"mode": "agent", "answer": r["answer"], "tool_calls": [{"tool": "playbook", "input": {"name": pb}}],
                    "table": None, "steps": r["steps"], "actions": r["actions"], "ui": r["ui"]}
    item = re.search(r"\b([a-z]{2}-\d{5})\b", q)
    if item:
        d = run_tool(m, "get_item", {"item_no": item.group(1)})
        if "error" not in d:
            s = d["summary"]
            cover = "n/a" if s.get("cover_days") is None else f"{s['cover_days']:,.1f} days"
            ans = (f"**{d['item_no']} – {s.get('description')}** ({C.CLASS_LABELS.get(s.get('item_class'), '')})\n\n"
                   f"- Stock on hand: {s.get('stock_qty') or 0:,.2f} {s.get('uom') or ''} "
                   f"(value {s.get('stock_value') or 0:,.0f} UGX)\n"
                   f"- Daily usage: {s.get('daily_usage') or 0:,.2f}; cover: {cover}\n"
                   f"- Incoming (open PO, excl. stale): {s.get('incoming_qty') or 0:,.2f}; overdue: {s.get('overdue_qty') or 0:,.2f}\n"
                   f"- Cover status: {s.get('cover_status', '-')}; supply: {s.get('supply_status', '-')}; "
                   f"risk: {s.get('risk_level', '-')}\n\nBasis: item position, cover and risk measures.")
            return {"mode": "offline", "answer": ans, "tool_calls": [{"tool": "get_item", "input": item.group(1)}],
                    "table": d["stock_by_location"]}
    view = next((v for pat, v in ROUTES if re.search(pat, q)), "inventory_position")
    cls = "RM" if re.search(r"raw material|\brm\b", q) else "WIP" if re.search(r"\bwip\b|work in progress", q) \
        else "FG" if re.search(r"finished|\bfg\b", q) else None
    site = re.search(r"ship\s?(\d)", q)
    f = {"item_class": cls, "site": f"SHIP {site.group(1)}" if site else None}
    out = V.render(m, view, f)
    kp = "\n".join(f"- {k['label']}: {_fmt(k)}" for k in out["kpis"])
    t = out["tables"][0]
    filt = ", ".join(f"{k}={v}" for k, v in f.items() if v)
    ans = (f"**{out['title']}**{f' ({filt})' if filt else ''} — {out['question']}\n\n{kp}\n\n"
           f"Top rows of *{t['title']}* ({t['total_rows']:,} rows) are shown below.\n\n"
           f"Basis: view {out['no']} {out['title']}. _Offline mode: set ANTHROPIC_API_KEY for full natural-language "
           f"answers._")
    return {"mode": "offline", "answer": ans, "tool_calls": [{"tool": "query_view", "input": {"view": view, **f}}],
            "table": {**t, "rows": t["rows"][:MAX_ROWS]}}


def _fmt(k):
    v = k["value"]
    if v is None:
        return "–"
    if k["kind"] == "money":
        return f"{v:,.0f} UGX"
    if k["kind"] == "days":
        return f"{v:,.1f} days"
    if isinstance(v, (int, float)):
        return f"{v:,.0f}"
    return str(v)
