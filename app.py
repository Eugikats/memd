"""Pulse – Supply Chain POC web app.   Run:  python app.py   then open http://localhost:8050"""
import json
import sys
import threading
from datetime import datetime

from flask import Flask, Response, jsonify, request, send_file, send_from_directory

from pulse import agent as A
from pulse import command, documents, extract, model
from pulse import excel_jobs
from pulse import mail
from pulse import reports as R
from pulse import scheduler
from pulse import views as V
from pulse.bc_client import DATA_DIR, ROOT, load_env

load_env()
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
app = Flask(__name__, static_folder="static")
# Load the model in the background so gunicorn can bind and serve the page while the
# supply-chain cache is opening. A full rebuild from raw files can take minutes after a restart.
STATE = {"model": None, "refreshing": False, "last_error": None, "loading": True, "load_error": None}
REFRESH_LOG = DATA_DIR / "refresh_log.jsonl"
excel_jobs.set_model_getter(lambda: STATE["model"])



def _boot_model():
    try:
        print("pulse: loading supply chain model…", flush=True)
        STATE["model"] = model.load()
        STATE["loading"] = False
        print("pulse: model ready", flush=True)
    except Exception as e:
        STATE["load_error"] = f"{type(e).__name__}: {e}"[:400]
        STATE["loading"] = False
        print(f"pulse: model load failed: {STATE['load_error']}", flush=True)


threading.Thread(target=_boot_model, daemon=True, name="pulse-model").start()


def _model_busy():
    if STATE["model"] is not None:
        return None
    if STATE.get("load_error"):
        return jsonify({"error": STATE["load_error"], "loading": False}), 503
    return jsonify({"error": "The supply chain model is still loading.", "loading": True}), 503


def filters():
    a = request.args
    out = {"item_class": a.get("item_class") or None, "site": a.get("site") or None,
           "location": a.get("location") or None, "q": a.get("q") or None}
    if a.get("sheet"):
        out["sheet"] = a.get("sheet")
    return out


def current_user():
    """Signed-in user from App Service Authentication (Entra ID); 'demo' when running locally."""
    return request.headers.get("X-MS-CLIENT-PRINCIPAL-NAME") or "demo"


def display_name():
    import base64
    raw = request.headers.get("X-MS-CLIENT-PRINCIPAL")
    if raw:
        try:
            claims = json.loads(base64.b64decode(raw)).get("claims", [])
            name = next((c["val"] for c in claims if c.get("typ") == "name"), "")
            if name:
                return name.split()[0]
        except Exception:
            pass
    return __import__("os").environ.get("PULSE_USER_NAME", "")


@app.get("/")
def index():
    return send_from_directory("static", "index.html")


@app.get("/api/meta")
def meta():
    if STATE["model"] is None:
        return jsonify({"loading": True, "load_error": STATE.get("load_error"),
                        "views": [], "notes": {}, "definitions": {}, "sites": [], "locations": [],
                        "refreshing": False, "last_error": STATE.get("last_error"),
                        "playbooks": [], "alert_metrics": {}, "pending_actions": 0,
                        "user_name": display_name(),
                        "user_email": current_user() if current_user() != "demo" else "",
                        "command_mode": command.provider(), "ai_model": ""})
    m = STATE["model"]
    sl = m["stock_loc"]
    return jsonify({"loading": False, "load_error": None,
                    "views": [{"no": n, "id": i, "title": t, "question": q} for n, i, t, q in V.VIEWS],
                    "notes": m["notes"], "definitions": V.DEFINITIONS,
                    "sites": sorted(sl.site.dropna().unique().tolist()),
                    "locations": sorted(sl.location.dropna().unique().tolist()),
                    "refreshing": STATE["refreshing"], "last_error": STATE["last_error"],
                    "playbooks": A.PLAYBOOKS, "alert_metrics": A.ALERT_METRICS,
                    "pending_actions": sum(a["status"] == "proposed" for a in A.list_actions()),
                    "user_name": display_name(), "user_email": current_user() if current_user() != "demo" else "",
                    "command_mode": command.provider(),
                    "ai_model": __import__("os").environ.get("AZURE_OPENAI_DEPLOYMENT", "") if command.provider() == "azure" else "",
                    "excel_engine": {
                        "blob": __import__("pulse.excel_blob", fromlist=["x"]).configured(),
                        "servicebus": excel_jobs.servicebus_configured(),
                        "search": bool(__import__("os").environ.get("PULSE_EXCEL_SEARCH_ENDPOINT")),
                        "keyvault": bool(__import__("os").environ.get("PULSE_KEY_VAULT_URI") or __import__("os").environ.get("AZURE_KEY_VAULT_URI")),
                    }})


@app.before_request
def _gate_until_model():
    path = request.path or ""
    if path == "/" or path.startswith("/static/") or path in ("/api/meta", "/api/mail/status"):
        return None
    # Upload + Excel job status do not need the model in-process (rebuild worker waits for it).
    if path == "/api/documents" or path.startswith("/api/excel/"):
        return None
    if path.startswith("/api/"):
        return _model_busy()
    return None


@app.get("/api/cockpit")
def cockpit():
    """Command Center totals and the measured comparisons that rotate under them."""
    feed = command.cockpit_feed(STATE["model"])
    return _json({"key": feed["key"], "metrics": feed["metrics"], "frames": feed["frames"]})


@app.get("/api/cockpit/narrate")
def cockpit_narrate():
    """Extra comparison lines written from the same facts. Empty when no model is configured."""
    return _json({"frames": command.cockpit_narration(STATE["model"])})


@app.get("/api/insights")
def insights():
    """Headline insight cards for the Command Center home, from the same governed model."""
    m = STATE["model"]
    it, ex, pl, rd = m["item"], m["exceptions"], m["purchase_lines"], m["readiness"]
    soon = it[it.stockout_exposed & (it.cover_status != "Stockout") & (it.stockout_day <= 7)]
    top = soon.sort_values("daily_usage_value", ascending=False).head(1)
    od = pl[pl.overdue & ~pl.stale]
    cons = rd[rd.readiness == "Constrained"]
    high = ex[ex.severity == "High"]
    cards = [
        {"tone": "critical", "label": "Stockout risk", "value": f"{len(soon):,} items",
         "text": "run out within 7 days before replenishment arrives"
                 + (f" — highest impact: {top.item_no.iloc[0]} {top.description.iloc[0]}" if len(top) else ""),
         "question": "Which items will run out in the next 7 days, and is any supply on the way?"},
        {"tone": "serious", "label": "Overdue receipts", "value": f"{len(od):,} PO lines",
         "text": f"overdue (≤90 days) worth {od.outstanding_value.sum() / 1e6:,.0f} M UGX; "
                 f"{int(pl.stale.sum())} more are stale",
         "question": "Which purchase orders are overdue, from which vendors, and for which items?"},
        {"tone": "warning", "label": "Production constraints", "value": f"{len(cons):,} products",
         "text": f"have under 7 days of components; {len(high):,} high-severity exceptions in total",
         "question": "Which products can't we produce this week and which components constrain them?"},
    ]
    return jsonify(cards)


@app.get("/api/view/<view_id>")
def view(view_id):
    if view_id not in V.VIEW_IDS and view_id != "overview":
        return jsonify({"error": "unknown view"}), 404
    return app.response_class(json.dumps(V.render(STATE["model"], view_id, filters()), default=str),
                              mimetype="application/json")


@app.get("/api/item/<item_no>")
def item(item_no):
    d = V.item_detail(STATE["model"], item_no.upper())
    if not d:
        return jsonify({"error": "not found"}), 404
    return app.response_class(json.dumps(d, default=str), mimetype="application/json")


@app.post("/api/ask")
def ask():
    body = request.get_json(force=True)
    res = command.ask(STATE["model"], body.get("question", "")[:2000], user=current_user(),
                      history=body.get("history"), documents=body.get("documents") or [],
                      emails=body.get("emails") or [], context=(body.get("context") or "")[:6000])
    return app.response_class(json.dumps(res, default=str), mimetype="application/json")


@app.post("/api/playbook/<name>")
def playbook(name):
    res = command.run_playbook(STATE["model"], name, user=current_user(),
                               args=request.get_json(silent=True) or {})
    if res is None:
        return jsonify({"error": "unknown playbook"}), 404
    return app.response_class(json.dumps(res, default=str), mimetype="application/json")


@app.get("/api/actions")
def actions():
    return app.response_class(json.dumps(A.list_actions(), default=str), mimetype="application/json")


@app.post("/api/actions/<aid>")
def action_status(aid):
    status = (request.get_json(force=True) or {}).get("status")
    if status not in ("approved", "dismissed", "proposed"):
        return jsonify({"error": "bad status"}), 400
    a = A.set_action_status(aid, status, user=current_user())
    return jsonify(a) if a else (jsonify({"error": "not found"}), 404)


@app.get("/api/actions/<aid>/download")
def action_download(aid):
    a = A.get_action(aid)
    if not a:
        return jsonify({"error": "not found"}), 404
    p = a["payload"]
    if a["kind"] == "transfer":
        cols = ["item_no", "description", "item_class", "from_site", "to_site", "qty", "uom"]
        body = ",".join(cols) + "\n" + "\n".join(",".join(f'"{r.get(c, "")}"' for c in cols) for r in p["rows"])
        return app.response_class(body, mimetype="text/csv",
                                  headers={"Content-Disposition": f"attachment; filename=transfer_plan_{aid}.csv"})
    if a["kind"] == "report":
        return report_history_download(p["history_id"])
    if a["kind"] == "briefing":
        return app.response_class(p["markdown"], mimetype="text/markdown",
                                  headers={"Content-Disposition": f"attachment; filename=briefing_{aid}.md"})
    if a["kind"] == "expedite":
        body = f"Subject: {p['subject']}\n\n{p['body']}"
        return app.response_class(body, mimetype="text/plain",
                                  headers={"Content-Disposition": f"attachment; filename=expedite_{aid}.txt"})
    return app.response_class(json.dumps(p, default=str), mimetype="application/json")


@app.get("/api/alerts")
def alerts():
    return app.response_class(json.dumps(A.list_alerts(STATE["model"]), default=str), mimetype="application/json")


@app.delete("/api/alerts/<rid>")
def alert_delete(rid):
    A.delete_alert(rid)
    return jsonify({"ok": True})


@app.get("/api/watchlist")
def watch():
    return app.response_class(json.dumps(A.watchlist(STATE["model"]), default=str), mimetype="application/json")


@app.post("/api/watchlist/<item_no>")
def watch_add(item_no):
    return _json_resp(A.watch_items(STATE["model"], [item_no]))


def _json_resp(o):
    return app.response_class(json.dumps(o, default=str), mimetype="application/json")


@app.delete("/api/watchlist/<item_no>")
def unwatch(item_no):
    A.unwatch(item_no)
    return jsonify({"ok": True})


def _refresh(days, trigger="manual", user=None):
    t = datetime.now()
    rec = {"started": R.now().isoformat(timespec="seconds"), "trigger": trigger, "user": user}
    try:
        log = extract.run(days)
        STATE["model"] = model.build()
        fired = [r["name"] for r in A.evaluate_alerts(STATE["model"]) if r.get("triggered")]
        r = STATE["model"]["notes"].get("reconciliation") or {}
        rec.update(status="ok", tables=log["tables"], alerts_triggered=fired,
                   records=sum(v["rows"] for v in log["tables"].values()),
                   recon_pct=round(100 * r["items_matching"] / max(r["items_compared"], 1), 2) if r else None)
        STATE["last_error"] = None
    except Exception as e:  # logged and surfaced in the UI
        rec.update(status="failed", error=str(e)[:500])
        STATE["last_error"] = str(e)[:500]
    finally:
        rec["seconds"] = round((datetime.now() - t).total_seconds(), 1)
        with REFRESH_LOG.open("a", encoding="utf-8") as f:
            f.write(json.dumps(rec) + "\n")
        STATE["refreshing"] = False
        if rec.get("status") == "ok":
            from pulse import mail
            mail.notice("refresh", "Data refresh finished",
                        "<p>The Business Central load finished. Manufacturing Intelligence Platform is using the new data.</p>")
        elif rec.get("status") == "failed":
            from pulse import mail
            mail.notice("refresh", "Data refresh failed",
                        f"<p>The Business Central load failed. {rec.get('error', '')}</p>")


@app.post("/api/refresh")
def refresh():
    if STATE["refreshing"]:
        return jsonify({"status": "already running"})
    STATE["refreshing"] = True
    threading.Thread(target=_refresh, args=(int(request.args.get("days", 90)), "manual", current_user()),
                     daemon=True).start()
    return jsonify({"status": "started"})


# ---------------------------------------------------------------- reporting
def _json(obj, status=200):
    return app.response_class(json.dumps(obj, default=str), status=status, mimetype="application/json")


def _params():
    a = request.args
    return {k: a.get(k) for k in ("item_class", "site", "q") if a.get(k)}


@app.get("/api/reports")
def reports_catalogue():
    return _json({"catalogue": R.catalogue(), "datasets": R.dataset_meta(STATE["model"]), "aggs": R.AGGS})


@app.get("/api/reports/run/<path:rid>")
def report_run(rid):
    try:
        rep_ = R.run_report(STATE["model"], rid, _params())
    except KeyError as e:
        return _json({"error": str(e)}, 404)
    for s_ in rep_["sections"]:          # preview: first 1,000 rows; exports carry everything
        s_["rows"] = s_["rows"][:1000]
    return _json(rep_)


def _send_report(rep_, fmt):
    fname = f"{rep_['name'].replace(' ', '_').replace('&', 'and')}_{rep_['reporting_date']}"
    if fmt == "csv":
        return Response(R.to_csv(rep_), mimetype="text/csv",
                        headers={"Content-Disposition": f"attachment; filename={fname}.csv"})
    return Response(R.to_xlsx(rep_), mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    headers={"Content-Disposition": f"attachment; filename={fname}.xlsx"})


@app.get("/api/reports/export/<path:rid>")
def report_export(rid):
    try:
        rep_ = R.run_report(STATE["model"], rid, _params())
    except KeyError as e:
        return _json({"error": str(e)}, 404)
    command._audit({"at": datetime.now().isoformat(timespec="seconds"), "user": current_user(),
                    "question": f"[export] {rep_['name']} ({request.args.get('fmt', 'xlsx')})", "mode": "report",
                    "answer": "", "actions": []})
    return _send_report(rep_, request.args.get("fmt", "xlsx"))


@app.post("/api/reports/custom/run")
def custom_run():
    body = request.get_json(force=True)
    try:
        out = R.run_custom(STATE["model"], body.get("spec") or {})
    except (KeyError, ValueError) as e:
        return _json({"error": str(e)}, 400)
    return _json(out)


@app.post("/api/reports/custom/export")
def custom_export():
    body = request.get_json(force=True)
    spec = body.get("spec") or {}
    out = R.run_custom(STATE["model"], spec)
    n = STATE["model"]["notes"]
    rep_ = {"name": body.get("name") or "Custom report", "description": "", "params": spec.get("filters") or {},
            "company": n["company"], "environment": n["environment"], "reporting_date": n["reporting_date"],
            "data_loaded": n["extracted_at"], "generated_at": R.now().strftime("%Y-%m-%d %H:%M"), **out}
    return _send_report(rep_, body.get("fmt", "xlsx"))


@app.post("/api/reports/custom")
def custom_save():
    body = request.get_json(force=True)
    if not body.get("name") or not body.get("spec"):
        return _json({"error": "name and spec required"}, 400)
    return _json(R.save_custom(body["name"], body["spec"], body.get("description", ""), current_user(), body.get("id")))


@app.delete("/api/reports/custom/<path:rid>")
def custom_delete(rid):
    R.delete_custom(rid)
    return _json({"ok": True})


@app.post("/api/reports/generate/<path:rid>")
def report_generate(rid):
    try:
        e = R.generate(STATE["model"], rid, request.get_json(silent=True) or {}, current_user(), "manual")
    except KeyError as ex:
        return _json({"error": str(ex)}, 404)
    return _json(e)


@app.get("/api/reports/history")
def report_history():
    return _json(R.history())


@app.get("/api/reports/history/<hid>/download")
def report_history_download(hid):
    e, p = R.history_file(hid)
    if not e:
        return _json({"error": "not found"}, 404)
    ext = p.suffix or ".xlsx"
    return send_file(p, as_attachment=True,
                     download_name=f"{e['name'].replace(' ', '_').replace('&', 'and')}_{e['generated_at'][:10]}{ext}")


@app.get("/api/mail/status")
def mail_status():
    return _json({"configured": mail.configured(), "prefs": mail.load_prefs()})


@app.post("/api/notifications/mail")
def notification_mail():
    b = request.get_json(force=True) or {}
    return _json(mail.save_prefs({"email": b.get("email"), "refresh": b.get("refresh"), "reports": b.get("reports")}))


@app.post("/api/documents")
def document_upload():
    f = request.files.get("file")
    if not f:
        return _json({"error": "Choose a file."}, 400)
    try:
        saved = documents.save(f.filename or "upload.txt", f.read())
    except ValueError as e:
        return _json({"error": str(e)}, 400)
    except Exception as e:
        detail = str(e).strip().splitlines()[0][:120]
        return _json({"error": f"Could not read that file ({type(e).__name__}{': ' + detail if detail else ''})."}, 400)
    return _json(saved)


@app.post("/api/excel/jobs")
def excel_job_create():
    body = request.get_json(force=True) or {}
    doc_id = body.get("document_id")
    if not doc_id or not documents.get(doc_id):
        return _json({"error": "Attach a workbook first."}, 400)
    job = excel_jobs.create(doc_id, body.get("instruction") or body.get("question") or "rebuild from upload",
                            user=current_user(), view=body.get("view"))
    return _json(job)


@app.get("/api/excel/jobs")
def excel_jobs_list():
    return _json({"jobs": excel_jobs.list_recent(40)})


@app.get("/api/excel/jobs/<job_id>")
def excel_job_get(job_id):
    job = excel_jobs.get(job_id, redrive=True)
    if not job:
        return _json({"error": "Unknown Excel job"}, 404)
    out = dict(job)
    if out.get("history_id"):
        out["download"] = f"/api/reports/history/{out['history_id']}/download"
    return _json(out)


@app.post("/api/excel/jobs/<job_id>/deliver")
def excel_job_deliver(job_id):
    """Email or webhook-deliver a ready job (Logic Apps URL used when no recipients)."""
    job = excel_jobs.get(job_id)
    if not job:
        return _json({"error": "Unknown Excel job"}, 404)
    if job.get("status") != "ready":
        return _json({"error": "Job is not ready", "status": job.get("status")}, 400)
    body = request.get_json(silent=True) or {}
    recipients = body.get("recipients") or []
    if recipients:
        from excel_engine import delivery
        return _json(delivery.email_history(job["history_id"], recipients))
    from excel_engine import delivery
    return _json(delivery.notify_ready(job))


@app.get("/api/schedules")
def schedules_get():
    s = R.schedules()
    names = {c["id"]: c["name"] for c in R.catalogue()}
    for x in s["reports"]:
        x["name"] = names.get(x["report_id"], x["report_id"])
    return _json(s)


@app.post("/api/schedules/report")
def schedule_add():
    b = request.get_json(force=True)
    if b.get("frequency") not in ("daily", "weekdays", "weekly", "monthly", "custom"):
        return _json({"error": "bad frequency"}, 400)
    if b.get("frequency") == "custom" and b.get("unit") not in ("hours", "days"):
        return _json({"error": "custom schedules need a period of hours or days"}, 400)
    sch = R.add_schedule(b["report_id"], b.get("params"), b["frequency"], b.get("time", "07:00"),
                         b.get("weekday", 0), b.get("monthday", 1), current_user(),
                         recipients=b.get("recipients"), every=b.get("every"), unit=b.get("unit"))
    from pulse import mail
    sch["mail_configured"] = mail.configured()
    return _json(sch)


@app.post("/api/schedules/report/<sid>")
def schedule_update(sid):
    b = request.get_json(force=True)
    s = R.schedules()
    for x in s["reports"]:
        if x["id"] == sid:
            if "enabled" in b:
                x["enabled"] = bool(b["enabled"])
            if x["enabled"]:
                x["next_run"] = R.next_run(x).isoformat(timespec="minutes")
    R.save_schedules(s)
    return _json({"ok": True})


@app.delete("/api/schedules/report/<sid>")
def schedule_delete(sid):
    s = R.schedules()
    s["reports"] = [x for x in s["reports"] if x["id"] != sid]
    R.save_schedules(s)
    return _json({"ok": True})


@app.post("/api/schedules/refresh")
def schedule_refresh():
    b = request.get_json(force=True)
    s = R.schedules()
    s["refresh"].update({"enabled": bool(b.get("enabled")), "time": b.get("time", "06:00"),
                         "days": [int(d) for d in b.get("days", [])]})
    R.save_schedules(s)
    return _json(s["refresh"])


@app.get("/api/refresh/log")
def refresh_log():
    rows = []
    if REFRESH_LOG.exists():
        rows = [json.loads(x) for x in REFRESH_LOG.read_text(encoding="utf-8").splitlines() if x.strip()]
    return _json(list(reversed(rows))[:200])


@app.get("/api/trends")
def trends():
    p = DATA_DIR / "kpi_history.jsonl"
    hist = [json.loads(x) for x in p.read_text(encoding="utf-8").splitlines() if x.strip()] if p.exists() else []
    w = STATE["model"]["weekly"]
    return _json({"kpi": hist, "weekly": w.assign(week=w.week.astype(str)).to_dict("records")})


@app.get("/api/audit")
def audit():
    rows = []
    p = DATA_DIR / "audit.jsonl"
    if p.exists():
        for x in p.read_text(encoding="utf-8").splitlines()[-1000:]:
            try:
                r = json.loads(x)
                rows.append({"at": r.get("at"), "user": r.get("user"), "type": r.get("mode"),
                             "event": r.get("question"),
                             "detail": __import__("re").sub(r"[*#_`|]+", "", r.get("answer") or "").strip()[:240]})
            except json.JSONDecodeError:
                pass
    if REFRESH_LOG.exists():
        for x in REFRESH_LOG.read_text(encoding="utf-8").splitlines():
            r = json.loads(x)
            rows.append({"at": r.get("started"), "user": r.get("user") or r.get("trigger"), "type": "refresh",
                         "event": f"Data refresh ({r.get('trigger', 'manual')}) – {r.get('status')}",
                         "detail": r.get("error") or f"{r.get('records', '')} records, {r.get('seconds')} s, "
                                                     f"reconciliation {r.get('recon_pct', '–')}%"})
    for a in A.list_actions():
        if a.get("decided"):
            rows.append({"at": a["decided"], "user": a.get("decided_by"), "type": "decision",
                         "event": f"{a['status'].title()}: {a['title']}", "detail": a.get("summary", "")})
    for h in R.history():
        rows.append({"at": h["generated_at"], "user": h["by"], "type": "report",
                     "event": f"Report generated ({h['trigger']}): {h['name']}", "detail": f"{h['rows']:,} rows"})
    for e in excel_jobs.audit_rows(200):
        rows.append({"at": e.get("at"), "user": e.get("user") or "system", "type": "excel_engine",
                     "event": f"{e.get('type')}: job {e.get('job_id')}",
                     "detail": e.get("error") or f"history={e.get('history_id') or '–'} tenant={e.get('tenant') or '–'}"})
    rows.sort(key=lambda r: str(r["at"] or ""), reverse=True)
    return _json(rows[:1000])


scheduler.start(STATE, _refresh)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=int(__import__("os").environ.get("PORT", 8050)), debug=False)
