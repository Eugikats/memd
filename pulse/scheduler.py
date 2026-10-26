"""Background scheduler (one thread per app process): scheduled Business Central refresh and scheduled reports.

Times are East Africa Time (Africa/Kampala). App Service runs one gunicorn worker with Always On, so one
thread owns the schedule.
"""
import threading
import time
import traceback
from datetime import datetime

from . import mail
from . import reports as R

_started = False


def _due_refresh(cfg, now):
    if not cfg.get("enabled"):
        return False
    hh, mm = (int(x) for x in cfg.get("time", "06:00").split(":"))
    if now.weekday() not in cfg.get("days", list(range(7))):
        return False
    if (now.hour, now.minute) < (hh, mm):
        return False
    return (cfg.get("last_run") or "")[:10] != now.date().isoformat()


def tick(state, refresh_fn):
    now = R.now()
    s = R.schedules()
    if _due_refresh(s["refresh"], now) and not state.get("refreshing"):
        s["refresh"]["last_run"] = now.isoformat(timespec="minutes")
        R.save_schedules(s)
        state["refreshing"] = True
        refresh_fn(90, "schedule", "scheduler")   # runs in this thread; logs its own result
    s = R.schedules()
    changed = False
    for sch in s["reports"]:
        if not sch.get("enabled") or not sch.get("next_run"):
            continue
        if datetime.fromisoformat(sch["next_run"]) <= now:
            try:
                e = R.generate(state["model"], sch["report_id"], sch.get("params"), user=sch.get("created_by"),
                               trigger="schedule")
                sch["last_status"], sch["last_file"] = "ok", e["id"]
                try:
                    _entry, path = R.history_file(e["id"])
                    attachment = None
                    if path:
                        attachment = (f"{e['name'].replace(' ', '_')}.xlsx", path.read_bytes(),
                                      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    sch["last_email"] = mail.send(
                        sch.get("recipients"), f"Scheduled report: {e['name']}",
                        f"<p>The scheduled report <b>{e['name']}</b> is attached. It was built from the Manufacturing Intelligence Platform data.</p>",
                        attachment)
                    mail.notice("reports", f"Report ready: {e['name']}",
                                f"<p>{e['name']} ran on schedule. {e.get('rows', 0)} rows were saved to report history.</p>")
                except Exception as ex:
                    sch["last_email"] = str(ex)[:160]
            except Exception as ex:  # keep the schedule alive; surface the error in the UI
                sch["last_status"] = f"failed: {ex}"[:200]
                sch["last_email"] = "not sent"
            sch["last_run"] = now.isoformat(timespec="minutes")
            nr = R.next_run(sch, now)
            sch["next_run"] = nr.isoformat(timespec="minutes") if nr else None
            changed = True
    if changed:
        R.save_schedules(s)


def start(state, refresh_fn):
    global _started
    if _started:
        return
    _started = True

    def loop():
        while True:
            try:
                tick(state, refresh_fn)
            except Exception:
                traceback.print_exc()
            time.sleep(30)

    threading.Thread(target=loop, daemon=True, name="pulse-scheduler").start()
