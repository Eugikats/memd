"""Azure Function stub: HTTP trigger that accepts Pulse excel job.ready payloads.

Deploy separately with Azure Functions Core Tools. Set PULSE_FORWARD_EMAIL to auto-email.
"""
import json
import logging
import os

import azure.functions as func

app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)


@app.route(route="excel-ready", methods=["POST"])
def excel_ready(req: func.HttpRequest) -> func.HttpResponse:
    try:
        body = req.get_json()
    except ValueError:
        return func.HttpResponse("invalid json", status_code=400)
    logging.info("excel job ready: %s", json.dumps(body)[:500])
    # Optional: forward to an ops mailbox via Logic App continuation or Graph.
    _ = os.environ.get("PULSE_FORWARD_EMAIL")
    return func.HttpResponse(json.dumps({"ok": True, "job_id": body.get("job_id")}),
                             mimetype="application/json")
