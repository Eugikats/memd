"""Container Apps Job entry: python -m excel_engine [--job-id ID]"""
from __future__ import annotations

import argparse
import json
import os
import sys


def main(argv=None):
    from excel_engine import keyvault as kv
    kv.hydrate_env_from_vault()

    parser = argparse.ArgumentParser(description="Pulse Excel Intelligence Engine worker")
    parser.add_argument("--job-id", default=os.environ.get("PULSE_EXCEL_JOB_ID"))
    parser.add_argument("--from-servicebus", action="store_true",
                        help="Receive one message from the excel-jobs queue and process it")
    args = parser.parse_args(argv)

    job_id = args.job_id
    if args.from_servicebus or not job_id:
        job_id = job_id or _receive_one()
    if not job_id:
        print("No job id provided", file=sys.stderr)
        return 2

    from excel_engine.pipeline import run_job
    result = run_job(job_id)
    print(json.dumps({"job_id": job_id, "status": result.get("status"), "history_id": result.get("history_id"),
                      "error": result.get("error")}, default=str))
    return 0 if result.get("status") == "ready" else 1


def _receive_one() -> str | None:
    conn = os.environ.get("PULSE_EXCEL_SERVICEBUS") or os.environ.get("AZURE_SERVICEBUS_CONNECTION_STRING")
    queue = os.environ.get("PULSE_EXCEL_SERVICEBUS_QUEUE", "excel-jobs")
    if not conn:
        return None
    from azure.servicebus import ServiceBusClient
    with ServiceBusClient.from_connection_string(conn) as client:
        with client.get_queue_receiver(queue, max_wait_time=30) as receiver:
            for msg in receiver:
                try:
                    body = str(msg)
                    data = json.loads(body)
                    job_id = data.get("job_id")
                    receiver.complete_message(msg)
                    return job_id
                except Exception:
                    receiver.dead_letter_message(msg, reason="bad-payload", error_description="invalid json")
                    return None
    return None


if __name__ == "__main__":
    raise SystemExit(main())
