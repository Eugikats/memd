"""Optional Microsoft Graph Excel recalculation for supported cloud copies."""
from __future__ import annotations

import os


def configured() -> bool:
    return bool(os.environ.get("PULSE_GRAPH_EXCEL_DRIVE_ID") and
                (os.environ.get("PULSE_MAIL_TENANT_ID") or os.environ.get("AZURE_TENANT_ID")))


def recalculate(output_bytes: bytes, filename: str = "report.xlsx") -> dict:
    """
    Best-effort: upload to a configured drive item and request workbook calculation.
    If Graph Excel is not configured or fails, return recalculated=False with a flag.
    """
    if not configured():
        return {"recalculated": False, "flag": "not recalculated in cloud (Graph Excel not configured)",
                "bytes": output_bytes}
    try:
        import requests
        from azure.identity import DefaultAzureCredential, ClientSecretCredential
    except ImportError:
        return {"recalculated": False, "flag": "not recalculated in cloud (deps missing)", "bytes": output_bytes}

    tenant = os.environ.get("PULSE_MAIL_TENANT_ID") or os.environ.get("AZURE_TENANT_ID")
    client_id = os.environ.get("PULSE_MAIL_CLIENT_ID") or os.environ.get("AZURE_CLIENT_ID")
    secret = os.environ.get("PULSE_MAIL_CLIENT_SECRET")
    drive_id = os.environ["PULSE_GRAPH_EXCEL_DRIVE_ID"]
    folder = os.environ.get("PULSE_GRAPH_EXCEL_FOLDER", "pulse-excel")

    try:
        if secret and client_id and tenant:
            cred = ClientSecretCredential(tenant, client_id, secret)
        else:
            cred = DefaultAzureCredential()
        token = cred.get_token("https://graph.microsoft.com/.default").token
        headers = {"Authorization": f"Bearer {token}"}
        # Upload
        put = requests.put(
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{folder}/{filename}:/content",
            headers={**headers, "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
            data=output_bytes, timeout=120)
        if put.status_code >= 300:
            return {"recalculated": False,
                    "flag": f"not recalculated in cloud (upload HTTP {put.status_code})",
                    "bytes": output_bytes}
        item = put.json()
        item_id = item.get("id")
        # Request calculation
        calc = requests.post(
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/workbook/application/calculate",
            headers={**headers, "Content-Type": "application/json"},
            json={"calculationType": "Full"}, timeout=60)
        if calc.status_code >= 300:
            return {"recalculated": False,
                    "flag": f"uploaded but calculate HTTP {calc.status_code}",
                    "bytes": output_bytes, "drive_item_id": item_id}
        # Download calculated file (values may still be formula text without session)
        get = requests.get(
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content",
            headers=headers, timeout=120)
        data = get.content if get.status_code < 300 else output_bytes
        return {"recalculated": True, "flag": None, "bytes": data, "drive_item_id": item_id}
    except Exception as e:
        return {"recalculated": False,
                "flag": f"not recalculated in cloud ({type(e).__name__}: {str(e)[:120]})",
                "bytes": output_bytes}
