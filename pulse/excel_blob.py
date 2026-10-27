"""Workbook object storage: Azure Blob when configured, otherwise local disk under DATA_DIR."""
from __future__ import annotations

import os
import re
from pathlib import Path

from .bc_client import DATA_DIR

_CONTAINER = os.environ.get("PULSE_EXCEL_BLOB_CONTAINER", "excel-workbooks")
_ACCOUNT = os.environ.get("PULSE_EXCEL_STORAGE_ACCOUNT") or os.environ.get("AZURE_STORAGE_ACCOUNT")
_CONN = os.environ.get("PULSE_EXCEL_STORAGE_CONNECTION") or os.environ.get("AZURE_STORAGE_CONNECTION_STRING")
_LOCAL = Path(os.environ.get("PULSE_EXCEL_LOCAL_DIR") or (DATA_DIR / "excel" / "blobs"))


def configured() -> bool:
    return bool(_CONN or _ACCOUNT)


def tenant_prefix(tenant: str | None = None) -> str:
    raw = (tenant or os.environ.get("PULSE_TENANT_ID") or "default").strip().lower()
    safe = re.sub(r"[^a-z0-9._-]+", "-", raw)[:64] or "default"
    return safe


def _local_path(key: str) -> Path:
    path = _LOCAL / key
    path.parent.mkdir(parents=True, exist_ok=True)
    return path


def _client():
    if not configured():
        return None
    try:
        from azure.storage.blob import BlobServiceClient
    except ImportError:
        return None
    if _CONN:
        return BlobServiceClient.from_connection_string(_CONN)
    from azure.identity import DefaultAzureCredential
    url = f"https://{_ACCOUNT}.blob.core.windows.net"
    return BlobServiceClient(url, credential=DefaultAzureCredential())


def put_bytes(key: str, data: bytes, content_type: str = "application/octet-stream") -> dict:
    """Store bytes; returns {uri, backend}."""
    key = key.lstrip("/")
    client = _client()
    if client is not None:
        try:
            cc = client.get_container_client(_CONTAINER)
            try:
                cc.create_container()
            except Exception:
                pass
            cc.upload_blob(key, data, overwrite=True, content_type=content_type)
            return {"uri": f"azure://{_CONTAINER}/{key}", "backend": "azure", "key": key}
        except Exception:
            pass
    path = _local_path(key)
    path.write_bytes(data)
    return {"uri": f"file://{path}", "backend": "local", "key": key}


def get_bytes(key_or_uri: str) -> bytes:
    key = key_or_uri
    if key.startswith("azure://"):
        key = key.split("/", 3)[-1] if key.count("/") >= 3 else key.replace("azure://", "")
        if key.startswith(_CONTAINER + "/"):
            key = key[len(_CONTAINER) + 1:]
    elif key.startswith("file://"):
        return Path(key[7:]).read_bytes()
    client = _client()
    if client is not None and not str(key_or_uri).startswith("file://"):
        try:
            return client.get_container_client(_CONTAINER).download_blob(key).readall()
        except Exception:
            pass
    path = _local_path(key)
    if not path.exists():
        raise FileNotFoundError(f"Workbook blob not found: {key}")
    return path.read_bytes()


def uri_for(key: str) -> str:
    if configured() and _client() is not None:
        return f"azure://{_CONTAINER}/{key.lstrip('/')}"
    return f"file://{_local_path(key.lstrip('/'))}"
