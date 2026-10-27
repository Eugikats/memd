"""Index extracted workbook intelligence in Azure AI Search (no-op when not configured)."""
from __future__ import annotations

import hashlib
import os


def configured() -> bool:
    return bool(os.environ.get("PULSE_EXCEL_SEARCH_ENDPOINT") and
                (os.environ.get("PULSE_EXCEL_SEARCH_KEY") or os.environ.get("AZURE_CLIENT_ID")))


def index_name() -> str:
    return os.environ.get("PULSE_EXCEL_SEARCH_INDEX", "workbook-intelligence")


def _client():
    endpoint = os.environ.get("PULSE_EXCEL_SEARCH_ENDPOINT")
    if not endpoint:
        return None
    try:
        from azure.search.documents import SearchClient
        from azure.core.credentials import AzureKeyCredential
    except ImportError:
        return None
    key = os.environ.get("PULSE_EXCEL_SEARCH_KEY")
    if key:
        return SearchClient(endpoint=endpoint, index_name=index_name(), credential=AzureKeyCredential(key))
    try:
        from azure.identity import DefaultAzureCredential
        return SearchClient(endpoint=endpoint, index_name=index_name(),
                            credential=DefaultAzureCredential())
    except Exception:
        return None


def ensure_index():
    """Best-effort create a simple index if admin key is available."""
    endpoint = os.environ.get("PULSE_EXCEL_SEARCH_ENDPOINT")
    key = os.environ.get("PULSE_EXCEL_SEARCH_KEY")
    if not (endpoint and key):
        return False
    try:
        from azure.search.documents.indexes import SearchIndexClient
        from azure.search.documents.indexes.models import SearchIndex, SimpleField, SearchableField, SearchFieldDataType
        from azure.core.credentials import AzureKeyCredential
        ic = SearchIndexClient(endpoint, AzureKeyCredential(key))
        fields = [
            SimpleField(name="id", type=SearchFieldDataType.String, key=True),
            SimpleField(name="job_id", type=SearchFieldDataType.String, filterable=True),
            SimpleField(name="tenant", type=SearchFieldDataType.String, filterable=True),
            SimpleField(name="document_id", type=SearchFieldDataType.String, filterable=True),
            SearchableField(name="workbook_name", type=SearchFieldDataType.String),
            SearchableField(name="sheet", type=SearchFieldDataType.String),
            SearchableField(name="content", type=SearchFieldDataType.String),
            SimpleField(name="kind", type=SearchFieldDataType.String, filterable=True),
        ]
        try:
            ic.create_index(SearchIndex(name=index_name(), fields=fields))
        except Exception:
            pass
        return True
    except Exception:
        return False


def index_extract(job_id: str, document_id: str, tenant: str, extract: dict) -> dict:
    """Push sheet/formula chunks to AI Search. Returns {indexed, backend}."""
    docs = []
    name = extract.get("name") or "workbook"
    for sheet in extract.get("sheets") or []:
        content = f"Sheet {sheet.get('name')}. Headers: {', '.join(sheet.get('headers') or [])}. " \
                  f"Formulas: {sheet.get('formula_count', 0)}."
        for sample in (sheet.get("formula_samples") or [])[:10]:
            content += f" {sample.get('cell')}={sample.get('formula')}"
        rid = hashlib.sha1(f"{job_id}:{sheet.get('name')}:sheet".encode()).hexdigest()
        docs.append({"id": rid, "job_id": job_id, "tenant": tenant, "document_id": document_id,
                     "workbook_name": name, "sheet": sheet.get("name"), "content": content, "kind": "sheet"})
    for i, formula in enumerate((extract.get("formulas") or [])[:200]):
        rid = hashlib.sha1(f"{job_id}:f:{i}:{formula.get('cell')}".encode()).hexdigest()
        docs.append({"id": rid, "job_id": job_id, "tenant": tenant, "document_id": document_id,
                     "workbook_name": name, "sheet": formula.get("sheet"),
                     "content": f"{formula.get('sheet')}!{formula.get('cell')} {formula.get('formula')}",
                     "kind": "formula"})
    summary_id = hashlib.sha1(f"{job_id}:summary".encode()).hexdigest()
    docs.append({"id": summary_id, "job_id": job_id, "tenant": tenant, "document_id": document_id,
                 "workbook_name": name, "sheet": "", "content": extract.get("summary") or "", "kind": "summary"})

    client = _client()
    if client is None:
        return {"indexed": 0, "backend": "none", "docs": len(docs)}
    ensure_index()
    try:
        client.upload_documents(docs)
        return {"indexed": len(docs), "backend": "azure_ai_search", "docs": len(docs)}
    except Exception as e:
        return {"indexed": 0, "backend": "error", "error": f"{type(e).__name__}: {e}"[:200], "docs": len(docs)}


def search(query: str, tenant: str | None = None, top: int = 8) -> list:
    client = _client()
    if client is None or not query:
        return []
    filt = f"tenant eq '{tenant}'" if tenant else None
    try:
        results = client.search(query, filter=filt, top=top)
        return [{"sheet": r.get("sheet"), "content": r.get("content"), "kind": r.get("kind"),
                 "workbook_name": r.get("workbook_name"), "job_id": r.get("job_id")} for r in results]
    except Exception:
        return []
