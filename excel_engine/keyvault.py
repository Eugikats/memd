"""Optional Key Vault secret resolution for the Excel engine."""
from __future__ import annotations

import os


def configured() -> bool:
    return bool(os.environ.get("PULSE_KEY_VAULT_URI") or os.environ.get("AZURE_KEY_VAULT_URI"))


def get_secret(name: str, default: str | None = None) -> str | None:
    """Return a secret value, or env fallback `PULSE_SECRET_<NAME>`, or default."""
    env_key = "PULSE_SECRET_" + name.upper().replace("-", "_")
    if os.environ.get(env_key):
        return os.environ[env_key]
    uri = os.environ.get("PULSE_KEY_VAULT_URI") or os.environ.get("AZURE_KEY_VAULT_URI")
    if not uri:
        return default
    try:
        from azure.identity import DefaultAzureCredential
        from azure.keyvault.secrets import SecretClient
        client = SecretClient(vault_url=uri, credential=DefaultAzureCredential())
        return client.get_secret(name).value
    except Exception:
        return default


def hydrate_env_from_vault(mapping: dict[str, str] | None = None):
    """
    mapping: {env_var: secret_name}. Sets os.environ if missing.
    Default mapping covers storage, service bus, and search keys.
    """
    mapping = mapping or {
        "PULSE_EXCEL_STORAGE_CONNECTION": "pulse-excel-storage-connection",
        "PULSE_EXCEL_SERVICEBUS": "pulse-excel-servicebus",
        "PULSE_EXCEL_SEARCH_KEY": "pulse-excel-search-key",
    }
    if not configured():
        return
    for env_var, secret_name in mapping.items():
        if os.environ.get(env_var):
            continue
        value = get_secret(secret_name)
        if value:
            os.environ[env_var] = value
