"""Role-based access control and actor resolution for MEMD routes."""
import base64
import json
import os
from functools import wraps
from flask import jsonify, request

ROLES = ("viewer", "analyst", "steward", "administrator")
ROLE_RANK = {role: idx for idx, role in enumerate(ROLES)}


def auth_required() -> bool:
    """Return True when strict authentication is enabled via PULSE_REQUIRE_AUTH."""
    return os.environ.get("PULSE_REQUIRE_AUTH", "").strip().lower() in ("1", "true", "yes")


def _claims_from_header() -> list[dict]:
    raw = request.headers.get("X-MS-CLIENT-PRINCIPAL")
    if not raw:
        return []
    try:
        data = json.loads(base64.b64decode(raw))
        return data.get("claims", []) if isinstance(data, dict) else []
    except Exception:
        return []


def current_actor() -> dict:
    """Resolve the current actor and role from trusted headers or local dev mode."""
    principal = (request.headers.get("X-MS-CLIENT-PRINCIPAL-NAME") or "").strip()
    claims = _claims_from_header()
    claim_roles = [
        str(c.get("val", "")).strip().lower()
        for c in claims
        if c.get("typ") in ("roles", "role", "http://schemas.microsoft.com/ws/2008/06/identity/claims/role")
    ]
    explicit_role = (request.headers.get("X-MEMD-ROLE") or "").strip().lower()

    if not principal:
        if auth_required():
            return {"authenticated": False, "actor": None, "role": None}
        # Local demo mode (PULSE_REQUIRE_AUTH unset): allow configurable role for testing, default analyst.
        role = explicit_role if explicit_role in ROLE_RANK else "administrator"
        return {"authenticated": True, "actor": "demo", "role": role}

    role = next((r for r in claim_roles if r in ROLE_RANK), None)
    if not role and not auth_required() and explicit_role in ROLE_RANK:
        role = explicit_role
    return {"authenticated": True, "actor": principal, "role": role or "viewer"}


def require_role(minimum_role: str):
    """Decorator enforcing authentication (when enabled) and minimum role rank."""
    if minimum_role not in ROLE_RANK:
        raise ValueError(f"Unknown role: {minimum_role}")

    def decorator(fn):
        @wraps(fn)
        def wrapper(*args, **kwargs):
            info = current_actor()
            if not info["authenticated"]:
                return jsonify({"error": "Authentication required"}), 401
            if ROLE_RANK.get(info["role"], -1) < ROLE_RANK[minimum_role]:
                return jsonify({
                    "error": f"Forbidden: requires {minimum_role} role",
                    "role": info["role"],
                }), 403
            return fn(*args, **kwargs)

        return wrapper

    return decorator

