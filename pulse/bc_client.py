"""Read-only Business Central client: Entra ID client-credentials auth, OData paging, retry."""
import os
import time
from pathlib import Path
from urllib.parse import quote

import requests

ROOT = Path(__file__).resolve().parent.parent
# Persistent data folder. On Azure App Service set PULSE_DATA_DIR=/home/data so data survives redeploys.
DATA_DIR = Path(os.environ.get("PULSE_DATA_DIR") or ROOT / "data")


def load_env(path=ROOT / ".env"):
    if path.exists():
        for line in path.read_text().splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip().strip("'\""))


class BCClient:
    def __init__(self):
        load_env()
        e = os.environ
        self.tenant = e["BC_TENANT_ID"]
        self.env = e["BC_ENVIRONMENT"]
        self.company_id = e["BC_COMPANY_ID"]
        self.client_id = e["BC_CLIENT_ID"]
        self.secret = e["BC_CLIENT_SECRET"]
        self.scope = e.get("BC_SCOPE", "https://api.businesscentral.dynamics.com/.default")
        self.base = f"https://api.businesscentral.dynamics.com/v2.0/{self.tenant}/{self.env}"
        self._token, self._exp = None, 0
        self.session = requests.Session()
        self._company_name = e.get("BC_COMPANY_NAME")

    def token(self):
        if not self._token or time.time() > self._exp - 120:
            r = requests.post(
                f"https://login.microsoftonline.com/{self.tenant}/oauth2/v2.0/token",
                data=dict(grant_type="client_credentials", client_id=self.client_id,
                          client_secret=self.secret, scope=self.scope),
                timeout=30)
            r.raise_for_status()
            j = r.json()
            self._token, self._exp = j["access_token"], time.time() + int(j.get("expires_in", 3600))
        return self._token

    def _get(self, url, params=None):
        for attempt in range(6):
            r = self.session.get(url, params=params, timeout=180, headers={
                "Authorization": f"Bearer {self.token()}",
                "Prefer": "odata.maxpagesize=20000",
                "Accept": "application/json"})
            if r.status_code in (429, 500, 502, 503, 504):
                wait = int(r.headers.get("Retry-After", 2 ** attempt * 2))
                time.sleep(min(wait, 60))
                continue
            if not r.ok:
                raise RuntimeError(f"BC {r.status_code} {url}: {r.text[:400]}")
            return r.json()
        raise RuntimeError(f"BC request kept failing: {url}")

    @property
    def company_name(self):
        if not self._company_name:
            j = self._get(f"{self.base}/api/v2.0/companies({self.company_id})")
            self._company_name = j["name"]
        return self._company_name

    def api(self, entity, params=None):
        return self._paged(f"{self.base}/api/v2.0/companies({self.company_id})/{entity}", params)

    def odata(self, entity, params=None):
        company = quote(self.company_name)
        return self._paged(f"{self.base}/ODataV4/Company('{company}')/{entity}", params)

    def _paged(self, url, params=None, progress=None):
        rows, page = [], 0
        j = self._get(url, params)
        while True:
            rows.extend(j.get("value", []))
            page += 1
            print(f"    page {page}: {len(rows):,} rows", flush=True)
            nxt = j.get("@odata.nextLink")
            if not nxt:
                return rows
            j = self._get(nxt)
