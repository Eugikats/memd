import base64
import io
import json
import os
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from docx import Document as DocxDocument
from openpyxl import Workbook, load_workbook

from app import app
from pulse import ministry


def _principal_header(role: str) -> dict:
    payload = {"claims": [{"typ": "roles", "val": role}]}
    return {
        "X-MS-CLIENT-PRINCIPAL-NAME": f"{role}@memd.go.ug",
        "X-MS-CLIENT-PRINCIPAL": base64.b64encode(json.dumps(payload).encode()).decode(),
    }


class BrdMilestoneTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.original_db = ministry.DB
        cls.original_upload_dir = ministry.UPLOAD_DIR
        cls.temp_dir = tempfile.TemporaryDirectory(prefix="memd-brd-test-")
        ministry.DB = Path(cls.temp_dir.name) / "memd.sqlite3"
        ministry.UPLOAD_DIR = Path(cls.temp_dir.name) / "uploaded_templates"
        ministry.initialize()

    @classmethod
    def tearDownClass(cls):
        ministry.DB = cls.original_db
        ministry.UPLOAD_DIR = cls.original_upload_dir
        cls.temp_dir.cleanup()

    def setUp(self):
        self.client = app.test_client()
        env = patch.dict(os.environ, {"AZURE_OPENAI_ENDPOINT": ""})
        env.start()
        self.addCleanup(env.stop)

    def test_auth_and_rbac_enforcement(self):
        with patch.dict(os.environ, {"PULSE_REQUIRE_AUTH": "1"}):
            self.assertEqual(self.client.get("/api/memd/overview").status_code, 401)
            viewer_headers = _principal_header("viewer")
            self.assertEqual(self.client.get("/api/memd/overview", headers=viewer_headers).status_code, 200)
            self.assertEqual(
                self.client.post("/api/memd/indicator", json={"sector": "energy", "metric": "Distribution losses", "period": "2026 Q1", "value": 17.0}, headers=viewer_headers).status_code,
                403,
            )
            self.assertEqual(
                self.client.post("/api/memd/dispositions", json={"decision": "accept", "reason": "Verified"}, headers=viewer_headers).status_code,
                403,
            )

    def test_indicator_metadata_provenance_and_period_compatibility(self):
        ov = self.client.get("/api/memd/overview").get_json()
        sample = ov["indicators"][0]
        self.assertTrue(sample["definition"])
        self.assertTrue(sample["owner"])
        self.assertIn(sample["freshness_status"], ("fresh", "stale"))
        self.assertEqual(sample["data_mode"], "synthetic")
        self.assertGreaterEqual(sample["evidence_count"], 1)

        ev = self.client.get(f"/api/memd/evidence?sector={sample['sector']}&period={sample['period']}").get_json()
        self.assertTrue(ev)
        self.assertTrue(ev[0]["id"].startswith("EV-"))

        with self.assertRaisesRegex(ValueError, "Incompatible reporting intervals"):
            ministry.check_period_compatibility(["2024/25", "2026 Q1"])

    def test_grounded_ai_strips_fabricated_citations_and_logs_audit(self):
        fake_resp = SimpleNamespace(choices=[SimpleNamespace(message=SimpleNamespace(content="Mineral value is 342 [EV-FAK-9999]."))])
        fake_client = SimpleNamespace(chat=SimpleNamespace(completions=SimpleNamespace(create=lambda **kw: fake_resp)))
        with patch.dict(os.environ, {"AZURE_OPENAI_ENDPOINT": "https://example.openai.azure.com"}):
            with patch("pulse.command._azure", return_value=fake_client):
                res = ministry.answer("Summarise mineral production value", actor="analyst@memd.go.ug", role="analyst")
        self.assertIn("[UNVERIFIED]", res["answer"])
        self.assertNotIn("EV-FAK-9999", res["answer"])
        self.assertTrue(any("Unverified citation" in lim for lim in res["limitations"]))
        self.assertTrue(res["citations"])
        audit_rows = ministry.list_audit_events(10)
        self.assertEqual(audit_rows[0]["actor"], "analyst@memd.go.ug")
        self.assertEqual(audit_rows[0]["correlation_id"], res["run_id"])

    def test_excel_placeholder_uses_latest_period_not_oldest(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "{{MINERAL_PRODUCTION_VALUE}}"
        ws["A2"] = "{{MINERAL_PRODUCTION_VALUE::2023/24}}"
        stream = io.BytesIO()
        wb.save(stream)
        resp = self.client.post(
            "/api/memd/reports/template",
            data={"file": (io.BytesIO(stream.getvalue()), "periods.xlsx")},
            content_type="multipart/form-data",
        )
        self.assertEqual(resp.status_code, 200)
        data = resp.get_json()
        self.assertEqual(data["target_period"], "2026 Q1")
        out = self.client.get(data["download_url"])
        gen = load_workbook(io.BytesIO(out.data), data_only=False)
        self.assertEqual(gen.active["A1"].value, 342)
        self.assertEqual(gen.active["A2"].value, 378)
        gen.close()
        wb.close()
        out.close()

    def test_briefing_draft_revision_and_docx_export(self):
        res = self.client.post(
            "/api/memd/briefings",
            json={
                "title": "Q1 Minerals Brief",
                "sector": "minerals",
                "periods": ["2025 Q4", "2026 Q1"],
                "narrative": "Initial draft.",
                "citations": [{"evidence_id": "EV-MIN-0001", "metric": "Mineral production value", "period": "2026 Q1", "source_ref": "Minerals Return 2026 Q1"}],
            },
        )
        self.assertEqual(res.status_code, 201)
        bid = res.get_json()["id"]
        upd = self.client.post(f"/api/memd/briefings/{bid}", json={"narrative": "Updated draft.", "status": "in_review"})
        self.assertEqual(upd.status_code, 200)
        self.assertEqual(upd.get_json()["version"], 2)
        docx_resp = self.client.get(f"/api/memd/briefings/{bid}/docx")
        self.assertEqual(docx_resp.status_code, 200)
        doc = DocxDocument(io.BytesIO(docx_resp.data))
        text = "\n".join(p.text for p in doc.paragraphs)
        self.assertIn("Q1 Minerals Brief", text)
        self.assertIn("Updated draft.", text)
        self.assertIn("EV-MIN-0001", text)
        docx_resp.close()

    def test_rules_exceptions_followups_and_dispositions(self):
        steward_headers = _principal_header("steward")
        rule_resp = self.client.post(
            "/api/memd/rules",
            json={"id": "RULE-STOCK-LOW", "sector": "petroleum", "metric": "Strategic stock cover", "operator": "<", "threshold": 20, "unit": "days"},
            headers=steward_headers,
        )
        self.assertEqual(rule_resp.status_code, 201)
        first_eval = self.client.post("/api/memd/rules/evaluate", headers=steward_headers).get_json()
        second_eval = self.client.post("/api/memd/rules/evaluate", headers=steward_headers).get_json()
        self.assertEqual(len(first_eval["created"]), 1)
        self.assertEqual(len(second_eval["created"]), 0)

        exc_id = first_eval["created"][0]["id"]
        fup = self.client.post("/api/memd/followups", json={"title": "Review reserve replenishment", "source_type": "exception", "source_id": exc_id, "owner": "Director Petroleum", "due_date": "2026-12-01"}).get_json()
        self.assertEqual(fup["status"], "open")
        fup_upd = self.client.post(f"/api/memd/followups/{fup['id']}", json={"status": "resolved", "resolution_notes": "Shipment scheduled"}).get_json()
        self.assertEqual(fup_upd["status"], "resolved")

        dsp = self.client.post(
            "/api/memd/dispositions",
            json={"target_type": "exception", "target_id": exc_id, "decision": "accept", "reason": "Confirmed below 20-day threshold", "rule_version": 1},
            headers=steward_headers,
        )
        self.assertEqual(dsp.status_code, 201)
        self.assertEqual(dsp.get_json()["decision"], "accept")

    def test_orphaned_report_jobs_fail_on_restart(self):
        with ministry.connect() as db:
            db.execute(
                "INSERT INTO report_jobs(id,filename,status,created_at,updated_at) VALUES (?,?,?,?,?)",
                ("orphan-job-1", "stale.xlsx", "processing", "2026-01-01T00:00:00", "2026-01-01T00:00:00"),
            )
        ministry.initialize()
        job = ministry.report_job("orphan-job-1")
        self.assertEqual(job["status"], "failed")
        self.assertIn("restarted", job["error"])


if __name__ == "__main__":
    unittest.main()

