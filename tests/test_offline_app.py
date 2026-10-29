import unittest

from app import app


class OfflineAppTests(unittest.TestCase):
    def setUp(self):
        self.client = app.test_client()

    def test_metadata_identifies_synthetic_mode_and_disables_refresh(self):
        response = self.client.get("/api/meta")

        self.assertEqual(response.status_code, 200)
        metadata = response.get_json()
        self.assertTrue(metadata["notes"]["synthetic"])
        self.assertEqual(metadata["notes"]["data_mode"], "SYNTHETIC")
        self.assertFalse(metadata["refresh_available"])

    def test_business_central_refresh_is_rejected_offline(self):
        response = self.client.post("/api/refresh", json={})

        self.assertEqual(response.status_code, 409)
        self.assertEqual(response.get_json()["status"], "offline_demo")


if __name__ == "__main__":
    unittest.main()
