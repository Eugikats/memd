import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from pulse import model


class OfflineDemoTests(unittest.TestCase):
    def test_first_model_load_seeds_a_synthetic_snapshot(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp) / "repo"
            raw = root / "data" / "raw"
            root.mkdir()
            with patch.object(model, "ROOT", root), patch.object(model, "RAW", raw), \
                    patch.object(model, "MODEL_PATH", root / "data" / "model.pkl"), \
                    patch.object(model, "_snapshot"):
                result = model.load()

            self.assertEqual(result["notes"]["data_mode"], "SYNTHETIC")
            self.assertTrue(result["notes"]["synthetic"])
            self.assertEqual(result["notes"]["company"], "Pulse Offline Demo")
            self.assertEqual(len(result["items"]), 11)
            self.assertTrue((raw / "run_log.json").exists())
            self.assertTrue((raw / "items.json.gz").exists())

    def test_incomplete_snapshot_is_not_silently_replaced(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp) / "repo"
            raw = root / "data" / "raw"
            raw.mkdir(parents=True)
            (raw / "items.json.gz").write_bytes(b"partial extraction")
            with patch.object(model, "ROOT", root), patch.object(model, "RAW", raw), \
                    patch.object(model, "MODEL_PATH", root / "data" / "model.pkl"):
                with self.assertRaisesRegex(RuntimeError, "snapshot is incomplete"):
                    model.load()
            self.assertEqual((raw / "items.json.gz").read_bytes(), b"partial extraction")


if __name__ == "__main__":
    unittest.main()
