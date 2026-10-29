import io
import tempfile
import unittest
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

from app import app
from pulse import ministry


class MinistryPocTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.original_db = ministry.DB
        cls.original_upload_dir = ministry.UPLOAD_DIR
        cls.temp_dir = tempfile.TemporaryDirectory(prefix='memd-test-')
        ministry.DB = Path(cls.temp_dir.name) / 'memd.sqlite3'
        ministry.UPLOAD_DIR = Path(cls.temp_dir.name) / 'uploaded_templates'
        ministry.initialize()

    @classmethod
    def tearDownClass(cls):
        ministry.DB = cls.original_db
        ministry.UPLOAD_DIR = cls.original_upload_dir
        cls.temp_dir.cleanup()

    def setUp(self):
        self.client = app.test_client()

    def test_seeded_sector_data_is_persisted_and_classified(self):
        response = self.client.get('/api/memd/overview')
        self.assertEqual(response.status_code, 200)
        data = response.get_json()
        self.assertEqual(data['data_label'], 'MEMD workspace')
        self.assertGreaterEqual(data['asset_count'], 5)
        self.assertGreaterEqual(len(data['indicators']), 40)
        self.assertTrue(ministry.DB.exists())

    def test_indicator_update_changes_database_backed_chart_series(self):
        before = self.client.get('/api/memd/overview').get_json()
        record = next(r for r in before['indicators'] if r['sector'] == 'energy' and r['metric'] == 'Distribution losses' and r['period'] == '2026 Q1')
        updated = round(record['value'] + 0.2, 2)
        response = self.client.post('/api/memd/indicator', json={**record, 'value': updated})
        self.assertEqual(response.status_code, 200)
        after = self.client.get('/api/memd/overview').get_json()
        actual = next(r for r in after['indicators'] if r['sector'] == 'energy' and r['metric'] == 'Distribution losses' and r['period'] == '2026 Q1')
        self.assertEqual(actual['value'], updated)
        self.assertEqual(self.client.post('/api/memd/indicator', json={**record, 'value': record['value']}).status_code, 200)

    def test_asset_telemetry_can_be_updated_and_read_back(self):
        response = self.client.post('/api/memd/twin/PS-01/telemetry', json={'metric': 'vibration', 'value': 4.1})
        self.assertEqual(response.status_code, 200)
        twin = self.client.get('/api/memd/twin/PS-01').get_json()
        self.assertEqual(twin['telemetry'][-1]['vibration'], 4.1)
        self.assertEqual(twin['telemetry'][-1]['status'], 'Alert')

    def test_xlsx_template_generation_preserves_formula_and_appends_ministry_data(self):
        wb = Workbook()
        ws = wb.active
        ws.title = 'Quarterly report'
        ws['A1'] = 'Quarterly production value'
        ws['B1'] = '=2+2'
        ws['A2'] = 'Mineral production value'
        ws['A1'].font = Font(bold=True)
        stream = io.BytesIO()
        wb.save(stream)
        response = self.client.post('/api/memd/reports/template', data={'file': (io.BytesIO(stream.getvalue()), 'template.xlsx')}, content_type='multipart/form-data')
        self.assertEqual(response.status_code, 200)
        result = response.get_json()
        self.assertEqual(result['formula_count'], 1)
        self.assertEqual(len(result['matched_cells']), 1)
        output = self.client.get(result['download_url'])
        self.assertEqual(output.status_code, 200)
        generated = load_workbook(io.BytesIO(output.data), data_only=False)
        self.assertEqual(generated['Quarterly report']['B1'].value, '=2+2')
        self.assertEqual(generated['Quarterly report']['B2'].value, 342)
        self.assertTrue(generated['Quarterly report']['A1'].font.bold)
        self.assertIn('MEMD data extract', generated.sheetnames)
        generated.close()
        wb.close()
        output.close()


if __name__ == '__main__':
    unittest.main()
