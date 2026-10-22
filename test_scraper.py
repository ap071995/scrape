import importlib.util
import os
from pathlib import Path
import shutil
import tempfile
import unittest
from unittest.mock import Mock, patch

os.environ['LOCALAPPDATA'] = str(Path('.test_appdata').resolve())
spec = importlib.util.spec_from_file_location('scraper_under_test', os.environ.get('SCRAPER_TEST_SOURCE', 'Scraper.py'))
app = importlib.util.module_from_spec(spec)
spec.loader.exec_module(app)


class RegressionTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(dir='.')
        self.path = Path(self.temp.name) / 'results.xlsx'
        self.worker = app.SearchWorker({})

    def tearDown(self):
        self.temp.cleanup()

    def record(self, suffix='1'):
        return {'First Name': 'Test', 'Prospect LinkedIn Link': 'https://www.linkedin.com/sales/lead/' + suffix,
                'Phone Number': '+1 555 123 4567', 'Zip Code': '00123'}

    def test_existing_workbook_preserved_and_appended(self):
        source = Path(r'C:/Users/Aditya.paikine/Desktop/Scraper/Sales_Navigator_Results.xlsx')
        if source.exists():
            shutil.copy2(source, self.path)
        else:
            # Keep the regression runnable after the user's sample is moved.
            fixture = app.Workbook()
            fixture.active.title = 'Leads'
            fixture.active.append(app.HEADERS)
            for index in range(5):
                row = self.record('fixture-' + str(index))
                fixture.active.append([row.get(header, '-') for header in app.HEADERS])
            fixture.save(self.path)
            fixture.close()
        before = app.load_workbook(self.path)
        rows = list(before.active.values)
        before.close()
        manager = app.WorkbookManager(str(self.path))
        self.assertEqual(list(manager.ws.values), rows)
        manager.add_record(self.record())
        manager.save(force=True)
        loaded = app.load_workbook(self.path)
        self.assertEqual(list(loaded.active.values)[:len(rows)], rows)
        self.assertEqual(loaded.active.max_row, len(rows) + 1)
        loaded.close()

    def test_row_immediately_persisted(self):
        manager = app.WorkbookManager(str(self.path))
        manager.add_record(self.record())
        loaded = app.load_workbook(self.path)
        self.assertEqual(loaded.active.max_row, 2)
        self.assertEqual(loaded.active['Z2'].value, '00123')
        self.assertEqual(loaded.active['AC2'].value, '+1 555 123 4567')
        self.assertEqual(loaded.active.auto_filter.ref, 'A1:AP2')
        loaded.close()

    def test_duplicate_across_context_and_restart(self):
        manager = app.WorkbookManager(str(self.path))
        manager.add_record(self.record('ABC,NAME_SEARCH,context1?x=1'))
        manager.save(force=True)
        resumed = app.WorkbookManager(str(self.path))
        self.assertTrue(resumed.is_duplicate('http://linkedin.com/sales/lead/ABC,OTHER,context2?x=2'))
        with self.assertRaises(ValueError):
            resumed.add_record(self.record('ABC,OTHER,context2'))

    def test_atomic_save_failure_preserves_file(self):
        manager = app.WorkbookManager(str(self.path))
        manager.add_record(self.record())
        manager.save(force=True)
        original = self.path.read_bytes()
        with patch.object(manager.wb, 'save', side_effect=OSError('disk full')):
            with self.assertRaises(OSError):
                manager.save(force=True)
        self.assertEqual(self.path.read_bytes(), original)
        self.assertEqual(len(list(self.path.parent.glob('*.xlsx'))), 1)

    def test_locked_file_keeps_dirty_rows_for_retry(self):
        manager = app.WorkbookManager(str(self.path))
        with patch.object(app.os, 'replace', side_effect=PermissionError('locked')):
            with self.assertRaises(PermissionError):
                manager.add_record(self.record())
        self.assertGreater(manager._dirty_rows, 0)
        manager.save(force=True)
        loaded = app.load_workbook(self.path)
        self.assertEqual(loaded.active.max_row, 2)
        loaded.close()

    def test_control_events_bypass_queue(self):
        self.worker.send('pause')
        self.assertTrue(self.worker.pause_event.is_set())
        self.worker.send('resume')
        self.assertFalse(self.worker.pause_event.is_set())
        self.worker.send('stop')
        self.assertTrue(self.worker.stop_event.is_set())
        self.assertTrue(self.worker.commands.empty())

    def test_classification_boundaries(self):
        self.assertEqual(self.worker._classify_job_level('VP'), 'VP')
        self.assertEqual(self.worker._classify_job_level('Misleading Specialist'), '-')
        self.assertEqual(self.worker._classify_job_function('Taxi Driver'), '-')
        self.assertEqual(self.worker._classify_job_function('Associate Vice President External Manufacturing'), 'Operations')

    def test_scraped_control_characters_and_formula(self):
        manager = app.WorkbookManager(str(self.path))
        record = self.record()
        record['Company Name'] = 'Example\x00Company'
        record['Description'] = '=HYPERLINK("bad")'
        manager.add_record(record)
        manager.save(force=True)
        loaded = app.load_workbook(self.path)
        self.assertEqual(loaded.active['Q2'].value, 'ExampleCompany')
        self.assertEqual(loaded.active['AN2'].data_type, 's')
        loaded.close()

    def test_empty_maps_result_tries_fallback(self):
        page = Mock()
        page.url = 'https://www.google.com/maps/place/Example'
        self.worker._get_google_maps_page = Mock(return_value=page)
        self.worker._google_maps_queries = Mock(return_value=[{'query': 'one'}, {'query': 'two'}])
        self.worker._dismiss_google_consent = Mock()
        self.worker._wait_and_scroll_google_maps_place = Mock()
        empty = {'Place Name': 'Example', 'Full Address': '-', 'Phone Number': '-'}
        valid = {'Place Name': 'Example', 'Full Address': '123 Main St, Boston, MA 02110, USA',
                 'Phone Number': '+1 555 123 4567', 'Address Link': page.url}
        self.worker._extract_google_maps_place = Mock(side_effect=[empty.copy(), empty.copy(), valid])
        record = {'Company Name': 'Example', 'City': 'Boston', 'State': 'Massachusetts', 'Country': 'United States'}
        result = self.worker._get_google_maps_company_details(record, 'test')
        self.assertEqual(page.goto.call_count, 2)
        self.assertEqual(result['zip_code'], '02110')

    def test_maps_merge_preserves_existing_fields(self):
        record = self.record()
        record.update({'Company Name': 'Example', 'Address': '-', 'Zip Code': '-'})
        self.worker._get_google_maps_company_details = Mock(return_value={
            'company_matched': True, 'address': '123 Main St', 'zip_code': '00123',
            'phone': '+1 555 999 9999', 'address_link': 'https://maps.google.com/example'})
        result = self.worker._google_maps_enrich_record(record, 'test')
        manager = app.WorkbookManager(str(self.path))
        manager.add_record(result)
        manager.save(force=True)
        loaded = app.load_workbook(self.path)
        self.assertEqual(loaded.active['W2'].value, '123 Main St')
        self.assertEqual(loaded.active['Z2'].value, '00123')
        self.assertEqual(loaded.active['AC2'].value, '+1 555 123 4567')
        loaded.close()


if __name__ == '__main__':
    unittest.main(verbosity=2)
