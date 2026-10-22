"""Offline capacity, custom worksheet and lifecycle tests; no LinkedIn traffic."""
import csv
import tempfile
import time
from pathlib import Path
import unittest
from unittest.mock import Mock, patch

from link_import import load_link_inputs, ImportCancelled
from run_support import RunMetrics
from test_scraper import app


class CapacityAndWorksheetTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)

    def url(self, i):
        return f'https://www.linkedin.com/in/fixture-{i}'

    def book(self):
        result = app.WorkbookManager(str(self.folder / 'results.xlsx'), checkpoint_every=10000)
        self.addCleanup(result.wb.close)
        return result

    def test_excel_1001_rows_all_columns_sheets_hyperlinks_and_sparse_dimensions(self):
        path = self.folder / 'links.xlsx'
        wb = app.Workbook()
        ws = wb.active
        for i in range(1001):
            ws.append([self.url(i), ''])
        ws['XFD1048576'].number_format = '@'  # Excel formatting can inflate dimensions.
        ws['B2'] = 'Profile'
        ws['B2'].hyperlink = self.url(1001)
        ws['C3'] = f'=HYPERLINK("{self.url(1002)}","Person")'
        wb.create_sheet('More links').append([self.url(1003), self.url(1)])
        wb.save(path)
        wb.close()
        started = time.monotonic()
        found = load_link_inputs(path)
        self.assertEqual({item.canonical for item in found}, {self.url(i) for i in range(1004)})
        self.assertLess(time.monotonic() - started, 10)

    def test_csv_98084_entries_and_progress(self):
        path = self.folder / 'large.csv'
        with path.open('w', encoding='utf-8', newline='') as stream:
            csv.writer(stream).writerows([[self.url(i)] for i in range(98084)])
        updates = []
        found = load_link_inputs(path, progress=updates.append)
        self.assertEqual(len(found), 98084)
        self.assertIn('98,084', updates[-1])
        self.assertGreater(len(updates), 100)

    def test_cancel_does_not_silently_return_partial_import(self):
        path = self.folder / 'cancel.csv'
        path.write_text(self.url(1), encoding='utf-8')
        with self.assertRaises(ImportCancelled):
            load_link_inputs(path, cancelled=lambda: True)

    def test_custom_columns_manual_rows_survive_reopen_and_scraper_update(self):
        book = self.book()
        book.add_columns(['Territory', 'User score'])
        rows = book.add_manual_rows(2)
        book.update_record(rows[0], {'First Name': 'Jane', 'Territory': 'West',
            'User score': '=1+1', 'Prospect LinkedIn Link': self.url(1)})
        book.update_record(rows[0], {'Job Title': 'Director'})
        book.save(force=True)
        reopened = app.WorkbookManager(str(book.output_path))
        self.addCleanup(reopened.wb.close)
        self.assertEqual(reopened.ws.max_row, 3)
        self.assertEqual(reopened.headers[-2:], ['Territory', 'User score'])
        self.assertEqual(reopened.read_record(2)['Territory'], 'West')
        self.assertEqual(reopened.read_record(2)['User score'], "'=1+1")
        self.assertEqual(reopened.ws.auto_filter.ref, 'A1:AR3')

    def test_schema_and_linkless_rows_recover_idempotently_after_interrupted_save(self):
        book = self.book()
        book.add_columns(['Custom field'])
        book.add_record({'Custom field': 'Manual A'})
        book.add_record({'Custom field': 'Manual B'})
        book.update_record(2, {'Custom field': 'Updated'})
        journal = book.journal.path.read_bytes()
        recovered = app.WorkbookManager(str(book.output_path))
        self.addCleanup(recovered.wb.close)
        self.assertEqual(recovered.ws.max_row, 3)
        self.assertEqual(recovered.read_record(2)['Custom field'], 'Updated')
        recovered.journal.path.write_bytes(journal)  # Crash between Excel replace and journal clear.
        again = app.WorkbookManager(str(book.output_path))
        self.addCleanup(again.wb.close)
        self.assertEqual(again.ws.max_row, 3)
        self.assertEqual(again.read_record(3)['Custom field'], 'Manual B')

    def test_invalid_columns_and_duplicate_links_are_rejected(self):
        book = self.book()
        for name in ('', 'Company Name', 'company name', '=1+1', '_email_evidence', 'Bad\nname'):
            with self.subTest(name=name), self.assertRaises(ValueError):
                book.add_columns([name])
        book.add_record({'Prospect LinkedIn Link': self.url(1)})
        row = book.add_record({})
        with self.assertRaisesRegex(ValueError, 'Duplicate'):
            book.update_record(row, {'Prospect LinkedIn Link': self.url(1)})

    def test_manual_link_edits_replay_without_creating_duplicate_rows(self):
        book = self.book()
        book.add_columns(['Commentary'])
        row = book.add_record({'Commentary': 'Manual'})
        book.update_record(row, {'Prospect LinkedIn Link': self.url(1)})
        book.update_record(row, {'Prospect LinkedIn Link': self.url(2)})
        journal = book.journal.path.read_bytes()
        recovered = app.WorkbookManager(str(book.output_path))
        self.addCleanup(recovered.wb.close)
        self.assertEqual(recovered.ws.max_row, 2)
        self.assertEqual(recovered.read_record(2)['Prospect LinkedIn Link'], self.url(2))
        recovered.journal.path.write_bytes(journal)
        again = app.WorkbookManager(str(book.output_path))
        self.addCleanup(again.wb.close)
        self.assertEqual(again.ws.max_row, 2)
        self.assertEqual(again.read_record(2)['Prospect LinkedIn Link'], self.url(2))

    def test_75_link_run_writes_every_row_with_all_option(self):
        path = self.folder / 'run.csv'
        path.write_text('\n'.join(self.url(i) for i in range(75)), encoding='utf-8')
        callbacks = {key: Mock() for key in ('status', 'log', 'progress', 'record', 'harvest', 'done', 'stage')}
        worker = app.SearchWorker(callbacks)
        worker._capture_profile_once = Mock(return_value={})
        worker._parse_standard_current_experience = Mock(return_value={
            'First Name': 'Fixture', 'Last Name': 'Person', 'Job Title': 'Director',
            'Company Name': 'Example Industries', 'Tenure': 'Jan 2020 - Present',
            'Company Profile Link': 'https://www.linkedin.com/company/example'})
        worker._read_standard_company_about = Mock(return_value={'Company Website': 'https://example.com'})
        worker._enrich_offsite = lambda record, source_label: record
        worker._log_enrichment_audit = Mock()
        book = self.book()
        worker._process_link_file({'input_path': str(path), 'max_records': None, 'workbook': book})
        self.assertEqual(book.ws.max_row, 76)
        self.assertEqual(worker._capture_profile_once.call_count, 75)
        self.assertEqual(callbacks['progress'].call_args.args, (75, 75, 0, 0, 75))
        self.assertTrue(any(call.args[:2] == ('excel_save', 'done') for call in callbacks['stage'].call_args_list))
        reopened = app.load_workbook(book.output_path)
        self.addCleanup(reopened.close)
        self.assertEqual(reopened['Leads'].max_row, 76)

    def test_stage_errors_finish_bar_and_remain_errors(self):
        callback = Mock()
        metrics = RunMetrics(callback)
        with self.assertRaises(ValueError):
            with metrics.measure('email'):
                raise ValueError('fixture')
        self.assertEqual(callback.call_args_list[0].args[:2], ('email', 'running'))
        self.assertEqual(callback.call_args_list[-1].args[:2], ('email', 'error'))


if __name__ == '__main__':
    unittest.main(verbosity=2)
