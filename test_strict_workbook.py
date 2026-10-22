"""Reject invalid writes before mutating leads; preserve data across repairs."""
import copy
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from test_scraper import app


class StrictWorkbookTests(unittest.TestCase):
    def setUp(self):
        self.base = Path.cwd().resolve()
        self.temp = tempfile.TemporaryDirectory(prefix='strict workbook ', dir=self.base)
        self.folder = Path(self.temp.name).resolve()
        self.folder.relative_to(self.base)
        self.path = self.folder / 'leads.xlsx'
        self.manager = app.WorkbookManager(str(self.path))

    def tearDown(self):
        self.manager.wb.close()
        self.folder.resolve().relative_to(self.base)
        self.temp.cleanup()

    def row(self, slug='jane-doe'):
        row = {header: '-' for header in app.HEADERS}
        row.update({'First Name': 'Jane', 'Last Name': 'Doe',
                    'Prospect LinkedIn Link': 'https://www.linkedin.com/in/' + slug,
                    'Company Name': 'Example Industries', 'Company Website': 'https://example.com',
                    'Business Email ID': 'jane.doe@example.com', 'Email Status': 'Manually reviewed',
                    'Notes': 'Keep manual note', 'Campaign Name': 'Keep campaign',
                    'Zip Code': '00123', 'Phone Number': '+15551234567'})
        return row

    def evidence(self, row):
        return {'person': 'Jane Doe', 'profile': row['Prospect LinkedIn Link'],
                'company': row['Company Name'], 'email': row['Business Email ID'],
                'method': 'Found online (unverified)', 'pattern': '-',
                'url': 'https://example.com/team', 'excerpt': 'Jane Doe jane.doe@example.com',
                'checked_at': '2026-10-08T00:00:00Z'}

    def state(self):
        return ({s.title: list(s.values) for s in self.manager.wb},
                set(self.manager.seen_links), self.path.read_bytes())

    def saved(self):
        return dict(zip(app.HEADERS, [c.value for c in self.manager.ws[2]]))

    def test_header_or_nonexistent_row_cannot_be_updated(self):
        self.manager.add_record(self.row())
        before = self.state()
        for row_index in (0, 1, -1, 3, 100, None, '2', 2.5):
            with self.subTest(row=row_index):
                with self.assertRaises(ValueError):
                    self.manager.update_record(row_index, self.row())
                self.assertEqual(self.state(), before)

    def test_duplicate_url_correction_is_rejected_without_changes(self):
        self.manager.add_record(self.row('one'))
        self.manager.add_record(self.row('two'))
        before = self.state()
        with self.assertRaises(ValueError):
            self.manager.update_record(2, self.row('two'))
        self.assertEqual(self.state(), before)

    def test_url_correction_updates_duplicate_index_and_survives_restart(self):
        old, new = self.row('old'), self.row('new')
        self.manager.add_record(old)
        self.manager.update_record(2, new)
        self.assertFalse(self.manager.is_duplicate(old['Prospect LinkedIn Link']))
        self.assertTrue(self.manager.is_duplicate(new['Prospect LinkedIn Link']))
        with self.assertRaises(ValueError):
            self.manager.add_record(new)
        self.manager.save(force=True)
        reopened = app.WorkbookManager(str(self.path))
        try:
            self.assertTrue(reopened.is_duplicate(new['Prospect LinkedIn Link']))
            self.assertFalse(reopened.is_duplicate(old['Prospect LinkedIn Link']))
        finally:
            reopened.wb.close()

    def test_partial_update_keeps_unspecified_cells(self):
        original = self.row()
        self.manager.add_record(original)
        self.manager.update_record(2, {'Job Title': 'Director'})
        self.assertEqual(self.saved(), {**original, 'Job Title': 'Director'})

    def test_explicit_empty_field_clears_stale_hyperlink_only(self):
        self.manager.add_record(self.row())
        self.manager.update_record(2, {'Company Website': None})
        col = app.HEADERS.index('Company Website') + 1
        self.assertIsNone(self.manager.ws.cell(2, col).hyperlink)
        self.assertEqual(self.saved()['Company Website'], '-')
        self.assertEqual(self.saved()['Notes'], 'Keep manual note')

    def test_bad_email_sheet_schema_rejects_append_before_mutation(self):
        self.manager.wb.create_sheet('Email Evidence').append(['Incorrect schema'])
        self.manager.save(force=True)
        row = self.row()
        row['_email_evidence'] = [self.evidence(row)]
        before = self.state()
        with self.assertRaises(ValueError):
            self.manager.add_record(row)
        self.assertEqual(self.state(), before)

    def test_bad_company_sheet_schema_rejects_update_before_mutation(self):
        self.manager.add_record(self.row())
        self.manager.wb.create_sheet('Company Evidence').append(['Incorrect schema'])
        self.manager.save(force=True)
        row = self.row()
        row.update({'Company Name': 'New Company', '_company_evidence': [
            {'company':'New Company', 'field':'Revenue', 'value':'$1 million', 'url':'https://example.com/report'}]})
        before = self.state()
        with self.assertRaises(ValueError):
            self.manager.update_record(2, row)
        self.assertEqual(self.state(), before)

    def test_malformed_optional_email_entries_do_not_drop_valid_lead(self):
        for index, entries in enumerate(('bad', {'not':'a list'}, 42, [None, 'bad', {}])):
            with self.subTest(entries=entries):
                row = self.row(str(index))
                row['_email_evidence'] = entries
                self.manager.add_record(row)
        self.assertEqual(self.manager.ws.max_row, 5)
        self.assertNotIn('Email Evidence', self.manager.wb.sheetnames)

    def test_mixed_email_evidence_keeps_valid_entries_only(self):
        row = self.row()
        row['_email_evidence'] = [None, 'bad', {}, self.evidence(row)]
        self.manager.add_record(row)
        self.assertEqual(self.manager.wb['Email Evidence'].max_row, 2)

    def test_evidence_for_a_different_person_or_email_is_not_attached(self):
        row = self.row()
        row['_email_evidence'] = [dict(self.evidence(row), **override) for override in (
            {'profile':'https://www.linkedin.com/in/another-person'},
            {'person':'John Smith'}, {'company':'Other Company'}, {'email':'john.smith@example.com'})]
        self.manager.add_record(row)
        self.assertNotIn('Email Evidence', self.manager.wb.sheetnames)

    def test_same_evidence_upserts_after_reopening(self):
        row = self.row()
        row['_email_evidence'] = [self.evidence(row)]
        self.manager.add_record(row)
        self.manager.wb.close()
        self.manager = app.WorkbookManager(str(self.path))
        row['_email_evidence'][0]['checked_at'] = '2026-10-08T01:00:00Z'
        self.manager.update_record(2, row)
        self.manager.save(force=True)
        self.assertEqual(self.manager.wb['Email Evidence'].max_row, 2)
        self.assertEqual(self.manager.wb['Email Evidence']['I2'].value, '2026-10-08T01:00:00Z')

    def test_update_save_failure_keeps_repair_for_successful_retry(self):
        self.manager.add_record(self.row())
        before = self.path.read_bytes()
        self.manager.update_record(2, {'Job Title':'Director'})
        with patch.object(app.os, 'replace', side_effect=PermissionError('Excel lock')):
            with self.assertRaises(PermissionError):
                self.manager.save(force=True)
        self.assertEqual(self.path.read_bytes(), before)
        self.assertGreater(self.manager._dirty_rows, 0)
        self.manager.save(force=True)
        reopened = app.load_workbook(self.path)
        try:
            row = dict(zip(app.HEADERS, [c.value for c in reopened['Leads'][2]]))
            self.assertEqual(row['Job Title'], 'Director')
            self.assertEqual(row['Campaign Name'], 'Keep campaign')
        finally:
            reopened.close()

    def test_user_sheet_formulas_styles_and_merges_survive_lead_repair(self):
        self.manager.add_record(self.row())
        sheet = self.manager.wb.create_sheet('Manual review')
        sheet['A1'] = '=SUM(1,2)'
        sheet['B1'] = 'Keep this'
        sheet['B1'].font = app.Font(bold=True)
        sheet.merge_cells('B1:D1')
        self.manager.update_record(2, self.row())
        self.manager.save(force=True)
        reopened = app.load_workbook(self.path)
        try:
            sheet = reopened['Manual review']
            self.assertEqual(sheet['A1'].value, '=SUM(1,2)')
            self.assertEqual(sheet['A1'].data_type, 'f')
            self.assertTrue(sheet['B1'].font.bold)
            self.assertIn('B1:D1', sheet.merged_cells)
        finally:
            reopened.close()

    def test_malformed_old_website_does_not_abort_contact_repair(self):
        old, current = self.row(), self.row()
        old['Company Website'] = 'https://[broken'
        current['Business Email ID'] = '-'
        result = app.SearchWorker({})._retain_same_employer_email(current, old)
        self.assertEqual(result['Business Email ID'], '-')

    def test_phone_control_characters_are_removed_without_losing_plus(self):
        row = self.row()
        row['Phone Number'] = '+1\x00555\x0b1234567'
        self.manager.add_record(row)
        self.manager.update_record(2, {'Phone Number': '+1\x015551234567'})
        self.manager.save(force=True)
        reopened = app.load_workbook(self.path)
        try:
            cell = reopened['Leads'].cell(2, app.HEADERS.index('Phone Number') + 1)
            self.assertEqual(cell.value, '+15551234567')
            self.assertEqual(cell.data_type, 's')
        finally:
            reopened.close()

    def test_unknown_profile_markers_do_not_become_duplicate_keys(self):
        for marker in ('-', '', None, 'unknown', 'N/A'):
            row = self.row()
            row['Prospect LinkedIn Link'] = marker
            self.manager.add_record(row)
            self.assertIsNone(self.manager.find_link_row(marker))
        self.assertEqual(self.manager.ws.max_row, 6)
        self.assertFalse(self.manager.seen_links)
        self.manager.save(force=True)
        self.manager.wb.close()
        self.manager = app.WorkbookManager(str(self.path))
        self.assertFalse(self.manager.seen_links)

    def test_invalid_evidence_source_urls_are_ignored(self):
        row = self.row()
        bad_urls = ('https://[broken', 'https://', 'http://127.0.0.1/team',
                    'javascript:alert(1)', 'https://example.com@localhost/team')
        row['_email_evidence'] = [dict(self.evidence(row), url=url) for url in bad_urls]
        row['_company_evidence'] = [
            {'company': row['Company Name'], 'field': 'Revenue', 'value': '$1 million', 'url': url}
            for url in bad_urls]
        self.manager.add_record(row)
        self.assertEqual(self.manager.ws.max_row, 2)
        self.assertNotIn('Email Evidence', self.manager.wb.sheetnames)
        self.assertNotIn('Company Evidence', self.manager.wb.sheetnames)


if __name__ == '__main__':
    unittest.main()
