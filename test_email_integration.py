"""Offline application hooks and workbook regressions for business email."""
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from company_public_data import SourceDocument
from test_scraper import app


class EmailIntegrationTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory(dir='.')
        self.addCleanup(self.temporary.cleanup)
        self.folder = Path(self.temporary.name)
        self.manager = app.WorkbookManager(str(self.folder / 'results.xlsx'))
        self.addCleanup(self.manager.wb.close)
        self.worker = app.SearchWorker({key: Mock() for key in (
            'log', 'trace', 'status', 'progress', 'record', 'harvest', 'done',
        )})
        self.worker.context = Mock()
        self.worker.page = Mock()
        self.worker._launch_browser = Mock()
        self.worker._wait_for_manual_login_if_needed = Mock()
        self.worker._google_maps_enrich_record = Mock(side_effect=lambda row, **_: row)
        self.worker._ensure_llm_available = Mock(return_value=(True, 'fixture'))
        self.worker._smart_validate_record = Mock(side_effect=lambda row, **_: row)
        self.web = Mock()
        self.web.search.return_value = []
        self.web.fetch.return_value = None
        self.web.enrich.side_effect = lambda row: row
        self.factory = Mock(return_value=self.web)
        factory_patch = patch('company_public_data.PublicCompanyEnricher', self.factory)
        factory_patch.start()
        self.addCleanup(factory_patch.stop)

    def record(self, **updates):
        row = {header: '-' for header in app.HEADERS}
        row.update({
            'First Name': 'Jane', 'Last Name': 'Doe',
            'Prospect LinkedIn Link': 'https://www.linkedin.com/in/jane-doe',
            'Job Title': 'Chief Information Officer', 'Job Level': 'C-Level',
            'Job Function': 'Information Technology',
            'Company Name': 'Example Industries', 'Company Website': 'https://example.com/',
            'Company Size': '1,001-5,000 employees', 'Industry': 'Software Development',
            'Tenure': 'Jan 2020 - Present · 6 yrs', 'City': 'Boston',
            'State': 'Massachusetts', 'Country': 'United States',
            'Address': '123 Main St, Boston, MA 02110, United States',
            'Zip Code': '02110', 'Phone Number': '+16175550123',
            'Phone Number Link': 'tel:+16175550123',
            'Address Link': 'https://www.google.com/maps/place/Example+Industries/',
            'Revenue': '$12 million', 'Revenue Link': 'https://example.com/annual-report',
            'SIC Code': '0111', 'NAICS Code': '111140',
            'SIC/NAICS Code Link': 'https://directory.example/company/example-industries',
            'Campaign Name': 'Keep campaign', 'Source File Name': 'original.xlsx',
            'Notes': 'Keep manual note', 'Bounce Email Status': 'Not evaluated',
        })
        row.update(updates)
        return row

    def published_source(self, email='jane.doe@example.com'):
        doc = SourceDocument('https://example.com/staff/jane-doe',
                             'Jane Doe | Example Industries',
                             'Jane Doe\nDirector\nEmail: ' + email,
                             headings=['Jane Doe'])
        self.web.search.side_effect = None
        self.web.search.return_value = [{
            'url': doc.url, 'title': doc.title, 'text': doc.text,
            'provider': 'google', 'kind': 'organic',
            'search_url': 'https://www.google.com/search?q=Jane+Doe+Example+Industries+email',
        }]
        self.web.fetch.side_effect = lambda url: doc if url == doc.url else None

    def saved_row(self, manager=None):
        manager = manager or self.manager
        return dict(zip(app.HEADERS, [cell.value for cell in manager.ws[2]]))

    def run_links(self, mode='FAST', manager=None):
        source = self.folder / 'profiles.csv'
        source.write_text('URL\nhttps://www.linkedin.com/in/jane-doe\n', encoding='utf-8')
        self.worker._process_link_file({
            'input_path': str(source), 'max_records': 1,
            'workbook': manager or self.manager, 'extraction_mode': mode,
        })

    def stub_new_profile(self, role=None):
        role = role or self.record()
        role = {**role, 'Location': 'Boston, Massachusetts, United States',
                'Company Profile Link': 'https://www.linkedin.com/company/example-industries/'}
        self.worker._capture_standard_profile = Mock(return_value={})
        self.worker._parse_standard_current_experience = Mock(return_value=role)
        self.worker._read_standard_company_about = Mock(return_value=role)
        return role

    def assert_email_evidence(self, manager, email):
        self.assertEqual(manager.ws.max_row, 2)
        self.assertEqual(manager.ws.max_column, 42)
        self.assertIn('Email Evidence', manager.wb.sheetnames)
        sheet = manager.wb['Email Evidence']
        self.assertEqual([cell.value for cell in sheet[1]], list(app.EMAIL_EVIDENCE_HEADERS))
        self.assertTrue(any(row[3] == email for row in sheet.iter_rows(min_row=2, values_only=True)))

    def test_saved_complete_company_row_retries_only_missing_email(self):
        original = self.record()
        self.manager.add_record(original)
        self.published_source()
        self.worker._capture_standard_profile = Mock(side_effect=AssertionError('No LinkedIn retry needed'))
        self.run_links()
        self.run_links()
        self.worker._capture_standard_profile.assert_not_called()
        self.worker._google_maps_enrich_record.assert_not_called()
        self.web.enrich.assert_not_called()
        self.web.search.assert_called_once()
        self.assert_email_evidence(self.manager, 'jane.doe@example.com')
        self.assertEqual(self.manager.wb['Email Evidence'].max_row, 2)
        self.manager.save(force=True)
        reopened = app.WorkbookManager(str(self.manager.output_path))
        try:
            saved = self.saved_row(reopened)
            self.assertEqual(saved['Business Email ID'], 'jane.doe@example.com')
            self.assertEqual(saved['Email Status'], 'Found online (unverified)')
            for field in app.HEADERS:
                if field not in ('Business Email ID', 'Email Status', 'Notes'):
                    self.assertEqual(saved[field], original[field], field)
            self.assertIn('Keep manual note', saved['Notes'])
            self.assert_email_evidence(reopened, 'jane.doe@example.com')
        finally:
            reopened.wb.close()

    def test_generated_saved_row_upgrades_on_later_run(self):
        row = self.record(**{'Business Email ID': 'jane.doe@example.com',
                             'Email Status': 'Generated (unverified)',
                             'Bounce Email Status': 'Hard bounce'})
        self.manager.add_record(row)
        self.run_links()
        self.assertEqual(self.saved_row()['Email Status'], 'Generated (unverified)')
        self.assertEqual(self.saved_row()['Bounce Email Status'], 'Hard bounce')
        self.published_source('jd742@example.com')
        self.run_links()
        saved = self.saved_row()
        self.assertEqual(saved['Business Email ID'], 'jd742@example.com')
        self.assertEqual(saved['Email Status'], 'Found online (unverified)')
        self.assertEqual(saved['Bounce Email Status'], 'Not evaluated')
        self.assert_email_evidence(self.manager, 'jd742@example.com')

    def test_lookup_exception_preserves_every_existing_cell(self):
        original = self.record()
        self.manager.add_record(original)
        self.web.search.side_effect = RuntimeError('Source unavailable')
        self.run_links()
        self.run_links()
        self.assertEqual(self.saved_row(), original)
        self.assertEqual(self.manager.ws.max_row, 2)
        self.assertNotIn('Email Evidence', self.manager.wb.sheetnames)
        self.assertEqual(self.web.search.call_count, 2)

    def test_manual_email_status_and_bounce_status_are_preserved(self):
        original = self.record(**{'Business Email ID': 'confirmed@example.com',
                                 'Email Status': 'Manually reviewed',
                                 'Bounce Email Status': 'Manual check pending'})
        self.manager.add_record(original)
        self.published_source()
        self.run_links()
        self.assertEqual(self.saved_row(), original)
        self.web.search.assert_not_called()

    def test_new_link_fast_and_smart_enrich_email_before_smart(self):
        for mode in ('FAST', 'SMART'):
            with self.subTest(mode=mode):
                manager = app.WorkbookManager(str(self.folder / ('link-' + mode + '.xlsx')))
                try:
                    self.stub_new_profile()
                    self.published_source()
                    def smart(row, **kwargs):
                        self.assertEqual(row['Business Email ID'], 'jane.doe@example.com')
                        self.assertEqual(row['Email Status'], 'Found online (unverified)')
                        return row
                    self.worker._smart_validate_record = Mock(side_effect=smart)
                    self.run_links(mode, manager)
                    self.assert_email_evidence(manager, 'jane.doe@example.com')
                    self.assertEqual(self.saved_row(manager)['Bounce Email Status'], 'Not evaluated')
                    self.assertEqual(self.worker._smart_validate_record.call_count, int(mode == 'SMART'))
                finally:
                    manager.wb.close()

    def test_new_sales_fast_and_smart_enrich_email_before_smart(self):
        for mode in ('FAST', 'SMART'):
            with self.subTest(mode=mode):
                manager = app.WorkbookManager(str(self.folder / ('sales-' + mode + '.xlsx')))
                try:
                    role = self.record(**{'Prospect LinkedIn Link': 'https://www.linkedin.com/sales/lead/jane,NAME_SEARCH'})
                    role['Company Profile Link'] = 'https://www.linkedin.com/company/example-industries/'
                    self.worker._collect_current_cards = Mock(return_value=[{'url': role['Prospect LinkedIn Link']}])
                    self.worker._parse_card_fast = Mock(return_value=role)
                    self.worker._open_lead_panel = Mock()
                    self.worker._close_lead_panel = Mock()
                    self.worker._capture_profile_panel = Mock(return_value={'text': '', 'links': []})
                    self.worker._parse_experience_blocks = Mock(return_value=[role])
                    self.worker._read_company_details = Mock(return_value=role)
                    self.published_source()
                    def smart(row, **kwargs):
                        self.assertEqual(row['Business Email ID'], 'jane.doe@example.com')
                        self.assertEqual(row['Email Status'], 'Found online (unverified)')
                        return row
                    self.worker._smart_validate_record = Mock(side_effect=smart)
                    self.worker._process_search({
                        'search_url': 'https://www.linkedin.com/sales/search/people',
                        'max_records': 1, 'workbook': manager, 'extraction_mode': mode,
                    })
                    self.assert_email_evidence(manager, 'jane.doe@example.com')
                    self.assertEqual(self.saved_row(manager)['Bounce Email Status'], 'Not evaluated')
                    self.assertEqual(self.worker._smart_validate_record.call_count, int(mode == 'SMART'))
                finally:
                    manager.wb.close()

    def test_profile_repair_does_not_restore_previous_employer_email(self):
        row = self.record(**{'Company Name': 'Previous Employer',
                             'Company Website': 'https://previous-employer.example',
                             'Tenure': '-', 'Business Email ID': 'jane.doe@previous-employer.example',
                             'Email Status': 'Manually reviewed', 'Bounce Email Status': 'Hard bounce'})
        self.manager.add_record(row)
        self.stub_new_profile()
        self.run_links()
        saved = self.saved_row()
        self.assertEqual(saved['Company Name'], 'Example Industries')
        self.assertEqual(saved['Business Email ID'], '-')
        self.assertEqual(saved['Email Status'], 'Not found')
        self.assertNotEqual(saved['Bounce Email Status'], 'Hard bounce')
        self.assertEqual(saved['Campaign Name'], 'Keep campaign')

    def test_same_person_and_employer_repair_preserves_manual_email(self):
        row = self.record(**{'Tenure': '-', 'Business Email ID': 'confirmed@example.com',
                             'Email Status': 'Manually reviewed',
                             'Bounce Email Status': 'Manual check pending'})
        self.manager.add_record(row)
        self.stub_new_profile()
        self.run_links()
        saved = self.saved_row()
        self.assertNotEqual(saved['Tenure'], '-')
        self.assertEqual(saved['Business Email ID'], 'confirmed@example.com')
        self.assertEqual(saved['Email Status'], 'Manually reviewed')
        self.assertEqual(saved['Bounce Email Status'], 'Manual check pending')
        self.web.search.assert_not_called()

    def test_corrected_person_at_same_employer_does_not_inherit_manual_email(self):
        row = self.record(**{'Tenure': '-', 'Business Email ID': 'jane.doe@example.com',
                             'Email Status': 'Manually reviewed'})
        self.manager.add_record(row)
        self.stub_new_profile(self.record(**{'First Name': 'John', 'Last Name': 'Smith'}))
        self.run_links()
        saved = self.saved_row()
        self.assertEqual((saved['First Name'], saved['Last Name']), ('John', 'Smith'))
        self.assertEqual(saved['Business Email ID'], '-')
        self.assertEqual(saved['Email Status'], 'Not found')

    def test_no_context_and_stop_skip_lookup(self):
        original = self.record()
        self.worker.context = None
        self.assertEqual(self.worker._business_email_enrich_record(original, 'fixture'), original)
        self.worker.context = Mock()
        self.worker.stop_event.set()
        self.assertEqual(self.worker._business_email_enrich_record(original, 'fixture'), original)
        self.web.search.assert_not_called()


if __name__ == '__main__':
    unittest.main()
