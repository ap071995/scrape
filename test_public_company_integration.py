"""Offline public-company enrichment integration and workbook regressions."""
import sys
import tempfile
from pathlib import Path
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from test_scraper import app


class PublicCompanyIntegrationTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory(dir='.')
        self.folder = Path(self.temporary.name)
        self.manager = app.WorkbookManager(str(self.folder / 'results.xlsx'))
        self.worker = app.SearchWorker({key: Mock() for key in (
            'log', 'trace', 'status', 'progress', 'record', 'harvest', 'done',
        )})
        self.worker.context = Mock()
        self.worker.page = Mock()
        self.worker._launch_browser = Mock()
        self.worker._wait_for_manual_login_if_needed = Mock()
        self.worker._google_maps_enrich_record = Mock(side_effect=lambda record, **_: record)
        self.worker._business_email_needs_enrichment = Mock(return_value=False)
        self.worker._business_email_enrich_record = Mock(side_effect=lambda record, **_: record)
        self.client = Mock()
        self.client.enrich.side_effect = self.enrich_fixture
        self.factory = Mock(return_value=self.client)
        self.module_patch = patch.dict(sys.modules, {
            'company_public_data': SimpleNamespace(PublicCompanyEnricher=self.factory),
        })
        self.module_patch.start()

    def tearDown(self):
        self.module_patch.stop()
        self.manager.wb.close()
        self.temporary.cleanup()

    def record(self, url='https://www.linkedin.com/in/example'):
        row = {header: '-' for header in app.HEADERS}
        row.update({
            'First Name': 'Jane', 'Last Name': 'Doe', 'Job Title': 'Chief Information Officer',
            'Job Level': 'C-Level', 'Job Function': 'Information Technology',
            'Company Name': 'Example Company', 'Company Website': 'https://example.com/',
            'Industry': 'Software Development', 'Company Size': '1,001-5,000 employees',
            'Tenure': 'Jan 2020 - Present · 6 yrs', 'City': 'Boston',
            'State': 'Massachusetts', 'Country': 'United States',
            'Prospect LinkedIn Link': url, 'Campaign Name': 'Preserve campaign',
            'Source File Name': 'original.xlsx', 'Notes': 'Preserve manual note',
        })
        return row

    def enrich_fixture(self, record):
        result = dict(record)
        result.update({
            'Revenue': '$12 million', 'Revenue Link': 'https://example.com/annual-report',
            'SIC Code': '0111', 'NAICS Code': '111140',
            'SIC/NAICS Code Link': 'https://directory.example/company/example-company',
        })
        result['_company_evidence'] = [
            {'company': result['Company Name'], 'website': result['Company Website'],
             'field': field, 'value': result[field], 'url': result[source],
             'period': 'FY 2025' if field == 'Revenue' else '-',
             'basis': 'reported', 'excerpt': f"{field}: {result[field]}",
             'checked_at': '2026-09-18T00:00:00Z'}
            for field, source in [('Revenue', 'Revenue Link'), ('SIC Code', 'SIC/NAICS Code Link'),
                                  ('NAICS Code', 'SIC/NAICS Code Link')]
        ]
        return result

    def run_links(self, mode='FAST'):
        source = self.folder / 'profiles.csv'
        source.write_text('URL\nhttps://www.linkedin.com/in/example\n', encoding='utf-8')
        self.worker._process_link_file({'input_path': str(source), 'max_records': 1,
                                       'workbook': self.manager, 'extraction_mode': mode})

    def test_existing_complete_profile_enriches_in_place_and_preserves_core(self):
        original = self.record()
        self.manager.add_record(original)
        self.worker._capture_standard_profile = Mock(side_effect=AssertionError('no LinkedIn rescrape'))
        self.run_links()
        self.assertEqual(self.manager.ws.max_row, 2)
        saved = dict(zip(app.HEADERS, [cell.value for cell in self.manager.ws[2]]))
        for field in app.HEADERS:
            if field not in app.PUBLIC_COMPANY_FIELDS:
                self.assertEqual(saved[field], original[field], field)
        self.assertEqual(saved['Revenue'], '$12 million')
        self.assertEqual(saved['SIC Code'], '0111')
        self.worker._capture_standard_profile.assert_not_called()
        self.worker._google_maps_enrich_record.assert_called_once()
        self.assertEqual(self.manager.wb['Company Evidence'].max_row, 4)
        self.run_links()
        self.client.enrich.assert_called_once()
        self.assertEqual(self.manager.ws.max_row, 2)

    def test_existing_public_complete_row_retries_maps_and_saves_without_duplicate(self):
        original = self.enrich_fixture(self.record())
        self.manager.add_record(original)
        maps_values = {
            'Address': '123 Main St, Boston, MA 02110, United States',
            'Zip Code': '02110', 'Phone Number': '+16175550123',
            'Phone Number Link': 'tel:+16175550123',
            'Address Link': 'https://www.google.com/maps/place/Example+Company/',
        }
        self.worker._google_maps_enrich_record.side_effect = lambda record, **_: {
            **record, **maps_values,
        }
        self.worker._capture_standard_profile = Mock(side_effect=AssertionError('no LinkedIn rescrape'))
        self.run_links()
        self.run_links()
        self.worker._google_maps_enrich_record.assert_called_once()
        self.client.enrich.assert_not_called()
        self.worker._capture_standard_profile.assert_not_called()
        reopened = app.WorkbookManager(str(self.manager.output_path))
        try:
            self.assertEqual(reopened.ws.max_row, 2)
            saved = dict(zip(app.HEADERS, [cell.value for cell in reopened.ws[2]]))
            for field in app.HEADERS:
                self.assertEqual(saved[field], maps_values.get(field, original[field]), field)
            self.assertEqual(reopened.wb['Company Evidence'].max_row, 4)
        finally:
            reopened.wb.close()

    def test_existing_values_retry_missing_maps_source_links_only(self):
        for link_field in ('Address Link', 'Phone Number Link'):
            with self.subTest(link_field=link_field):
                row = self.enrich_fixture(self.record())
                row.update({
                    'Address': '123 Main St, Boston, MA 02110, United States',
                    'Zip Code': '02110', 'Phone Number': '+16175550123',
                    'Address Link': 'https://www.google.com/maps/place/Example+Company/',
                    'Phone Number Link': 'tel:+16175550123',
                })
                expected_link = row[link_field]
                row[link_field] = '-'
                if self.manager.ws.max_row == 1:
                    self.manager.add_record(row)
                else:
                    self.manager.update_record(2, row)
                self.worker._google_maps_enrich_record.reset_mock()
                self.worker._google_maps_enrich_record.side_effect = lambda record, **_: {
                    **record, link_field: expected_link,
                }
                self.run_links()
                self.worker._google_maps_enrich_record.assert_called_once()
                self.client.enrich.assert_not_called()
                self.assertEqual(self.manager.ws.max_row, 2)
                saved = dict(zip(app.HEADERS, [cell.value for cell in self.manager.ws[2]]))
                self.assertEqual(saved[link_field], expected_link)
                for field in app.HEADERS:
                    if field != link_field:
                        self.assertEqual(saved[field], row[field], field)

    def test_maps_retry_predicate_covers_values_links_and_complete_rows(self):
        complete = {
            'Address': '123 Main St, Boston, MA 02110, United States',
            'Zip Code': '02110', 'Phone Number': '+16175550123',
            'Address Link': 'https://www.google.com/maps/place/Example+Company/',
            'Phone Number Link': 'tel:+16175550123',
        }
        self.assertFalse(self.worker._google_maps_needs_enrichment(complete))
        for field in complete:
            with self.subTest(field=field):
                self.assertTrue(self.worker._google_maps_needs_enrichment({**complete, field: '-'}))

    def test_new_link_rows_use_public_sources_before_smart(self):
        for mode in ('FAST', 'SMART'):
            with self.subTest(mode=mode):
                manager = app.WorkbookManager(str(self.folder / f'{mode}.xlsx'))
                previous = self.manager
                self.manager = manager
                try:
                    role = self.record()
                    role.update({'Location': 'Boston, Massachusetts, United States',
                                 'Company Profile Link': 'https://www.linkedin.com/company/example-company/'})
                    self.worker._capture_standard_profile = Mock(return_value={})
                    self.worker._parse_standard_current_experience = Mock(return_value=role)
                    self.worker._read_standard_company_about = Mock(return_value={
                        field: role[field] for field in ('Company Name', 'Company Website', 'Industry', 'Company Size')
                    })
                    self.worker._ensure_llm_available = Mock(return_value=(True, 'fixture'))
                    def smart(record, **kwargs):
                        self.assertEqual(record['Revenue'], '$12 million')
                        return record
                    self.worker._smart_validate_record = Mock(side_effect=smart)
                    self.run_links(mode)
                    self.assertEqual(manager.ws.max_row, 2)
                    self.assertEqual(manager.wb['Company Evidence'].max_row, 4)
                    self.assertEqual(self.worker._smart_validate_record.call_count, int(mode == 'SMART'))
                finally:
                    manager.wb.close()
                    self.manager = previous

    def test_sales_navigator_uses_public_sources_before_smart(self):
        for mode in ('FAST', 'SMART'):
            with self.subTest(mode=mode):
                manager = app.WorkbookManager(str(self.folder / f'sales-{mode}.xlsx'))
                role = self.record('https://www.linkedin.com/sales/lead/example,NAME_SEARCH')
                role['Company Profile Link'] = 'https://www.linkedin.com/company/example-company/'
                self.worker._collect_current_cards = Mock(return_value=[{'url': role['Prospect LinkedIn Link']}])
                self.worker._parse_card_fast = Mock(return_value=role)
                self.worker._open_lead_panel = Mock()
                self.worker._close_lead_panel = Mock()
                self.worker._capture_profile_panel = Mock(return_value={'text': '', 'links': []})
                self.worker._parse_experience_blocks = Mock(return_value=[role])
                self.worker._read_company_details = Mock(return_value=role)
                self.worker._ensure_llm_available = Mock(return_value=(True, 'fixture'))
                def smart(record, **kwargs):
                    self.assertEqual(record['Revenue'], '$12 million')
                    return record
                self.worker._smart_validate_record = Mock(side_effect=smart)
                self.worker._process_search({'search_url': 'https://www.linkedin.com/sales/search/people',
                                             'max_records': 1, 'workbook': manager,
                                             'extraction_mode': mode})
                self.assertEqual(manager.ws.max_row, 2)
                self.assertEqual(manager.wb['Company Evidence'].max_row, 4)
                self.assertEqual(self.worker._smart_validate_record.call_count, int(mode == 'SMART'))
                manager.wb.close()

    def test_optional_lookup_failure_does_not_duplicate_or_erase_existing_row(self):
        original = self.record()
        self.manager.add_record(original)
        self.client.enrich.side_effect = RuntimeError('search unavailable')
        self.run_links()
        self.run_links()
        self.assertEqual(self.manager.ws.max_row, 2)
        self.assertEqual(dict(zip(app.HEADERS, [cell.value for cell in self.manager.ws[2]])), original)
        self.assertEqual(self.client.enrich.call_count, 2)

    def test_evidence_roundtrip_upserts_and_preserves_leading_zero_code(self):
        record = self.enrich_fixture(self.record())
        self.manager.add_record(record)
        self.manager.save(force=True)
        reopened = app.WorkbookManager(str(self.manager.output_path))
        record['_company_evidence'][0]['checked_at'] = '2026-09-19T00:00:00Z'
        reopened.update_record(2, record)
        reopened.save(force=True)
        self.assertEqual(reopened.ws.max_column, 42)
        evidence = reopened.wb['Company Evidence']
        self.assertEqual(evidence.max_row, 4)
        self.assertEqual(evidence['I2'].value, '2026-09-19T00:00:00Z')
        self.assertEqual(evidence['D3'].value, '0111')
        self.assertEqual(evidence['E3'].hyperlink.target, record['SIC/NAICS Code Link'])
        self.assertEqual(reopened.ws.cell(2, app.HEADERS.index('SIC Code') + 1).value, '0111')
        reopened.wb.close()

    def test_smart_preserves_source_links_and_ignores_unsolicited_company_facts(self):
        record = self.enrich_fixture(self.record())
        self.worker._ensure_llm_available = Mock(return_value=(True, 'fixture'))
        self.worker._collect_official_website_evidence = Mock(return_value='-')
        self.worker.smart_job_cache[self.worker._smart_job_cache_key(record)] = {'Job Level': 'C-Level'}
        self.worker.llm.enrich_missing_fields = Mock(return_value={
            'values': {'Revenue': '$999 billion', 'SIC Code': '9999', 'NAICS Code': '999999',
                       'SIC/NAICS Code Link': 'https://wrong.example/999999'},
            'confidence': {field: 100 for field in app.PUBLIC_COMPANY_FIELDS},
        })
        actual = self.worker._smart_validate_record(dict(record), source_label='fixture')
        for field in app.PUBLIC_COMPANY_FIELDS:
            self.assertEqual(actual[field], record[field])
        requested = self.worker.llm.enrich_missing_fields.call_args.kwargs['requested_fields']
        self.assertFalse(set(requested) & set(app.PUBLIC_COMPANY_FIELDS))
        missing = self.record()
        missing['NAICS Code'] = '111140'
        actual = self.worker._smart_validate_record(missing, source_label='fixture')
        self.assertEqual(actual['Revenue'], '-')
        self.assertEqual(actual['SIC Code'], '-')
        self.assertEqual(actual['SIC/NAICS Code Link'], '-')

    def test_no_browser_context_skips_public_fetch_and_close_clears_service(self):
        self.worker.context = None
        record = self.record()
        self.assertEqual(self.worker._public_company_enrich_record(record, 'fixture'), record)
        self.factory.assert_not_called()
        self.worker.public_company_enricher = self.client
        self.worker._close_browser()
        self.client.close.assert_called_once()
        self.assertIsNone(self.worker.public_company_enricher)

    def test_sourced_currencies_do_not_trigger_unnecessary_repeat_enrichment(self):
        record = self.enrich_fixture(self.record())
        for revenue in ('CHF 12 million', 'CNY 80 million', 'US$12 million'):
            record['Revenue'] = revenue
            self.assertFalse(self.worker._public_company_needs_enrichment(record), revenue)


if __name__ == '__main__':
    unittest.main(verbosity=2)
