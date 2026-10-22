"""Offline worker-to-Excel checks for email provenance, DNS and office scope."""
import copy
from pathlib import Path
import tempfile
import unittest
from unittest.mock import Mock, patch

from business_email import FOUND, GENERATED, build_email_details
from company_public_data import SourceDocument
from email_domain_check import DomainMXChecker
from test_scraper import app


class FeatureIntegrationTests(unittest.TestCase):
    def setUp(self):
        folder = tempfile.TemporaryDirectory(dir='.')
        self.addCleanup(folder.cleanup)
        self.path = Path(folder.name) / 'features.xlsx'
        self.manager = app.WorkbookManager(str(self.path))
        self.addCleanup(self.manager.wb.close)
        self.worker = app.SearchWorker({})
        self.worker.context = Mock()
        no_network = patch('requests.get', side_effect=AssertionError('Offline test made a network request'))
        no_network.start()
        self.addCleanup(no_network.stop)

    def record(self, **extra):
        return {**{key: '-' for key in app.HEADERS},
                'First Name': 'Jane', 'Last Name': 'Doe',
                'Prospect LinkedIn Link': 'https://www.linkedin.com/in/jane-doe',
                'Company Name': 'Example Industries', 'Company Website': 'https://example.com/',
                'City': 'Boston', 'State': 'Massachusetts', 'Country': 'United States',
                'Notes': 'Keep manual note', **extra}

    def details(self, **extra):
        return {'company_matched': True, 'company_name': 'Example Industries',
                'company_website': 'https://example.com/', 'location_matched': True,
                'address': '123 Main St, Boston, MA 02110, USA', 'zip_code': '02110',
                'phone': '+16175551234', 'phone_link': 'tel:+16175551234',
                'address_link': 'https://www.google.com/maps/place/example', **extra}

    def sheet_row(self, name, row=2, manager=None):
        sheet = (manager or self.manager).wb[name]
        return dict(zip([cell.value for cell in sheet[1]],
                        [cell.value for cell in sheet[row]]))

    def lead_row(self, manager=None):
        manager = manager or self.manager
        return dict(zip(app.HEADERS, [cell.value for cell in manager.ws[2]]))

    def checked_email(self, **extra):
        row = self.record(**{'Business Email ID': 'jane@example.com', 'Email Status': FOUND, **extra})
        query = Mock(return_value={'Status': 0, 'Answer': [
            {'type': 15, 'data': '10 mail.example.com.', 'TTL': 300}]})
        self.worker.email_domain_checker = DomainMXChecker(query=query)
        self.worker.email_domain_checks_enabled = True
        return self.worker._business_email_enrich_record(row, 'fixture'), query

    def test_three_email_origins_remain_separate_in_workbook(self):
        for index, (status, target) in enumerate(((FOUND, 'Published Email ID'),
                (GENERATED, 'Generated Email Candidate'),
                ('Manually reviewed', 'Existing Email (origin unknown)')), start=2):
            row = self.record(**{'Prospect LinkedIn Link': f'https://www.linkedin.com/in/person-{index}',
                                 'Business Email ID': 'jane@example.com', 'Email Status': status})
            self.manager.add_record(row)
            saved = self.sheet_row('Email Details', index)
            self.assertEqual(saved[target], 'jane@example.com')
            for other in ('Published Email ID', 'Generated Email Candidate', 'Existing Email (origin unknown)'):
                if other != target:
                    self.assertEqual(saved[other], '-')
            self.assertEqual(saved['Mailbox Verification'], 'Not verified')
        self.assertEqual(self.manager.ws.max_column, 42)

    def test_published_search_runs_through_worker_and_creates_details(self):
        doc = SourceDocument('https://example.com/staff/jane', 'Jane Doe | Example Industries',
                             'Jane Doe\nDirector\nEmail: jane.doe@example.com', headings=['Jane Doe'])
        web = Mock(source_index={}, documents={})
        web.search.return_value = [{'url': doc.url, 'title': doc.title, 'text': doc.text,
                                  'provider': 'google', 'kind': 'organic'}]
        web.fetch.side_effect = lambda url: doc if url == doc.url else None
        self.worker.public_company_enricher = web
        result = self.worker._business_email_enrich_record(self.record(), 'fixture')
        self.manager.add_record(result)
        self.assertEqual(self.lead_row()['Business Email ID'], 'jane.doe@example.com')
        self.assertEqual(self.sheet_row('Email Details')['Published Email ID'], 'jane.doe@example.com')
        self.assertIn('Email Evidence', self.manager.wb.sheetnames)

    def test_disabled_domain_checks_make_no_dns_queries(self):
        self.worker.email_domain_checker = Mock()
        record = self.record(**{'Business Email ID': 'jane@example.com', 'Email Status': 'Manually reviewed'})
        result = self.worker._business_email_enrich_record(record, 'fixture')
        self.manager.add_record(result)
        self.worker.email_domain_checker.check.assert_not_called()
        self.assertEqual(self.sheet_row('Email Details')['Domain/MX Status'], 'Not checked (disabled)')

    def test_enabled_checks_existing_email_without_searching_or_changing_bounce(self):
        row, query = self.checked_email(**{'Email Status': 'Manually reviewed',
                                          'Bounce Email Status': 'Hard bounce'})
        self.manager.add_record(row)
        self.assertIsNone(self.worker.public_company_enricher)
        query.assert_called_once()
        self.assertEqual(query.call_args.args[:2], ('example.com', 'MX'))
        self.assertEqual(self.sheet_row('Email Details')['Domain/MX Status'], 'MX records found')
        self.assertEqual(self.sheet_row('Email Details')['Mailbox Verification'], 'Not verified')
        self.assertEqual(self.lead_row()['Bounce Email Status'], 'Hard bounce')

    def test_enabled_dns_uses_email_domain_not_company_website_domain(self):
        row, query = self.checked_email(**{'Business Email ID': 'jane@staff.example.net'})
        self.assertEqual(query.call_args.args[0], 'staff.example.net')
        self.assertEqual(row['_email_details']['email_domain'], 'staff.example.net')

    def test_dns_failure_keeps_email_and_lead(self):
        self.worker.email_domain_checks_enabled = True
        self.worker.email_domain_checker = Mock()
        self.worker.email_domain_checker.check.side_effect = TimeoutError('Resolver unavailable')
        row = self.worker._business_email_enrich_record(self.record(**{
            'Business Email ID': 'jane@example.com', 'Email Status': FOUND}), 'fixture')
        self.manager.add_record(row)
        self.assertEqual(self.lead_row()['Business Email ID'], 'jane@example.com')
        self.assertEqual(self.sheet_row('Email Details')['Domain/MX Status'],
                         'Inconclusive (DNS lookup failed)')

    def test_stopped_worker_does_not_start_domain_query(self):
        self.worker.email_domain_checks_enabled = True
        self.worker.email_domain_checker = Mock()
        self.worker.stop_event.set()
        self.worker._business_email_enrich_record(self.record(**{'Business Email ID': 'jane@example.com'}), 'fixture')
        self.worker.email_domain_checker.check.assert_not_called()

    def test_email_details_upsert_and_dns_survive_restart_partial_update(self):
        row, query = self.checked_email()
        self.manager.add_record(row)
        checked = self.sheet_row('Email Details')
        self.manager.save(force=True)
        resumed = app.WorkbookManager(str(self.path))
        self.addCleanup(resumed.wb.close)
        resumed.update_record(2, {'Notes': 'Updated manual note'})
        resumed.save(force=True)
        self.assertEqual(resumed.wb['Email Details'].max_row, 2)
        after = self.sheet_row('Email Details', manager=resumed)
        for field in ('Domain/MX Status', 'Checked At', 'Detail'):
            self.assertEqual(after[field], checked[field], field)
        self.assertEqual(self.lead_row(resumed)['Notes'], 'Updated manual note')
        query.assert_called_once()

    def test_changed_email_does_not_reuse_old_dns_result_even_same_domain(self):
        row, _ = self.checked_email()
        self.manager.add_record(row)
        self.manager.update_record(2, {'Business Email ID': 'new.jane@example.com'})
        detail = self.sheet_row('Email Details')
        self.assertEqual(detail['Published Email ID'], 'new.jane@example.com')
        self.assertEqual(detail['Domain/MX Status'], 'Not checked (disabled)')
        self.assertEqual(detail['Checked At'], '-')

    def test_generation_to_publication_upserts_without_duplicate_or_old_candidate(self):
        row = self.record(**{'Business Email ID': 'jane@example.com', 'Email Status': GENERATED})
        self.manager.add_record(row)
        self.manager.update_record(2, {'Business Email ID': 'jd42@example.com', 'Email Status': FOUND})
        detail = self.sheet_row('Email Details')
        self.assertEqual(self.manager.wb['Email Details'].max_row, 2)
        self.assertEqual(detail['Published Email ID'], 'jd42@example.com')
        self.assertEqual(detail['Generated Email Candidate'], '-')

    def test_wrong_email_details_cannot_associate_different_address_or_dns(self):
        row = self.record(**{'Business Email ID': 'jane@example.com', 'Email Status': FOUND})
        provided = build_email_details({**row, 'Business Email ID': 'someone.else@other.example'})
        provided.update(domain_status='MX records found', checked_at='2026-10-09T00:00:00+00:00',
                        mailbox_status='Verified')
        self.manager.add_record({**row, '_email_details': provided})
        detail = self.sheet_row('Email Details')
        self.assertEqual(detail['Published Email ID'], 'jane@example.com')
        self.assertEqual(detail['Email Domain'], 'example.com')
        self.assertEqual(detail['Domain/MX Status'], 'Not checked (disabled)')
        self.assertEqual(detail['Mailbox Verification'], 'Not verified')

    def test_matching_provided_details_cannot_claim_mailbox_verification(self):
        row = self.record(**{'Business Email ID': 'jane@example.com', 'Email Status': FOUND})
        row['_email_details'] = {**build_email_details(row), 'mailbox_status': 'Verified'}
        self.manager.add_record(row)
        self.assertEqual(self.sheet_row('Email Details')['Mailbox Verification'], 'Not verified')

    def test_bad_email_details_headers_reject_append_atomically(self):
        self.manager.wb.create_sheet('Email Details').append(['Wrong schema'])
        record = self.record(**{'Business Email ID': 'jane@example.com'})
        with self.assertRaisesRegex(ValueError, 'Email Details'):
            self.manager.add_record(record)
        self.assertEqual(self.manager.ws.max_row, 1)
        self.assertFalse(self.manager.is_duplicate(record['Prospect LinkedIn Link']))

    def test_bad_maps_evidence_headers_reject_update_atomically(self):
        row = self.record()
        self.manager.add_record(row)
        original = list(self.manager.ws.values)
        self.manager.wb.create_sheet('Maps Evidence').append(['Wrong schema'])
        self.worker._get_google_maps_company_details = Mock(return_value=self.details())
        enriched = self.worker._google_maps_enrich_record(row, 'fixture')
        with self.assertRaisesRegex(ValueError, 'Maps Evidence'):
            self.manager.update_record(2, enriched)
        self.assertEqual(list(self.manager.ws.values), original)

    def test_hq_fresh_and_cached_keep_company_and_prospect_locations_separate(self):
        original = self.record(City='London', State='-', Country='United Kingdom',
                               Address='Headquarters, 123 Main St, Boston, MA 02110, USA')
        self.worker._get_google_maps_company_details = Mock(return_value=self.details(
            company_name='Example Industries Headquarters'))
        for iteration in range(2):
            enriched = self.worker._google_maps_enrich_record(copy.deepcopy(original), 'fixture')
            if iteration == 0:
                self.manager.add_record(enriched)
            else:
                self.manager.update_record(2, enriched)
            evidence = self.sheet_row('Maps Evidence')
            self.assertEqual(evidence['Location Type'], 'Headquarters')
            self.assertEqual(evidence['Prospect Location'], 'London, United Kingdom')
            self.assertIn('Boston', evidence['Company Location Address'])
            self.assertEqual(evidence['Prospect Works Here Verified'], 'No')
            self.assertEqual(evidence['Eligible for Merge'], 'Yes')
            self.assertEqual(evidence['Source URL'], self.details()['address_link'])
            self.assertIn('saved company address', evidence['Location Basis'])
            self.assertTrue(evidence['Outcome / Reason'])
            self.assertEqual(enriched['Country'], 'United Kingdom')
            self.assertEqual(enriched['Phone Number'], '+16175551234')
        self.worker._get_google_maps_company_details.assert_called_once()
        self.assertEqual(self.manager.wb['Maps Evidence'].max_row, 2)
        self.manager.save(force=True)
        resumed = app.WorkbookManager(str(self.path))
        self.addCleanup(resumed.wb.close)
        resumed.update_record(2, enriched)
        self.assertEqual(resumed.wb['Maps Evidence'].max_row, 2)

    def test_local_office_and_unclassified_are_never_prospect_office_verification(self):
        for label, expected in (('Local office', 'Local office'), ('-', 'Company location (unclassified)')):
            with self.subTest(label=label):
                row = self.record()
                self.worker._apply_google_maps_details(row, self.details(scope_label=label))
                self.assertEqual(row['_maps_evidence']['scope'], expected)
                self.assertFalse(row['_maps_evidence']['prospect_office_verified'])
                self.assertEqual(row['Address'], self.details()['address'])

    def test_phone_only_candidate_is_held_with_persisted_reason(self):
        row = self.record()
        self.worker._get_google_maps_company_details = Mock(return_value=self.details(address='-', address_link='-'))
        enriched = self.worker._google_maps_enrich_record(row, 'fixture')
        self.manager.add_record(enriched)
        self.assertEqual(enriched['Phone Number'], '-')
        self.assertEqual(enriched['Address'], '-')
        evidence = self.sheet_row('Maps Evidence')
        self.assertEqual(evidence['Eligible for Merge'], 'No')
        self.assertEqual(evidence['Prospect Works Here Verified'], 'No')
        self.assertNotEqual(evidence['Outcome / Reason'], '-')
        self.assertEqual(evidence['Fields Applied'], 'None')

    def test_locality_placeholder_can_upgrade_without_claiming_local_office(self):
        row = self.record(Address='Boston, MA')
        self.worker._apply_google_maps_details(row, self.details())
        self.assertEqual(row['Address'], self.details()['address'])
        self.assertEqual(row['_maps_evidence']['scope'], 'Company location (unclassified)')
        self.assertIn('Prospect locality', row['_maps_evidence']['location_basis'])
        self.assertFalse(row['_maps_evidence']['prospect_office_verified'])

    def test_cached_wrong_company_website_is_revalidated_before_merging(self):
        row = self.record()
        self.worker.google_maps_cache[self.worker._google_maps_cache_key(row)] = self.details(
            company_website='https://wrong-company.example/')
        self.worker._get_google_maps_company_details = Mock(side_effect=AssertionError('Expected cached branch'))
        result = self.worker._google_maps_enrich_record(row, 'fixture')
        self.assertEqual(result['Phone Number'], '-')
        self.assertFalse(result['_maps_evidence']['is_mergeable'])
        self.worker._get_google_maps_company_details.assert_not_called()

    def test_actual_maps_collector_rejects_website_conflict_despite_exact_name(self):
        page = Mock()
        page.url = 'https://www.google.com/maps/place/example'
        self.worker._get_google_maps_page = Mock(return_value=page)
        self.worker._google_maps_queries = Mock(return_value=[{'query': 'Example Industries Boston'}])
        self.worker._dismiss_google_consent = Mock()
        self.worker._wait_and_scroll_google_maps_place = Mock()
        self.worker._extract_google_maps_place = Mock(return_value={
            'Place Name': 'Example Industries',
            'Full Address': self.details()['address'], 'Phone Number': '+16175551234',
            'Company Website': 'https://wrong-company.example/', 'Address Link': page.url})
        result = self.worker._get_google_maps_company_details(self.record(), 'fixture')
        self.assertFalse(result['company_matched'])
        self.assertEqual(result['phone'], '-')

    def test_actual_collector_accepts_explicit_office_descriptor_and_labels_it(self):
        page = Mock()
        page.url = 'https://www.google.com/maps/place/example'
        self.worker._get_google_maps_page = Mock(return_value=page)
        self.worker._google_maps_queries = Mock(return_value=[{'query': 'Example Industries Boston'}])
        self.worker._dismiss_google_consent = Mock()
        self.worker._wait_and_scroll_google_maps_place = Mock()
        for suffix in ('Local Office', 'Branch Office', 'Regional Headquarters'):
            with self.subTest(suffix=suffix):
                self.worker._extract_google_maps_place = Mock(return_value={
                    'Place Name': 'Example Industries ' + suffix,
                    'Full Address': self.details()['address'], 'Phone Number': '+16175551234',
                    'Company Website': 'https://example.com/', 'Address Link': page.url})
                self.assertGreater(self.worker._maps_company_match_score(
                    'Example Industries', 'Example Industries ' + suffix), 0)
                record = self.record()
                result = self.worker._get_google_maps_company_details(record, 'fixture')
                self.assertTrue(result['company_matched'])
                self.worker._apply_google_maps_details(record, result)
                self.assertEqual(record['_maps_evidence']['scope'], 'Local office')
                self.assertFalse(record['_maps_evidence']['prospect_office_verified'])
                self.assertEqual(record['Phone Number'], '+16175551234')


if __name__ == '__main__':
    unittest.main()
