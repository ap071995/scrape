"""Crash recovery, cached identity isolation, parallel ownership and headless controls."""
import json
from pathlib import Path
import tempfile
import threading
import time
import unittest
from unittest.mock import Mock, patch

from run_support import EvidenceCache, EnrichmentSessions, RecoveryJournal
from test_scraper import app


class RunSupportTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(dir='.')
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / 'results.xlsx'
        self.worker = app.SearchWorker({})

    def record(self, suffix='person'):
        return {'First Name': 'Sejal', 'Last Name': 'Pattni', 'Job Title': 'Director',
                'Tenure': 'Jul 2013 - Present', 'Company Name': 'UBS',
                'Company Website': 'https://www.ubs.com/',
                'Prospect LinkedIn Link': 'https://www.linkedin.com/in/' + suffix,
                'City': 'London', 'State': 'England', 'Country': 'United Kingdom',
                'Address': 'Postbox, Zurich, 8001, CH', 'Zip Code': '8001',
                'Phone Number': '+41442341111', 'Phone Number Link': 'tel:+41442341111',
                'Address Link': 'https://www.linkedin.com/company/ubs/about/'}

    def details(self):
        return {'company_matched': True, 'company_name': 'UBS AG',
                'company_website': 'https://www.ubs.com/uk/en.html', 'location_matched': True,
                'address': '5 Broadgate, London EC2M 2QS, United Kingdom',
                'zip_code': 'EC2M 2QS', 'phone': '+442075678000',
                'phone_link': 'tel:+442075678000',
                'address_link': 'https://www.google.com/maps/place/UBS+AG/'}

    def test_cache_survives_restart_expires_and_isolates_values(self):
        clock = [100]
        path = self.path.with_suffix('.sqlite3')
        a = EvidenceCache(path, 'company', 20, clock=lambda: clock[0])
        a[('ubs', 'ubs.com')] = {'Industry': 'Finance'}
        other = EvidenceCache(path, 'maps', 20, clock=lambda: clock[0])
        self.assertNotIn(('ubs', 'ubs.com'), other)
        b = EvidenceCache(path, 'company', 20, clock=lambda: clock[0])
        result = b[('ubs', 'ubs.com')]
        result['Industry'] = 'Wrong'
        self.assertEqual(a[('ubs', 'ubs.com')]['Industry'], 'Finance')
        clock[0] = 121
        self.assertNotIn(('ubs', 'ubs.com'), b)

    def test_failed_cache_short_ttl(self):
        clock = [100]
        cache = EvidenceCache(self.path.with_suffix('.db'), 'maps', clock=lambda: clock[0])
        cache['fail'] = {'company_matched': False}
        cache['success'] = {'company_matched': True, 'address': '5 Broadgate, London', 'phone': '+442075678000'}
        clock[0] += 901
        self.assertNotIn('fail', cache)
        self.assertIn('success', cache)

    def test_batched_rows_recover_after_crash_with_details(self):
        manager = app.WorkbookManager(str(self.path), checkpoint_every=10)
        record = self.record()
        record['_profile_state'] = {'profile': record['Prospect LinkedIn Link'],
                                   'company': 'UBS', 'company_url': 'https://www.linkedin.com/company/ubs/',
                                   'captured_at': '2026-10-15T00:00:00+00:00'}
        manager.add_record(record)
        manager.update_record(2, {'Job Title': 'Executive Director'})
        disk = app.load_workbook(self.path)
        self.assertEqual(disk['Leads'].max_row, 1)
        disk.close()
        manager.wb.close()  # Simulated process loss without a final XLSX save.
        reopened = app.WorkbookManager(str(self.path), checkpoint_every=10)
        self.addCleanup(reopened.wb.close)
        self.assertEqual(reopened.ws.max_row, 2)
        self.assertEqual(reopened.read_record(2)['Job Title'], 'Executive Director')
        self.assertEqual(reopened.read_record(2)['Company LinkedIn URL'], 'https://www.linkedin.com/company/ubs/')
        self.assertFalse(reopened.journal.path.exists())

    def test_replay_after_excel_replace_before_journal_cleanup_is_idempotent(self):
        manager = app.WorkbookManager(str(self.path), checkpoint_every=10)
        manager.add_record(self.record())
        with patch.object(manager.journal, 'clear'):
            manager.save(force=True)
        manager.wb.close()
        reopened = app.WorkbookManager(str(self.path))
        self.addCleanup(reopened.wb.close)
        self.assertEqual(reopened.ws.max_row, 2)

    def test_locked_excel_preserves_journal_and_old_workbook(self):
        manager = app.WorkbookManager(str(self.path), checkpoint_every=10)
        self.addCleanup(manager.wb.close)
        original = self.path.read_bytes()
        manager.add_record(self.record())
        with patch.object(app.os, 'replace', side_effect=PermissionError('locked')):
            with self.assertRaises(PermissionError):
                manager.save(force=True)
        self.assertTrue(manager.journal.path.exists())
        self.assertEqual(self.path.read_bytes(), original)

    def test_torn_final_journal_line_preserves_durable_rows(self):
        journal = RecoveryJournal(self.path)
        journal.append(None, self.record())
        with journal.path.open('ab') as stream:
            stream.write(b'{"row":')
        self.assertEqual(len(journal.read()), 1)

    def test_corrupt_completed_journal_line_raises(self):
        journal = RecoveryJournal(self.path)
        journal.path.write_text('broken\n', encoding='utf-8')
        with self.assertRaises(ValueError):
            journal.read()

    def test_row_index_updates_on_url_change(self):
        manager = app.WorkbookManager(str(self.path))
        self.addCleanup(manager.wb.close)
        manager.add_record(self.record())
        old = self.record()['Prospect LinkedIn Link']
        manager.update_record(2, {'Prospect LinkedIn Link': old + '-new'})
        self.assertIsNone(manager.find_link_row(old))
        self.assertEqual(manager.find_link_row(old + '-new'), 2)
        self.assertFalse(manager.is_duplicate(old))

    def test_missing_company_and_location_do_not_reopen_complete_profile(self):
        manager = app.WorkbookManager(str(self.path))
        self.addCleanup(manager.wb.close)
        row = self.record()
        row.update({'City': '-', 'Company Website': '-', 'Company Size': '-'})
        manager.add_record(row)
        self.assertFalse(manager.profile_needs_visit(row['Prospect LinkedIn Link']))
        manager.update_record(2, {'Job Title': '-'})
        self.assertTrue(manager.profile_needs_visit(row['Prospect LinkedIn Link']))

    def test_profile_expiry(self):
        manager = app.WorkbookManager(str(self.path))
        self.addCleanup(manager.wb.close)
        row = self.record()
        row['_profile_state'] = {'profile': row['Prospect LinkedIn Link'], 'company': 'UBS',
                                 'captured_at': '2020-01-01T00:00:00+00:00'}
        manager.add_record(row)
        self.assertTrue(manager.profile_needs_visit(row['Prospect LinkedIn Link']))

    def test_recent_unresolved_rows_do_not_consume_target_or_repeat_lookup(self):
        manager = app.WorkbookManager(str(self.path))
        self.addCleanup(manager.wb.close)
        manager.add_record(self.record())
        worker = self.worker
        worker.cache_path = self.path.with_suffix('.db')
        worker._install_caches()
        worker._enrich_offsite = Mock(side_effect=lambda row, source: row)
        url = self.record()['Prospect LinkedIn Link']
        self.assertIsNone(worker._enrich_saved_company_row(manager, url, 'resume'))
        self.assertIsNone(worker._enrich_saved_company_row(manager, url, 'resume'))
        worker._enrich_offsite.assert_called_once()
        # A changed company/site must not inherit the previous attempt's skip.
        manager.update_record(2, {'Company Website': 'https://different.example/'})
        worker._enrich_saved_company_row(manager, url, 'resume')
        self.assertEqual(worker._enrich_offsite.call_count, 2)

    def test_cached_company_revalidates_identity_before_reuse(self):
        worker = self.worker
        worker.cache_path = self.path.with_suffix('.db')
        worker._install_caches()
        url = 'https://www.linkedin.com/company/ubs/'
        worker.company_cache['standard|' + app.normalize_url(url)] = {
            'Company Name': 'UBS', 'Company Website': 'https://ubs.com/',
            'Industry': 'Financial Services', 'Company Size': '10,001+ employees'}
        worker._launch_browser = Mock(side_effect=AssertionError('Cached company must not open browser'))
        result = worker._read_standard_company_about(url, 'UBS AG')
        self.assertEqual(result['Company Name'], 'UBS')
        with self.assertRaises(app.CompanyPageMismatchError):
            worker._read_standard_company_about(url, 'Unrelated Company')

    def test_headless_launch_passes_flag_and_switch_keeps_enrichment_workers(self):
        worker = self.worker
        worker.pw = Mock()
        context = worker.pw.chromium.launch_persistent_context.return_value
        context.pages = [Mock()]
        worker.headless = True
        worker._launch_browser()
        self.assertTrue(worker.pw.chromium.launch_persistent_context.call_args.kwargs['headless'])
        sessions = worker.enrichment_sessions = Mock()
        worker.headless = False
        worker._launch_browser()
        self.assertFalse(worker.pw.chromium.launch_persistent_context.call_args.kwargs['headless'])
        sessions.close.assert_not_called()

    def test_missing_company_fields_retry_company_only_from_saved_url(self):
        manager = app.WorkbookManager(str(self.path))
        self.addCleanup(manager.wb.close)
        row = self.record()
        row['Company Website'] = '-'
        row['_profile_state'] = {'profile': row['Prospect LinkedIn Link'], 'company': 'UBS',
                                 'company_url': 'https://www.linkedin.com/company/ubs/'}
        manager.add_record(row)
        worker = self.worker
        worker._capture_standard_profile = Mock(side_effect=AssertionError('No profile visit needed'))
        worker._read_standard_company_about = Mock(return_value={'Company Website': 'https://ubs.com/'})
        worker._enrich_offsite = Mock(side_effect=lambda row, source: row)
        worker._enrich_saved_company_row(manager, row['Prospect LinkedIn Link'], 'resume')
        worker._read_standard_company_about.assert_called_once()
        self.assertEqual(manager.read_record(2)['Company Website'], 'https://ubs.com/')

    def test_ubs_london_replaces_contact_group_preserves_zurich(self):
        worker = self.worker
        worker.prefer_local_office = True
        requests = []
        def lookup(record, source):
            requests.append(dict(record))
            return self.details()
        worker._get_google_maps_company_details = Mock(side_effect=lookup)
        row = worker._google_maps_enrich_record(self.record(), 'screenshot')
        self.assertEqual(row['Address'], self.details()['address'])
        self.assertEqual(row['Phone Number'], '+442075678000')
        self.assertEqual(row['Zip Code'], 'EC2M 2QS')
        self.assertIn('Zurich', row['_previous_company_location']['address'])
        self.assertFalse(row['_maps_evidence']['prospect_office_verified'])
        requested = requests[0]
        self.assertEqual(requested['Address'], '-')
        self.assertIn('London', worker._google_maps_queries(requested)[0]['query'])
        self.assertNotIn('Zurich', worker._google_maps_queries(requested)[0]['query'])
        manager = app.WorkbookManager(str(self.path))
        self.addCleanup(manager.wb.close)
        manager.add_record(row)
        reopened = app.WorkbookManager(str(self.path))
        self.addCleanup(reopened.wb.close)
        self.assertEqual(reopened.read_record(2)['Phone Number'], '+442075678000')
        self.assertIn('Zurich', reopened.read_record(2)['_previous_company_location']['address'])

    def test_local_mismatch_keeps_original_contact_group(self):
        worker = self.worker
        worker.prefer_local_office = True
        details = self.details()
        details['address'] = '5 Main Street, Manchester, United Kingdom'
        worker._get_google_maps_company_details = Mock(return_value=details)
        original = self.record()
        result = worker._google_maps_enrich_record(dict(original), 'wrong city')
        self.assertEqual(result['Address'], original['Address'])
        self.assertEqual(result['Phone Number'], original['Phone Number'])
        self.assertFalse(result['_maps_evidence']['is_mergeable'])

    def test_two_offices_do_not_share_maps_cache(self):
        a, b = self.record(), self.record('other')
        b['City'] = 'Manchester'
        self.assertNotEqual(self.worker._google_maps_cache_key(a), self.worker._google_maps_cache_key(b))

    def test_profile_snapshot_reparse_without_network(self):
        worker = self.worker
        worker.profile_cache = EvidenceCache(self.path.with_suffix('.db'), 'profile_capture')
        worker._launch_browser = Mock()
        worker._capture_standard_profile = Mock(return_value={'name': 'Sejal Pattni'})
        url = self.record()['Prospect LinkedIn Link']
        worker._capture_profile_once(url)
        worker._capture_profile_once(url)
        worker._capture_standard_profile.assert_called_once()
        worker.refresh_profiles = True
        worker._capture_profile_once(url)
        self.assertEqual(worker._capture_standard_profile.call_count, 2)

    def test_headless_authwall_exits_immediately(self):
        self.worker.page = Mock(url='https://www.linkedin.com/checkpoint/challenge')
        self.worker.headless = True
        with self.assertRaisesRegex(app.ProfileAccessError, 'Open Chrome'):
            self.worker._wait_for_manual_login_if_needed()

    def test_login_always_opens_visible(self):
        self.worker.headless = True
        self.worker.page = Mock()
        self.worker._launch_browser = Mock()
        self.worker._open_login()
        self.assertFalse(self.worker.headless)

    def test_sessions_overlap_own_threads_and_merge_only_their_fields(self):
        barrier = threading.Barrier(2)
        ownership = {}
        class Session:
            def __init__(self, stage):
                self.stage = stage
                self.thread = threading.get_ident()
                ownership[stage] = self.thread
            def process(self, row):
                assert threading.get_ident() == self.thread
                barrier.wait(timeout=5)
                if self.stage == 'maps':
                    row.update({'Address': 'London', 'Phone Number': '+442075678000', 'Notes': 'Maps source'})
                else:
                    row.update({'Revenue': '$1 million', 'Address': 'wrong', 'Notes': 'Revenue source'})
                return row
            def close(self):
                assert threading.get_ident() == self.thread
        sessions = EnrichmentSessions(Session)
        try:
            result = sessions.enrich(self.record())
        finally:
            sessions.close()
        self.assertNotEqual(ownership['maps'], ownership['company_email'])
        self.assertEqual(result['Address'], 'London')
        self.assertEqual(result['Revenue'], '$1 million')
        self.assertIn('Maps source', result['Notes'])
        self.assertIn('Revenue source', result['Notes'])

    def test_sequential_session_mode_and_error_cleanup(self):
        calls = []
        class Session:
            def __init__(self, stage):
                self.stage = stage
            def process(self, row):
                calls.append(self.stage)
                if self.stage == 'maps':
                    return {**row, 'Address': 'London'}
                assert row['Address'] == 'London'
                raise RuntimeError('fixture failure')
            def close(self):
                calls.append('close:' + self.stage)
        sessions = EnrichmentSessions(Session, concurrent=False)
        try:
            with self.assertRaises(RuntimeError):
                sessions.enrich(self.record())
        finally:
            sessions.close()
        self.assertEqual(calls[:2], ['maps', 'company_email'])
        self.assertIn('close:maps', calls)
        self.assertIn('close:company_email', calls)

    def test_partial_session_failure_preserves_other_stage(self):
        class Session:
            def __init__(self, stage):
                self.stage = stage
            def process(self, row):
                if self.stage == 'company_email':
                    raise RuntimeError('fixture failure')
                return {**row, 'Address': '5 Broadgate, London'}
            def close(self):
                pass
        errors = []
        sessions = EnrichmentSessions(Session, on_error=errors.append)
        try:
            result = sessions.enrich(self.record())
        finally:
            sessions.close()
        self.assertEqual(result['Address'], '5 Broadgate, London')
        self.assertIn('company_email enrichment pending', result['Notes'])
        self.assertEqual(len(errors), 1)

    def test_real_isolated_browsers_block_linkedin_and_share_no_cookies(self):
        from enrichment_session import IsolatedEnrichmentSession
        owner = self.worker
        owner.headless = True
        owners = {}
        owner.cache_path = self.path.with_suffix('.sqlite3')
        def factory(stage):
            session = IsolatedEnrichmentSession(owner, stage)
            context = session.worker.context
            self.assertEqual(context.cookies(), [])
            page = context.new_page()
            with self.assertRaises(Exception):
                page.goto('https://www.linkedin.com/in/blocked-fixture', timeout=3000)
            page.close()
            page = context.new_page()
            page.set_content('<main><h1>UBS AG</h1><button data-item-id="address">'
                '5 Broadgate, London EC2M 2QS, United Kingdom</button>'
                '<button data-item-id="phone:tel:+442075678000">Call</button></main>')
            owners[stage] = threading.get_ident()
            def process(row):
                assert owners[stage] == threading.get_ident()
                data = session.worker._extract_google_maps_place(page)
                if stage == 'maps':
                    return {**row, 'Address': data['Full Address'], 'Phone Number': data['Phone Number']}
                return {**row, 'Industry': 'Financial Services'}
            session.process = process
            return session
        sessions = EnrichmentSessions(factory)
        try:
            result = sessions.enrich(self.record())
        finally:
            sessions.close()
        self.assertEqual(result['Phone Number'], '+442075678000')
        self.assertEqual(result['Industry'], 'Financial Services')
        self.assertNotEqual(owners['maps'], owners['company_email'])


if __name__ == '__main__':
    unittest.main()
