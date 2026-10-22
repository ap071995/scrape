"""Real Chromium fixtures, no external requests or account required."""
import csv
from pathlib import Path
import tempfile
import unittest
from unittest.mock import Mock, patch
from test_scraper import app
from linkedin_links import classify_link, extract_link_inputs, LinkResolutionStore

SALES = 'https://www.linkedin.com/sales/lead/ACwAADv0aHgBhjpTrvfVOtoOH4d-brIoiRfKxcM'
PUBLIC = 'https://www.linkedin.com/in/fixture-person'


class MixedInputTests(unittest.TestCase):
    def test_host_and_route_validation(self):
        for url in ('https://notlinkedin.com/in/person', 'https://linkedin.com.evil.com/in/person',
                    'https://linkedin.com@evil.com/in/person', 'https://linkedin.com/in/',
                    'https://linkedin.com/in/a/details/experience', 'https://linkedin.com/sales/lead/',
                    'https://linkedin.com/company/company', 'javascript://linkedin.com/in/person',
                    'https://linkedin.com:999/in/person', 'https://linkedin.com/in/a%2Fb'):
            with self.subTest(url=url):
                self.assertEqual(classify_link(url).kind, 'Invalid')

    def test_regional_and_tracking_links(self):
        self.assertEqual(classify_link('http://ar.linkedin.com/in/fixture-person/?trk=a').canonical, PUBLIC)
        self.assertEqual(classify_link(SALES + ',NAME_SEARCH,anything?_ntb=anything').canonical, SALES)
        self.assertEqual(classify_link('linkedin.com/in/fixture-person').canonical, PUBLIC)

    def test_markdown_and_formula_links(self):
        for value in (f'[{SALES}]({SALES})', f'=HYPERLINK("{SALES}","Profile")'):
            self.assertEqual([x.canonical for x in extract_link_inputs(value)], [SALES])

    def test_excel_hyperlinks_and_invalid_urls_preserved(self):
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / 'input.xlsx'
            book = app.Workbook()
            book.active.append(['Profile URL'])
            book.active.append([PUBLIC])
            book.active.append(['Sales profile'])
            book.active['A3'].hyperlink = SALES + ',NAME_SEARCH,Yjan?_ntb=test'
            book.active.append(['https://wrong.example/in/person'])
            book.save(source)
            book.close()
            worker = app.SearchWorker({})
            values = worker._load_link_inputs(str(source))
            self.assertEqual([x.kind for x in values], ['LinkedIn profile', 'Sales Navigator lead', 'Invalid'])
            self.assertEqual(worker._load_profile_links(str(source)), [PUBLIC, SALES])

    def test_cache_persists_and_invalidates(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'cache.sqlite3'
            first = LinkResolutionStore(path)
            first.put(SALES, PUBLIC)
            reopened = LinkResolutionStore(path)
            self.assertEqual(reopened.get(SALES + ',NAME_SEARCH,Yjan?tracking=a'), PUBLIC)
            reopened.invalidate(SALES)
            self.assertIsNone(first.get(SALES))
            with self.assertRaises(ValueError):
                first.put(SALES, 'https://evil.example/in/person')

    def test_audit_crash_replay_and_update(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'results.xlsx'
            manager = app.WorkbookManager(str(path), checkpoint_every=100)
            manager.record_link_resolution(SALES, PUBLIC, 'Sales Navigator lead', 'Resolved', 'Ready')
            manager.record_link_resolution(SALES, PUBLIC, 'Sales Navigator lead', 'Resolved', 'Saved')
            manager.wb.close()
            recovered = app.WorkbookManager(str(path))
            self.assertEqual(recovered.wb['Link Resolution'].max_row, 2)
            self.assertEqual(recovered.wb['Link Resolution']['E2'].value, 'Saved')
            self.assertEqual(recovered.ws.max_row, 1)
            self.assertFalse(recovered.journal.path.exists())
            recovered.wb.close()

    def test_existing_custom_sheet_is_not_overwritten(self):
        with tempfile.TemporaryDirectory() as folder:
            manager = app.WorkbookManager(str(Path(folder) / 'results.xlsx'))
            manager.wb.create_sheet('Link Resolution').append(['User notes'])
            with self.assertRaisesRegex(ValueError, 'different headers'):
                manager.record_link_resolution(SALES, PUBLIC)
            self.assertEqual(manager.wb['Link Resolution']['A1'].value, 'User notes')
            self.assertFalse(manager.journal.path.exists())
            manager.wb.close()


class MixedBrowserTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.pw = app.sync_playwright().start()
        cls.addClassCleanup(cls.pw.stop)
        cls.browser = cls.pw.chromium.launch(channel='chrome', headless=True, chromium_sandbox=True)
        cls.addClassCleanup(cls.browser.close)

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.context = self.browser.new_context()
        self.addCleanup(self.context.close)
        self.page = self.context.new_page()
        self.worker = app.SearchWorker({key: Mock() for key in ('status', 'log', 'done', 'record', 'harvest', 'progress')})
        self.worker.page = self.page
        self.worker._launch_browser = Mock()
        self.worker.headless = True
        self.worker.link_resolutions = LinkResolutionStore(self.folder / 'cache.sqlite3')
        self.requests = []
        self.html = f'<main><h1>Fixture Person</h1><a href="{PUBLIC}">View LinkedIn profile</a></main>'
        def route(request):
            self.requests.append(request.request.url)
            request.fulfill(content_type='text/html; charset=utf-8', body=self.html)
        self.context.route('**/*', route)
        self.patch = patch.object(app, 'LINK_PROFILE_SETTLE_MS', 1)
        self.patch.start()
        self.addCleanup(self.patch.stop)

    def resolve(self):
        return self.worker._resolve_profile_link(classify_link(SALES))

    def test_direct_never_opens_browser(self):
        self.assertEqual(self.worker._resolve_profile_link(classify_link(PUBLIC)), (PUBLIC, 'Direct'))
        self.worker._launch_browser.assert_not_called()
        self.assertEqual(self.requests, [])

    def test_explicit_link_and_cached_reuse(self):
        self.assertEqual(self.resolve(), (PUBLIC, 'Resolved'))
        self.assertEqual(self.resolve(), (PUBLIC, 'Cached'))
        self.assertEqual(len(self.requests), 1)

    def test_more_menu(self):
        self.html = f'''<main><h1>Fixture Person</h1>
        <button onclick="document.querySelector('nav').style.display='block'">More</button></main>
        <nav role="menu" style="display:none"><a href="{PUBLIC}">View LinkedIn profile</a></nav>'''
        self.assertEqual(self.resolve()[0], PUBLIC)

    def test_button_opens_popup_and_closes_only_owned_popup(self):
        existing = self.context.new_page()
        self.html = f'<main><button onclick="window.open(\'{PUBLIC}\')">View LinkedIn profile</button></main>'
        self.assertEqual(self.resolve()[0], PUBLIC)
        self.assertEqual(set(self.context.pages), {self.page, existing})

    def test_current_sales_overflow_menu_and_delayed_link(self):
        self.html = f'''<main><h1>Fixture Person</h1><button aria-label="More">Unrelated tabs</button>
        <button aria-label="Open actions overflow menu" onclick="setTimeout(() => document.querySelector('#menu').innerHTML = '&lt;a href={PUBLIC}&gt;View LinkedIn profile&lt;/a&gt;', 450)">...</button>
        <div id="menu" role="menu"></div></main>'''
        self.assertEqual(self.resolve()[0], PUBLIC)

    def test_no_guess_from_recommendations(self):
        self.html = f'<main><h1>Fixture Person</h1></main><aside><a href="{PUBLIC}">View LinkedIn profile</a></aside>'
        with self.assertRaisesRegex(ValueError, 'No explicit'):
            self.resolve()
        self.assertIsNone(self.worker.link_resolutions.get(SALES))

    def test_ambiguous_links_not_cached(self):
        self.html += '<a href="https://www.linkedin.com/in/someone-else">View LinkedIn profile</a>'
        with self.assertRaisesRegex(ValueError, 'Multiple'):
            self.resolve()
        self.assertIsNone(self.worker.link_resolutions.get(SALES))

    def test_authwall_is_access_error(self):
        self.html = '<main><h1>Sign in to view</h1><input type="password"></main>'
        with self.assertRaises(app.ProfileAccessError):
            self.resolve()

    def test_missing_lead_can_be_retried(self):
        self.html = '<main>Profile not found</main>'
        with self.assertRaisesRegex(ValueError, 'unavailable'):
            self.resolve()
        self.html = f'<main><a href="{PUBLIC}">View LinkedIn profile</a></main>'
        self.assertEqual(self.resolve()[0], PUBLIC)

    def test_refresh_invalidates_old_mapping(self):
        self.worker.link_resolutions.put(SALES, PUBLIC)
        self.worker.refresh_profiles = True
        self.html = '<main>Profile not found</main>'
        with self.assertRaises(ValueError):
            self.resolve()
        self.assertIsNone(self.worker.link_resolutions.get(SALES))

    def test_mixed_input_through_workbook(self):
        from test_link_scraper import PROFILE, ABOUT
        self.context.unroute('**/*')
        def route(request):
            url = request.request.url
            self.requests.append(url)
            if '/sales/lead/' in url:
                body = self.html if 'ACwAADv0' in url else '<main>Profile not found</main>'
            else:
                body = PROFILE if '/in/' in url else ABOUT
            request.fulfill(content_type='text/html; charset=utf-8', body=body)
        self.context.route('**/*', route)
        source = self.folder / 'mixed.csv'
        with source.open('w', newline='') as stream:
            writer = csv.writer(stream)
            writer.writerows([[x] for x in ('Profile URL', SALES + ',NAME_SEARCH,test?_ntb=x', PUBLIC,
                                           SALES + ',OTHER,test', 'https://www.linkedin.com/sales/lead/missing',
                                           'https://wrong.example/in/person')])
        manager = app.WorkbookManager(str(self.folder / 'results.xlsx'), checkpoint_every=10)
        self.worker._enrich_offsite = lambda record, source_label: record
        with patch.multiple(app, LINK_PROFILE_MAX_SCROLLS=1, LINK_COMPANY_SETTLE_MS=1, LINK_COMPANY_MAX_SCROLLS=1):
            self.worker._process_link_file({'input_path': str(source), 'max_records': 10, 'workbook': manager, 'headless': True})
        reopened = app.load_workbook(manager.output_path)
        self.addCleanup(reopened.close)
        self.assertEqual(reopened['Leads'].max_row, 2)
        lead = dict(zip(app.HEADERS, list(reopened['Leads'].values)[1]))
        self.assertEqual(lead['Prospect LinkedIn Link'], PUBLIC)
        self.assertEqual(lead['Company Name'], 'Cognizant')
        self.assertEqual(lead['Industry'], 'IT Services and IT Consulting')
        audit = list(reopened['Link Resolution'].values)[1:]
        self.assertEqual([r[4] for r in audit], ['Saved', 'Duplicate', 'Duplicate', 'Not processed', 'Not processed'])
        self.assertEqual([r[3] for r in audit], ['Resolved', 'Direct', 'Cached', 'Unresolved', 'Invalid'])
        self.assertEqual(sum('/in/fixture-person' in u for u in self.requests), 1)
        self.assertEqual(sum(SALES in u for u in self.requests), 1)


if __name__ == '__main__':
    unittest.main(verbosity=2)
