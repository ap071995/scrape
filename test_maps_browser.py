"""Maps extraction, branch retries and Excel saving through real browser DOMs."""
import tempfile
from pathlib import Path
import unittest
from unittest.mock import patch

from test_scraper import app


class MapsBrowserTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.pw = app.sync_playwright().start()
        cls.addClassCleanup(cls.pw.stop)
        cls.browser = cls.pw.chromium.launch(channel='chrome', headless=True, chromium_sandbox=True)
        cls.addClassCleanup(cls.browser.close)

    def setUp(self):
        self.context = self.browser.new_context()
        self.context.route('**/*', lambda route: route.abort())
        self.page = self.context.new_page()
        self.worker = app.SearchWorker({})
        self.worker.context = self.context
        self.worker.google_maps_page = self.page
        self.delays = patch.multiple(app, GOOGLE_MAPS_DETAIL_INITIAL_WAIT_MS=30,
            GOOGLE_MAPS_DETAIL_SCROLL_WAIT_MS=50, GOOGLE_MAPS_POST_CLICK_WAIT_MS=30,
            GOOGLE_MAPS_RESULT_WAIT_MS=1000, GOOGLE_MAPS_RESULT_SCAN_ROUNDS=2)
        self.delays.start()

    def tearDown(self):
        self.delays.stop()
        self.context.close()

    @staticmethod
    def panel(address='123 Main St, Boston, MA 02110, USA', phone='+16175551234'):
        return (f'<main><h1 class="DUwDvf">Example Industries</h1>'
                f'<button data-item-id="address" aria-label="Address: {address}"></button>'
                + (f'<button data-item-id="phone:tel:{phone}" aria-label="Call phone number">Call</button>'
                   if phone else '') + '<a data-item-id="authority" href="https://example.com/">Website</a></main>')

    def record(self):
        return {'Company Name':'Example Industries', 'Company Website':'https://example.com/',
                'City':'Boston', 'State':'Massachusetts', 'Country':'United States'}

    def test_phone_uses_number_in_data_attribute_not_call_label(self):
        self.page.set_content(self.panel())
        result = self.worker._extract_google_maps_place(self.page)
        self.assertEqual(result['Phone Number'], '+16175551234')

    def test_tel_link_uses_href_not_call_label(self):
        self.page.set_content('<main><h1>Example</h1><a href="tel:+16175551234">Call</a></main>')
        self.assertEqual(self.worker._extract_google_maps_place(self.page)['Phone Number'], '+16175551234')

    def test_ignores_hidden_old_panels_and_neighbouring_feed_contacts(self):
        self.page.set_content('<main hidden><h1 class="DUwDvf">Old Company</h1>'
            '<button data-item-id="address">Old address</button></main>'
            '<div role="feed"><h1>Search results</h1><button data-item-id="address">Wrong result</button>'
            '<a href="tel:+442012345678">Call</a></div>' + self.panel())
        result = self.worker._extract_google_maps_place(self.page)
        self.assertEqual(result['Place Name'], 'Example Industries')
        self.assertIn('Boston', result['Full Address'])
        self.assertEqual(result['Phone Number'], '+16175551234')

    def test_scrolls_details_instead_of_larger_results_feed(self):
        self.page.set_content('''<div role="feed" style="height:200px;overflow:auto"><div style="height:4000px">Results</div></div>
            <main><h1 class="DUwDvf">Example Industries</h1>
            <div id="details" style="height:200px;overflow:auto"><div style="height:1200px">
            <button data-item-id="address">123 Main St, Boston, MA 02110, USA</button></div></div></main>
            <script>document.querySelector('#details').addEventListener('scroll', () => {
              if (!document.querySelector('a[href^="tel:"]')) document.querySelector('#details').insertAdjacentHTML(
                'afterbegin', '<a href="tel:+16175551234">Call</a>');
            });</script>''')
        state = self.worker._wait_and_scroll_google_maps_place(self.page, 'fixture', 'Example', True, True)
        self.assertTrue(state['phone_found'])
        self.assertEqual(self.page.locator('[role=feed]').evaluate('(e)=>e.scrollTop'), 0)
        self.assertEqual(self.worker._extract_google_maps_place(self.page)['Phone Number'], '+16175551234')

    def serve_candidates(self, first_address, first_phone, second_address, second_phone):
        # Each new search reloads the feed. Selecting a candidate exposes a place panel.
        import json
        html = '<h1>Search results</h1><div role="feed">'
        for index, address, phone in [(1, first_address, first_phone), (2, second_address, second_phone)]:
            html += (f'<a aria-label="Example Industries" href="https://www.google.com/maps/place/{index}" '
                     f'onclick="event.preventDefault(); openPlace({index});">Example Industries</a>')
        html += '</div><script>const panels=' + json.dumps({1:self.panel(first_address, first_phone),
                   2:self.panel(second_address, second_phone)}) + ''';
            function openPlace(i) {document.body.innerHTML=panels[i];
                history.replaceState({}, '', '/maps/place/'+i);}</script>'''
        self.context.unroute('**/*')
        self.context.route('**/*', lambda route: route.fulfill(status=200, content_type='text/html', body=html))
        self.worker._google_maps_queries = lambda record: [{'type':'FIXTURE', 'query':'Example+Boston'}]

    def test_wrong_city_first_result_retries_correct_second_result(self):
        self.serve_candidates('100 Broadway, New York, NY 10001, USA', '+12125551234',
                              '123 Main St, Boston, MA 02110, USA', '+16175551234')
        details = self.worker._get_google_maps_company_details(self.record(), 'fixture')
        self.assertTrue(details['company_matched'])
        self.assertIn('Boston', details['address'])
        self.assertEqual(details['phone'], '+16175551234')
        self.assertTrue(details['address_link'].endswith('/2'))

    def test_missing_phone_first_result_retries_without_losing_address(self):
        address = '123 Main St, Boston, MA 02110, USA'
        self.serve_candidates(address, '', address, '+16175551234')
        details = self.worker._get_google_maps_company_details(self.record(), 'fixture')
        self.assertEqual(details['address'], address)
        self.assertEqual(details['phone'], '+16175551234')

    def test_browser_to_excel_saves_contact_for_equivalent_existing_address(self):
        html = self.panel()
        self.context.unroute('**/*')
        self.context.route('**/*', lambda route: route.fulfill(status=200, content_type='text/html', body=html))
        record = {**self.record(), 'Address':'123 Main Street, Boston, MA 02110-1234, US',
                  'Prospect LinkedIn Link':'https://www.linkedin.com/in/maps-fixture'}
        record = self.worker._google_maps_enrich_record(record, 'fixture')
        with tempfile.TemporaryDirectory(dir='.') as folder:
            path = Path(folder) / 'maps.xlsx'
            manager = app.WorkbookManager(str(path))
            manager.add_record(record)
            manager.wb.close()
            book = app.load_workbook(path)
            try:
                saved = dict(zip(app.HEADERS, [cell.value for cell in book['Leads'][2]]))
                self.assertEqual(saved['Address'], record['Address'])
                self.assertEqual(saved['Zip Code'], '02110')
                self.assertEqual(saved['Phone Number'], '+16175551234')
                self.assertEqual(saved['Phone Number Link'], 'tel:+16175551234')
                self.assertTrue(saved['Address Link'].startswith('https://www.google.com/maps/'))
            finally:
                book.close()

    def test_phone_normalization_rejects_nonphone_digits(self):
        self.assertEqual(self.worker._normalize_maps_phone('Call 1'), '-')
        self.assertEqual(self.worker._normalize_maps_phone('+1 617 555 1234 ext.123'), '+16175551234')

    def test_domain_fallback_can_verify_a_trading_name_on_opened_panel(self):
        self.serve_candidates('123 Main St, Boston, MA 02110, USA', '+16175551234',
                              '123 Main St, Boston, MA 02110, USA', '')
        self.worker._google_maps_queries = lambda record: [{'type':'FALLBACK_WEBSITE', 'query':'example.com+Boston'}]
        record = self.record()
        record['Company Name'] = 'Different Legal Name'
        details = self.worker._get_google_maps_company_details(record, 'fixture')
        self.assertTrue(details['company_matched'])
        self.assertEqual(details['phone'], '+16175551234')

    def test_domain_query_without_matching_website_does_not_prove_identity(self):
        self.serve_candidates('123 Main St, Boston, MA 02110, USA', '+16175551234',
                              '123 Main St, Boston, MA 02110, USA', '+16175551234')
        self.worker._google_maps_queries = lambda record: [{'type':'FALLBACK_WEBSITE', 'query':'different.example+Boston'}]
        record = self.record()
        record.update({'Company Name':'Different Legal Name', 'Company Website':'https://different.example/'})
        details = self.worker._get_google_maps_company_details(record, 'fixture')
        self.assertFalse(details['company_matched'])


if __name__ == '__main__':
    unittest.main(verbosity=2)
