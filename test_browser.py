"""Real Chromium DOM tests; no login or external network needed."""
import time
import unittest
from test_scraper import app


class BrowserTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.pw = app.sync_playwright().start()
        cls.addClassCleanup(cls.pw.stop)
        cls.browser = cls.pw.chromium.launch(channel='chrome', headless=True, chromium_sandbox=True)
        cls.addClassCleanup(cls.browser.close)

    def setUp(self):
        self.page = self.browser.new_page()
        self.worker = app.SearchWorker({})

    def tearDown(self):
        self.page.close()

    def test_extract_place_dom(self):
        self.page.set_content('''<h1 class="DUwDvf">Example</h1>
        <button data-item-id="address" aria-label="Address: 123 Main St, Boston, MA 02110, USA"></button>
        <button data-item-id="phone:tel:15551234567" aria-label="Phone: +1 555 123 4567"></button>
        <a data-item-id="authority" href="https://example.com">Website</a>''')
        result = self.worker._extract_google_maps_place(self.page)
        self.assertEqual(result['Place Name'], 'Example')
        self.assertEqual(result['Full Address'], '123 Main St, Boston, MA 02110, USA')
        self.assertIn('555', result['Phone Number'])
        self.assertEqual(result['Company Website'], 'https://example.com/')

    def test_late_results_without_initial_scroll_container(self):
        self.page.set_content('''<h1>Search results</h1><script>
        setTimeout(() => {
          document.body.insertAdjacentHTML('beforeend',
            '<div role="feed"><a aria-label="Example" href="https://www.google.com/maps/place/Example" onclick="event.preventDefault()">Example</a></div>');
        }, 250);
        </script>''')
        result = self.worker._find_google_maps_company_result(self.page, 'Example', 'fixture', 'Example')
        self.assertTrue(result['found'])


if __name__ == '__main__':
    unittest.main(verbosity=2)
