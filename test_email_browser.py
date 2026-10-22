"""Rendered local fixtures for business email extraction and Excel evidence."""
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock

from playwright.sync_api import sync_playwright

from business_email import BusinessEmailEnricher
from company_public_data import PublicCompanyEnricher
from company_search_dom import SEARCH_RESULT_SCRIPT
from test_business_email import company_match, record
from test_scraper import app


EMAIL_HEADERS = [
    'Person', 'Prospect LinkedIn Link', 'Company Name', 'Business Email ID',
    'Method', 'Pattern', 'Source URL', 'Supporting Text', 'Retrieved At',
]


class EmailBrowserWorkbookTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.playwright = sync_playwright().start()
        cls.addClassCleanup(cls.playwright.stop)
        cls.browser = cls.playwright.chromium.launch(channel='chrome', headless=True, chromium_sandbox=True)
        cls.addClassCleanup(cls.browser.close)

    def setUp(self):
        self.context = self.browser.new_context()
        self.addCleanup(self.context.close)
        self.temp = tempfile.TemporaryDirectory(dir='.')
        self.addCleanup(self.temp.cleanup)

    def save_and_reopen(self, result):
        path = Path(self.temp.name) / 'email_results.xlsx'
        manager = app.WorkbookManager(str(path))
        try:
            manager.add_record(result)
            manager.save(force=True)
        finally:
            manager.wb.close()
        reopened = app.WorkbookManager(str(path))
        self.addCleanup(reopened.wb.close)
        self.assertEqual(reopened.ws.max_row, 2)
        self.assertEqual([cell.value for cell in reopened.ws[1]], app.HEADERS)
        self.assertIn('Email Evidence', reopened.wb.sheetnames)
        evidence = reopened.wb['Email Evidence']
        self.assertEqual([cell.value for cell in evidence[1]], EMAIL_HEADERS)
        self.assertGreater(evidence.max_row, 1)
        row = dict(zip(app.HEADERS, [cell.value for cell in reopened.ws[2]]))
        evidence_rows = [dict(zip(EMAIL_HEADERS, values))
                         for values in evidence.iter_rows(min_row=2, values_only=True)]
        return row, evidence_rows

    def test_real_fetch_extracts_person_owned_mailto_then_saves_evidence(self):
        url = 'https://example.com/staff/jane-doe'
        markup = '''<!doctype html><html><head>
          <title>Jane Doe | Example Industries</title></head><body><main>
          <article><h1>Jane Doe</h1><p>Director</p>
            <a href="mailto:jd742@example.com">Email Jane Doe</a>
          </article></main></body></html>'''

        def route_fixture(route):
            if route.request.url == url:
                route.fulfill(status=200, content_type='text/html', body=markup)
            else:
                route.fulfill(status=404, content_type='text/plain', body='No fixture')

        self.context.route('**/*', route_fixture)
        web = PublicCompanyEnricher(self.context, company_match)
        self.addCleanup(web.close)
        web.search = Mock(return_value=[{
            'url': url, 'title': 'Jane Doe | Example Industries',
            'text': 'Jane Doe, Director at Example Industries. View contact information.',
            'provider': 'google', 'kind': 'organic',
            'search_url': 'https://www.google.com/search?q=Jane+Doe+Example+Industries+email',
        }])
        result = BusinessEmailEnricher(web, company_match).enrich(record())
        self.assertEqual(result['Business Email ID'], 'jd742@example.com')
        self.assertEqual(result['Email Status'], 'Found online (unverified)')
        row, evidence_rows = self.save_and_reopen(result)
        self.assertEqual(row['Business Email ID'], 'jd742@example.com')
        self.assertEqual(row['Email Status'], 'Found online (unverified)')
        matching = [item for item in evidence_rows if item['Business Email ID'] == 'jd742@example.com']
        self.assertTrue(matching)
        self.assertEqual(matching[0]['Prospect LinkedIn Link'], record()['Prospect LinkedIn Link'])
        self.assertEqual(matching[0]['Source URL'], url)
        self.assertIn('Jane Doe', matching[0]['Supporting Text'])

    def test_google_rendered_snippet_to_workbook_with_unavailable_target(self):
        page = self.context.new_page()
        page.set_content('''<!doctype html><html><body><main id="search">
          <div class="result"><div><a href="https://example.com/staff/jane-doe">
            <h3>Jane Doe | Example Industries</h3></a></div>
            <p>Jane Doe, Director at Example Industries. Email: jane.doe@example.com</p>
          </div></main></body></html>''')
        rows = page.evaluate(SEARCH_RESULT_SCRIPT, 'google')
        self.assertEqual(len(rows), 1)
        self.assertIn('jane.doe@example.com', rows[0]['text'])
        search_url = 'https://www.google.com/search?q=Jane+Doe+Example+Industries+email'
        for item in rows:
            item.update(provider='google', search_url=search_url)
        web = Mock()
        web.search.return_value = rows
        web.fetch.return_value = None
        result = BusinessEmailEnricher(web, company_match).enrich(record())
        row, evidence_rows = self.save_and_reopen(result)
        self.assertEqual(row['Business Email ID'], 'jane.doe@example.com')
        self.assertEqual(row['Email Status'], 'Found online (unverified)')
        self.assertTrue(any(item['Source URL'] == search_url for item in evidence_rows))
        self.assertTrue(any('jane.doe@example.com' in item['Supporting Text'] for item in evidence_rows))


if __name__ == '__main__':
    unittest.main()
