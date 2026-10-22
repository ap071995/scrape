"""Local rendered-DOM fixtures; these do not navigate to a search provider."""
import unittest
import tempfile
from pathlib import Path
from unittest.mock import Mock

from playwright.sync_api import sync_playwright

from company_search_dom import SEARCH_RESULT_SCRIPT
from company_public_data import PublicCompanyEnricher, parse_search_result
from test_scraper import app


class SearchResultDomTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.playwright = sync_playwright().start()
        try:
            cls.browser = cls.playwright.chromium.launch(channel='chrome', headless=True, chromium_sandbox=True)
        except Exception:
            cls.playwright.stop()
            raise

    @classmethod
    def tearDownClass(cls):
        cls.browser.close()
        cls.playwright.stop()

    def setUp(self):
        self.context = self.browser.new_context()
        self.page = self.context.new_page()
        self.addCleanup(self.context.close)

    def collect(self, markup, provider='google'):
        self.page.set_content('<!doctype html><html><body>' + markup + '</body></html>')
        return self.page.evaluate(SEARCH_RESULT_SCRIPT, provider)

    @staticmethod
    def result(name, amount, url='https://directory.example/company'):
        return ('<div class="result"><div><div><div><div><a href="' + url + '"><h3>' +
                name + '</h3></a></div></div></div></div><div class="snippet">' +
                name + ' annual revenue: ' + amount + '</div></div>')

    def test_snippet_more_than_three_ancestors_from_heading(self):
        rows = self.collect('<main id="search">' + self.result('Example Industries', '$12M') + '</main>')
        self.assertEqual(len(rows), 1)
        self.assertIn('annual revenue: $12M', rows[0]['text'])
        self.assertEqual(rows[0]['title'], 'Example Industries')

    def test_two_company_cards_do_not_mix_their_values(self):
        markup = ('<main id="search"><section>' +
                  self.result('Example Industries', '$12M', 'https://example.com/') +
                  self.result('Other Company', '$99B', 'https://other.example/') + '</section></main>')
        rows = self.collect(markup)
        self.assertEqual(len(rows), 2)
        self.assertIn('$12M', rows[0]['text'])
        self.assertNotIn('$99B', rows[0]['text'])
        self.assertNotIn('Other Company', rows[0]['text'])
        self.assertNotIn('$12M', rows[1]['text'])

    def test_hidden_ads_related_questions_and_ai_are_excluded(self):
        bad = self.result('Wrong Company', '$999B')
        markup = ('<main id="search"><div hidden>' + bad + '</div>' +
                  '<div style="display:none">' + bad + '</div>' +
                  '<div aria-hidden="true">' + bad + '</div>' +
                  '<div id="tads">' + bad + '</div>' +
                  '<div class="related-question-pair">' + bad + '</div>' +
                  '<section aria-label="AI Overview">' + bad + '</section>' +
                  self.result('Example Industries', '$12M') + '</main>')
        rows = self.collect(markup)
        self.assertEqual([row['title'] for row in rows], ['Example Industries'])
        self.assertNotIn('Wrong Company', rows[0]['text'])

    def test_adjacent_widget_without_heading_is_not_absorbed(self):
        markup = ('<main id="search"><section>' + self.result('Example Industries', '$12M') +
                  '<div data-ai-overview>Wrong Company annual revenue: $99B</div></section></main>')
        rows = self.collect(markup)
        self.assertEqual(len(rows), 1)
        self.assertNotIn('$99B', rows[0]['text'])

    def test_adjacent_other_entity_heading_is_not_absorbed(self):
        markup = ('<main id="search"><section>' + self.result('Example Industries', '$12M') +
                  '<aside><h2>Other Company</h2>Annual revenue: $99B</aside></section></main>')
        rows = self.collect(markup)
        self.assertEqual(len(rows), 1)
        self.assertNotIn('Other Company', rows[0]['text'])

    def test_organic_company_faq_answer_is_preserved(self):
        markup = ('<main id="search"><div><a href="https://directory.example/company">'
                  '<h3>Example Industries - Company Profile</h3></a>'
                  '<p>What is Example Industries NAICS code? '
                  'The NAICS codes for Example Industries are 541512.</p></div></main>')
        rows = self.collect(markup)
        self.assertEqual(len(rows), 1)
        self.assertIn('are 541512', rows[0]['text'])

    def test_visible_knowledge_property_uses_entity_title_and_suffix_currency(self):
        rows = self.collect('''<main id="search"><aside>
            <div data-attrid="title"><h2>Example Industries</h2></div>
            <div data-attrid="kc:/business/business_operation:revenue">
              <b>Revenue</b><span>48.61 billion USD (2024)</span>
            </div></aside></main>''')
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]['kind'], 'knowledge')
        self.assertEqual(rows[0]['title'], 'Example Industries')
        self.assertIn('48.61 billion USD (2024)', rows[0]['text'])
        self.assertEqual(rows[0]['url'], '')

    def test_knowledge_card_missing_entity_does_not_borrow_query(self):
        rows = self.collect('''<main id="search">
          <input name="q" value="Example Industries revenue">
          <div data-attrid="kc:/business/business_operation:revenue">Revenue: $99B</div>
        </main>''')
        self.assertEqual(rows, [])

    def test_separate_knowledge_panels_keep_entity_identity(self):
        rows = self.collect('''<main id="search">
          <aside><div data-attrid="title">Example Industries</div>
            <div data-attrid="sic_code">SIC Code: 0111</div></aside>
          <aside><div data-attrid="title">Other Company</div>
            <div data-attrid="naics_code">NAICS Code: 541512</div></aside>
        </main>''')
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]['title'], 'Example Industries')
        self.assertNotIn('541512', rows[0]['text'])
        self.assertEqual(rows[1]['title'], 'Other Company')

    def test_featured_answer_includes_scoped_answer_and_source_link(self):
        rows = self.collect('''<main id="search"><div data-featured-snippet>
          <p>Example Industries annual revenue: $12 million (2024)</p>
          <div><a href="https://example.com/results"><h3>Example Industries results</h3></a></div>
        </div></main>''')
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]['kind'], 'featured')
        self.assertIn('$12 million', rows[0]['text'])
        self.assertEqual(rows[0]['url'], 'https://example.com/results')

    def test_bing_organic_cards_have_same_contract(self):
        rows = self.collect('''<ol id="b_results"><li class="b_algo">
          <h2><a href="https://example.com/">Example Industries</a></h2>
          <div>Example Industries revenue: $12M</div></li>
          <li class="b_algo" hidden><h2><a href="https://other.example/">Other</a></h2></li>
        </ol>''', provider='bing')
        self.assertEqual(rows, [{'title': 'Example Industries',
                                 'text': 'Example Industries\nExample Industries revenue: $12M',
                                 'url': 'https://example.com/', 'kind': 'organic'}])

    def test_google_visible_facts_reach_saved_workbook_when_sources_unavailable(self):
        markup = '''<main id="search"><aside>
          <div data-attrid="title">Example Industries</div>
          <div data-attrid="kc:/organization/organization:revenue"><b>Revenue</b><span>48.61 billion USD (2024)</span></div>
        </aside><div><a href="https://directory.example/company"><h3>Example Industries - Company Profile</h3></a>
          <p>SIC Code: 0111; NAICS Code: 111110</p></div></main>'''
        self.page.context.route('https://www.google.com/**', lambda route: route.fulfill(
            status=200, content_type='text/html', body=markup))
        with tempfile.TemporaryDirectory() as folder:
            observations = []
            service = PublicCompanyEnricher(self.page.context, app.SearchWorker._link_company_names_match,
                                            sources_path=Path(folder)/'none.json',
                                            trace=lambda event, **detail: observations.append((event, detail)))
            service.fetch = Mock(return_value=None)
            record = {'Company Name':'Example Industries', 'Company Website':'https://example.com/',
                      'First Name':'Test', 'Prospect LinkedIn Link':'https://www.linkedin.com/in/fixture'}
            try:
                enriched = service.enrich(record)
            finally:
                service.close()
                self.page.context.unroute('https://www.google.com/**')
            self.assertIn('Revenue', enriched, observations)
            self.assertEqual(enriched['Revenue'], 'USD 48.61 billion')
            self.assertEqual(enriched['SIC Code'], '0111')
            self.assertEqual(enriched['NAICS Code'], '111110')
            output = Path(folder)/'result.xlsx'
            manager = app.WorkbookManager(str(output))
            manager.add_record(enriched)
            manager.save(force=True)
            manager.wb.close()
            workbook = app.load_workbook(output, read_only=True)
            try:
                values = list(workbook['Leads'].values)
                saved = dict(zip(values[0], values[1]))
                self.assertEqual(saved['Revenue'], 'USD 48.61 billion')
                self.assertEqual(saved['SIC Code'], '0111')
                evidence = list(workbook['Company Evidence'].values)[1:]
                self.assertEqual(len(evidence), 3)
                self.assertTrue(all(row[4].startswith('https://www.google.com/search?') for row in evidence))
                self.assertTrue(all('search result' in row[6] for row in evidence))
            finally:
                workbook.close()

    def test_featured_result_title_and_snippet_parse_together(self):
        rows = self.collect('''<main id="search"><div data-featured-snippet>
          <p>Example Industries annual revenue: $12 million (2024)</p>
          <a href="https://example.com/results"><h3>Example Industries results</h3></a>
        </div></main>''')
        rows[0].update(provider='google', search_url='https://www.google.com/search?q=Example+Industries+revenue')
        facts = parse_search_result(rows[0], {'Company Name':'Example Industries', 'Company Website':'https://example.com/'},
                                    app.SearchWorker._link_company_names_match)
        self.assertEqual(facts[0]['value'], '$12 million')


if __name__ == '__main__':
    unittest.main()
