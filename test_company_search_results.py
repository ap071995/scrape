"""Regression coverage for company facts visible directly in search results.

These tests use saved-format examples, not live searches. Search snippets are
weaker evidence than a fetched source and must carry their own provenance.
"""
import copy
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock

from test_scraper import app
from company_public_data import (PublicCompanyEnricher, SourceDocument,
                                 parse_document, parse_search_result, select_facts)


class SearchResultParsingTests(unittest.TestCase):
    record = {'Company Name': 'Example Industries',
              'Company Website': 'https://example.com/'}
    search_url = 'https://www.google.com/search?q=%22Example+Industries%22+revenue'
    target_url = 'https://directory.example/company/example-industries'
    match = staticmethod(app.SearchWorker._link_company_names_match)

    def row(self, text, **overrides):
        result = {'title': 'Example Industries - Company Profile', 'text': text,
                  'url': self.target_url, 'provider': 'google',
                  'search_url': self.search_url, 'kind': 'organic'}
        result.update(overrides)
        return result

    def parse(self, text, **overrides):
        return parse_search_result(self.row(text, **overrides), self.record, self.match)

    def values(self, facts):
        return {item['field']: item['value'] for item in facts}

    def test_organic_result_has_value_period_and_search_provenance(self):
        facts = self.parse('Example Industries annual revenue was $12.4 billion in 2024.')
        self.assertEqual(self.values(facts), {'Revenue': '$12.4 billion'})
        fact = facts[0]
        self.assertEqual(fact['period'], '2024')
        self.assertEqual(fact['url'], self.search_url)
        self.assertIn('search', fact['basis'].lower())
        self.assertIn(self.target_url, fact['excerpt'])
        self.assertIn('$12.4 billion', fact['excerpt'])
        self.assertLess(fact['_rank'], 0)

    def test_named_knowledge_card_accepts_suffix_currency(self):
        facts = self.parse('Revenue\n48.61 billion USD (2024)',
                           title='Example Industries', kind='knowledge')
        self.assertEqual(self.values(facts), {'Revenue': 'USD 48.61 billion'})
        self.assertEqual(facts[0]['period'], '2024')
        self.assertNotIn('Reported annual', facts[0]['basis'])

    def test_featured_result_can_supply_labelled_classifications(self):
        facts = self.parse('Example Industries\nSIC Code: 0111\nNAICS Code: 111110',
                           kind='featured')
        self.assertEqual(self.values(facts), {'SIC Code': '0111', 'NAICS Code': '111110'})
        for fact in facts:
            self.assertEqual(fact['url'], self.search_url)
            self.assertLess(fact['_rank'], 0)

    def test_company_faq_answer_can_supply_explicit_naics(self):
        facts = self.parse('What is Example Industries NAICS code? '
                           'The NAICS codes for Example Industries are 541512.')
        self.assertEqual(self.values(facts), {'NAICS Code': '541512'})

    def test_estimated_revenue_remains_explicitly_estimated(self):
        facts = self.parse("Example Industries's estimated annual revenue is $10M - $50M.")
        self.assertEqual(self.values(facts), {'Revenue': '$10M - $50M'})
        self.assertIn('estimat', facts[0]['basis'].lower())

    def test_query_company_is_not_identity_evidence(self):
        # The company is present only in the Google URL, not in the result.
        self.assertEqual(self.parse('Revenue: $99 billion\nSIC Code: 1234',
                                    title='Company Revenue and Industry Codes'), [])

    def test_different_company_or_subsidiary_is_rejected(self):
        for name in ('Other Corporation', 'Example Industries UK'):
            with self.subTest(name=name):
                self.assertEqual(self.parse(name + '\nAnnual revenue: $7 billion\nNAICS Code: 541512',
                                            title=name + ' - Revenue'), [])

    def test_parent_mention_does_not_validate_subsidiary_fact(self):
        self.assertEqual(self.parse('SubsidiaryCo, part of Example Industries. '
                                    'Annual revenue: $7 billion. SIC Code: 1234.',
                                    title='SubsidiaryCo - Company Profile'), [])

    def test_explicit_different_website_rejects_company_namesake(self):
        self.assertEqual(self.parse('Example Industries\nWebsite: other.example\n'
                                    'Annual revenue: $9 million\nNAICS Code: 541512'), [])

    def test_unanswered_related_question_is_not_classification(self):
        for text in ('What is Example Industries NAICS code?',
                     'People also ask\nIs Example Industries in NAICS 541512?',
                     'Related questions\nDoes SIC 1234 apply to Example Industries?'):
            with self.subTest(text=text):
                self.assertEqual(self.parse(text), [])

    def test_generic_code_definition_or_crosswalk_is_not_company_fact(self):
        for title, text in [
                ('NAICS Code 541512 - Computer Systems Design Services',
                 'NAICS Code: 541512. Examples include Example Industries.'),
                ('Example Industries - SIC to NAICS crosswalk',
                 'SIC Code: 7371\nNAICS Code: 541511\nCrosswalk and industry definitions.')]:
            with self.subTest(title=title):
                self.assertEqual(self.parse(text, title=title), [])

    def test_broad_sectors_and_unlabelled_numbers_are_rejected(self):
        for text in ('Example Industries\nNAICS: 54\nSIC: 73',
                     'Example Industries\nNAICS: 5415',
                     'Example Industries has 541512 staff and 1234 locations.'):
            with self.subTest(text=text):
                self.assertEqual(self.parse(text), [])

    def test_quarter_forecast_and_non_revenue_metrics_rejected(self):
        for text in ('Example Industries Q2 2025 revenue: $3 billion',
                     'Example Industries annual revenue forecast: $3 billion',
                     'Example Industries annual revenue per employee: $300,000',
                     'Example Industries annual revenue per share: $3.50',
                     'Example Industries assets: $3 billion',
                     'Example Industries net income: $3 billion',
                     'Example Industries revenue\nNot disclosed\nAssets: $3 billion'):
            with self.subTest(text=text):
                self.assertEqual(self.parse(text), [])

    def test_quarterly_revenue_does_not_discard_independent_classification(self):
        facts = self.parse('Example Industries Q2 2025 revenue: $3 billion\nNAICS Code: 541512')
        self.assertEqual(self.values(facts), {'NAICS Code': '541512'})

    def test_currency_cannot_be_inferred_from_company_country(self):
        facts = self.parse('Example Industries\nRevenue: 48.61 billion (2024)',
                           title='Example Industries', kind='knowledge')
        self.assertEqual(facts, [])

    def test_multiple_codes_are_not_arbitrarily_selected(self):
        self.assertEqual(self.parse('Example Industries\nNAICS Code: 541511\nNAICS Code: 541512'), [])

    def test_conflicting_equally_ranked_search_results_remain_unresolved(self):
        first = self.parse('Example Industries annual revenue: $12 million (2024)')
        second = self.parse('Example Industries annual revenue: $18 million (2024)',
                            url='https://another.example/example-industries')
        self.assertTrue(first)
        self.assertTrue(second)
        self.assertEqual(select_facts(first + second), [])

    def test_fetched_page_preferred_over_newer_search_result(self):
        snippet = self.parse('Example Industries annual revenue: $18 million (2025)')
        document = SourceDocument('https://example.com/results', 'Example Industries',
                                  'Example Industries\n2024 annual revenue: $12 million')
        page_facts = parse_document(document, self.record, self.match)
        selected = select_facts(snippet + page_facts)
        self.assertEqual(self.values(selected), {'Revenue': '$12 million'})
        self.assertEqual(selected[0]['url'], document.url)

    def test_parser_does_not_mutate_result_or_record(self):
        row = self.row('Example Industries annual revenue: $12 million (2024)')
        record = copy.deepcopy(self.record)
        original = copy.deepcopy(row)
        parse_search_result(row, record, self.match)
        self.assertEqual(row, original)
        self.assertEqual(record, self.record)

    def test_other_company_sentence_in_matched_result_is_not_borrowed(self):
        facts = self.parse('Example Industries annual revenue: $12 million (2024). '
                           'Other Corporation annual revenue: $99 billion (2025). '
                           'Other Corporation NAICS Code: 541511.')
        self.assertEqual(self.values(facts), {'Revenue':'$12 million'})
        self.assertEqual(facts[0]['period'], '2024')

    def test_company_and_year_subjects_remain_supported(self):
        for text in ('Example Industries reported annual revenue: $12 million (2024)',
                     'The company annual revenue: $12 million (2024)',
                     '2024 annual revenue: $12 million'):
            self.assertEqual(self.values(self.parse(text)), {'Revenue':'$12 million'})


class SearchResultEnrichmentTests(unittest.TestCase):
    record = {'Company Name': 'Example Industries',
              'Company Website': 'https://example.com/',
              'Revenue': '-', 'SIC Code': '-', 'NAICS Code': '-'}
    result_url = 'https://directory.example/company/example-industries'
    search_url = 'https://www.google.com/search?q=Example+Industries+revenue'

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.service = PublicCompanyEnricher(
            Mock(), app.SearchWorker._link_company_names_match,
            sources_path=Path(self.temp.name) / 'no-source-index.json')
        self.addCleanup(self.service.close)
        self.service.search = Mock(return_value=[{
            'title': 'Example Industries - Revenue and Company Profile',
            'text': 'Example Industries annual revenue: $18 million (2025)\n'
                    'SIC Code: 0111\nNAICS Code: 111110',
            'url': self.result_url, 'provider': 'google',
            'search_url': self.search_url, 'kind': 'organic',
        }])

    def test_search_facts_survive_unavailable_destination_pages(self):
        self.service.fetch = Mock(return_value=None)
        result = self.service.enrich(self.record)
        self.assertEqual([result[name] for name in ('Revenue', 'SIC Code', 'NAICS Code')],
                         ['$18 million', '0111', '111110'])
        self.assertEqual(result['Revenue Link'], self.search_url)
        self.assertEqual(result['SIC/NAICS Code Link'], self.search_url)
        self.assertEqual(len(result['_company_evidence']), 3)
        for item in result['_company_evidence']:
            self.assertIn(self.result_url, item['excerpt'])
            self.assertIn('search', item['basis'].lower())

    def test_fetched_source_overrides_snippet_in_same_enrichment(self):
        document = SourceDocument('https://example.com/results', 'Example Industries',
                                  'Example Industries\n2024 annual revenue: $12 million\n'
                                  'SIC Code: 1234\nNAICS Code: 541512')
        self.service.fetch = Mock(return_value=document)
        result = self.service.enrich(self.record)
        self.assertEqual([result[name] for name in ('Revenue', 'SIC Code', 'NAICS Code')],
                         ['$12 million', '1234', '541512'])
        self.assertEqual(result['Revenue Link'], document.url)
        self.assertEqual(result['SIC/NAICS Code Link'], document.url)

    def test_cached_search_results_do_not_repeat_queries_or_drop_evidence(self):
        self.service.fetch = Mock(return_value=None)
        first = self.service.enrich(self.record)
        search_count = self.service.search.call_count
        self.assertGreater(search_count, 0)
        second = self.service.enrich(self.record)
        self.assertEqual(self.service.search.call_count, search_count)
        self.assertEqual(first, second)
        self.assertEqual(len(second['_company_evidence']), 3)


if __name__ == '__main__':
    unittest.main()
