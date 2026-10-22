import json
import tempfile
import threading
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from test_scraper import app
from company_public_data import (PublicCompanyEnricher, SourceDocument, parse_document,
                                 select_facts, unwrap_search_url, public_url)


class PublicDataParsingTests(unittest.TestCase):
    record = {'Company Name': 'Example Industries', 'Company Website': 'https://example.com/'}

    def parse(self, body, title='Example Industries', url='https://example.com/annual-results'):
        return parse_document(SourceDocument(url, title, body, [title]), self.record,
                              app.SearchWorker._link_company_names_match)

    def test_annual_amount_period_and_link(self):
        facts = self.parse('Example Industries\n2025 Financial Results\n$24.9\nSales (billions)')
        self.assertEqual(facts[0]['value'], '$24.9 billion')
        self.assertEqual(facts[0]['period'], '2025')
        self.assertIn('Reported annual', facts[0]['basis'])

    def test_quarter_budget_forecast_and_profit_rejected(self):
        for line in ['Q2 2025 revenue was $3 billion', '2025 annual revenue budget $3 billion',
                     '2025 annual revenue forecast $3 billion', 'Annual net income $3 billion',
                     '2025 adjusted annual revenue $3 billion']:
            self.assertEqual(self.parse(line), [], line)

    def test_separate_quarter_header_is_not_annual(self):
        for body in ['Quarterly Results Q1 2025\n(Unaudited)\nRevenue: $12 million\nAnnual revenue growth was 8%',
                     '2025 Financial Results\nFirst Quarter\n$3\nSales (billions)']:
            self.assertEqual(self.parse(body), [])

    def test_comparative_years_do_not_mix_amount_and_period(self):
        self.assertEqual(self.parse('Annual revenue was $10 million in 2024 and $12 million in 2025'), [])

    def test_subsidiary_revenue_and_future_expectations_rejected(self):
        for body in ['Annual results\nRevenue for Acme Subsidiary: $4 million',
                     'The company expects annual revenue of $50 million in 2027.']:
            self.assertEqual(self.parse(body), [])

    def test_other_company_section_is_excluded(self):
        self.assertEqual(self.parse('Company profile\nOther companies\nAcme Subsidiary\nSIC Code: 1234\nNAICS Code: 123456'), [])

    def test_revenue_per_employee_is_not_company_revenue(self):
        self.assertEqual(self.parse('Annual results\nRevenue per employee $7,000'), [])

    def test_later_tax_amount_is_not_revenue(self):
        self.assertEqual(self.parse('Annual results\nTaxes are a contribution to fiscal revenue. '
                                   'Example Industries and its employees paid around CHF 24 billion in taxes.'), [])

    def test_annual_table_preserves_year_alignment(self):
        doc = SourceDocument('https://example.com/revenue', 'Example Industries', 'Total Revenues',
             tables=[[['', '2023', '2024'], ['Total Revenues', '$10,000', '$12,000']]])
        facts = parse_document(doc, self.record, app.SearchWorker._link_company_names_match)
        self.assertEqual(facts[0]['value'], '$12,000')
        self.assertEqual(facts[0]['period'], '2024')

    def test_quarterly_page_title_blocks_all_revenue_paths(self):
        for body, tables in [
                ('First-quarter revenue', [[['', '2024', '2025'], ['Total Revenue', '$10 million', '$12 million']]]),
                ('2025 Financial Results\n$12\nSales (billions)', []),
                ('Annual results\n2025 Total revenue: $12 million', [])]:
            doc = SourceDocument('https://example.com/q1', 'Example Industries Q1 2025 results',
                                 body + '\nSIC Code: 1234', tables=tables)
            facts = parse_document(doc, self.record, app.SearchWorker._link_company_names_match)
            self.assertEqual([fact['field'] for fact in facts], ['SIC Code'])

    def test_quarterly_table_body_header_is_not_annual(self):
        for body, header in [('Three months ended June 30', ''), ('', 'First-quarter results')]:
            doc = SourceDocument('https://example.com/results', 'Example Industries Financial Results', body,
                 tables=[[[header, '2024', '2025'], ['Total Revenue', '$10 million', '$12 million']]])
            self.assertEqual(parse_document(doc, self.record, app.SearchWorker._link_company_names_match), [])

    def test_standalone_revenue_does_not_borrow_another_metric(self):
        for body in ['Annual results\nRevenue\nNot disclosed\nTotal assets: $10 billion',
                     'Annual results\nRevenue\nTotal assets: $10 billion',
                     'Annual results\nRevenue\n$10 billion of assets']:
            self.assertEqual(self.parse(body), [], body)

    def test_standalone_revenue_amount_is_still_accepted(self):
        for source, expected in [('$12 million', '$12 million'), ('$5 - 10 million', '$5 million - $ 10 million')]:
            facts = self.parse('Annual results\nRevenue\n' + source)
            self.assertEqual(facts[0]['value'], expected)

    def test_official_acronym_heading(self):
        doc = SourceDocument('https://budget.ecu.edu/revenue', 'Revenue | ECU', 'Total Revenues',
             tables=[[['', '2023', '2024'], ['Total Revenues', '$10,000', '$12,000']]])
        facts = parse_document(doc, {'Company Name':'East Carolina University', 'Company Website':'https://ecu.edu/'},
                               app.SearchWorker._link_company_names_match)
        self.assertEqual(facts[0]['value'], '$12,000')

    def test_estimate_range_is_not_reported_fact(self):
        facts = self.parse("Example Industries\nWebsite\nexample.com\nExample Industries's revenue is estimated to be in the range of $1B$10B",
                           url='https://leadiq.com/c/example/123')
        self.assertEqual(facts[0]['value'], '$1B - $10B')
        self.assertNotIn('Reported annual', facts[0]['basis'])

    def test_wrong_company_or_website_rejected(self):
        for title, body in [('Example Industries International', 'Revenue $7M'),
                            ('Example Industries', 'Website: other.example\nRevenue: $7M')]:
            self.assertEqual(self.parse(body, title, 'https://naics.com/company-profile-page/?co=4'), [])

    def test_unknown_website_does_not_mean_a_conflicting_website(self):
        document = SourceDocument('https://naics.com/company-profile-page/?co=4', 'Example Industries',
                                  'Website: example.com\nNAICS Code: 111110')
        for website in ('-', 'NA', 'Not found'):
            record = dict(self.record, **{'Company Website':website})
            facts = parse_document(document, record, app.SearchWorker._link_company_names_match)
            self.assertEqual(facts[0]['value'], '111110')

    def test_official_parent_domain_does_not_validate_subsidiary_page(self):
        for headings in (['SubsidiaryCo'], []):
            document = SourceDocument('https://example.com/subsidiary', 'SubsidiaryCo | Example Industries',
                                      'SubsidiaryCo\nAnnual revenue: $3 billion\nNAICS Code: 111110', headings=headings)
            self.assertEqual(parse_document(document, self.record, app.SearchWorker._link_company_names_match), [])

    def test_generic_official_page_heading_still_validates_company(self):
        for heading in ('Investor Relations', 'Financial Results', '2025 Annual Results', 'Revenue'):
            document = SourceDocument('https://example.com/results', heading + ' | Example Industries',
                                      heading + '\nAnnual revenue: $3 billion', headings=[heading])
            facts = parse_document(document, self.record, app.SearchWorker._link_company_names_match)
            self.assertEqual(facts[0]['value'], '$3 billion')

    def test_specific_codes_and_commercial_sic_extension(self):
        facts = self.parse('SIC 1: 01110000\nNAICS 1: 111110\nNAICS 2: 111120')
        values = {x['field']:x['value'] for x in facts}
        self.assertEqual(values, {'SIC Code':'0111', 'NAICS Code':'111110'})

    def test_broad_sector_is_not_padded_to_six_digits(self):
        self.assertEqual(self.parse('NAICS Code: 531\nSIC Code: 65'), [])

    def test_competitor_code_is_not_company_code(self):
        self.assertEqual(self.parse('Competitor and target customer NAICS Code: 541512'), [])

    def test_conflicting_codes_on_same_page_rejected(self):
        self.assertEqual(self.parse('NAICS Code: 541512\nNAICS Code: 541511'), [])

    def test_primary_annual_result_beats_undated_directory(self):
        direct = self.parse('2025 Financial Results\n$24.9\nSales (billions)')
        directory = self.parse('Sales Volume: $35,355,000,000', url='https://naics.com/company-profile-page/?co=4')
        self.assertEqual(select_facts(direct+directory)[0]['value'], '$24.9 billion')

    def test_equal_rank_conflicts_are_unresolved(self):
        one = self.parse('2025 Financial Results\n$24.9\nSales (billions)')
        two = self.parse('2025 Financial Results\n$35.0\nSales (billions)')
        self.assertEqual(select_facts(one+two), [])

    def test_no_local_source_navigation(self):
        for url in ['file:///C:/secret','http://127.0.0.1/data','http://192.168.1.1','http://localhost','http://host.internal/']:
            self.assertFalse(public_url(url))
        self.assertTrue(public_url('https://example.com/annual-results'))

    def test_google_redirect_unwrapped(self):
        self.assertEqual(unwrap_search_url('https://www.google.com/url?q=https%3A%2F%2Fexample.com%2Ffacts'),
                         'https://example.com/facts')


class PublicDataEnrichmentTests(unittest.TestCase):
    def setUp(self):
        self.service = PublicCompanyEnricher(Mock(), app.SearchWorker._link_company_names_match)
        self.service._discover = Mock(return_value=['https://example.com/facts'])
        self.service.fetch = Mock(return_value=SourceDocument('https://example.com/facts', 'Example Industries',
            'Example Industries\n2025 Financial Results\n$24.9\nSales (billions)\nSIC Code: 0111\nNAICS Code: 111110'))
        self.record = {'Company Name':'Example Industries', 'Company Website':'https://example.com/'}

    def test_pairs_evidence_and_company_cache(self):
        first = self.service.enrich(self.record)
        second = self.service.enrich(dict(self.record, City='Another city'))
        self.assertEqual(first['Revenue'], '$24.9 billion')
        self.assertEqual(first['Revenue Link'], 'https://example.com/facts')
        self.assertEqual(len(first['_company_evidence']), 3)
        self.assertEqual(second['SIC Code'], '0111')
        self.service.fetch.assert_called_once()

    def test_different_existing_amount_does_not_receive_false_source(self):
        result = self.service.enrich(dict(self.record, Revenue='$10 million'))
        self.assertEqual(result['Revenue'], '$10 million')
        self.assertNotIn('Revenue Link', result)
        self.assertFalse(any(x['field']=='Revenue' for x in result['_company_evidence']))

    def test_missing_sentinels_and_invalid_existing_values_are_repaired(self):
        for revenue, sic, naics in [('NA', 'Not found', 'null'), ('not revenue', '011', '111')]:
            result = self.service.enrich(dict(self.record, Revenue=revenue,
                                             **{'SIC Code':sic, 'NAICS Code':naics}))
            self.assertEqual(result['Revenue'], '$24.9 billion')
            self.assertEqual(result['SIC Code'], '0111')
            self.assertEqual(result['NAICS Code'], '111110')
            self.assertEqual(len(result['_company_evidence']), 3)

    def test_existing_structurally_valid_conflicting_facts_are_preserved(self):
        for revenue in ('$10 million', '10 million', 'CHF 10 million', '$5 - 10 million'):
            original = dict(self.record, Revenue=revenue,
                            **{'SIC Code':'1234', 'NAICS Code':'123456',
                               'Revenue Link':'https://example.com/old-revenue',
                               'SIC/NAICS Code Link':'https://example.com/old-codes'})
            result = self.service.enrich(original)
            for field in ('Revenue', 'SIC Code', 'NAICS Code', 'Revenue Link', 'SIC/NAICS Code Link'):
                self.assertEqual(result[field], original[field])
            self.assertEqual(result['_company_evidence'], [])

    def test_failed_sources_are_cached_for_run(self):
        self.service.fetch.return_value = None
        self.assertNotIn('Revenue', self.service.enrich(self.record))
        self.service.enrich(self.record)
        self.service.fetch.assert_called_once()

    def test_failed_navigation_does_not_poison_next_source_tab(self):
        failed, healthy = Mock(), Mock()
        failed.goto.side_effect = RuntimeError('Navigation interrupted by chrome-error page')
        healthy.goto.return_value.status = 200
        healthy.url = 'https://example.com/facts'
        healthy.evaluate.return_value = {'title':'Example Industries', 'text':'Annual revenue: $3 million',
                                         'headings':[], 'links':[], 'tables':[]}
        context = Mock()
        context.new_page.side_effect = [failed, healthy]
        service = PublicCompanyEnricher(context, app.SearchWorker._link_company_names_match)
        self.assertIsNone(service.fetch('https://example.com/broken'))
        failed.close.assert_called_once()
        self.assertEqual(service.fetch('https://example.com/facts').title, 'Example Industries')
        self.assertEqual(context.new_page.call_count, 2)
        service.close()

    def test_stop_prevents_search(self):
        self.service.stop_event = threading.Event()
        self.service.stop_event.set()
        self.assertEqual(self.service.enrich(self.record), self.record)
        self.service._discover.assert_not_called()

    def test_configured_search_api(self):
        response = Mock()
        response.json.return_value = {'web':{'results':[{'title':'Example Industries revenue',
            'url':'https://example.com/facts','description':'Annual results'}]}}
        with patch.dict('os.environ', {'BRAVE_SEARCH_API_KEY':'test-token'}), \
             patch('company_public_data.requests.get', return_value=response) as get:
            rows = self.service.search('Example Industries revenue')
        self.assertEqual(rows[0]['url'], 'https://example.com/facts')
        self.assertEqual(get.call_args.kwargs['params']['q'], 'Example Industries revenue')

    def test_blocked_google_is_disabled_for_rest_of_run(self):
        page = Mock()
        page.url = 'https://www.google.com/search?q=test'
        page.locator.return_value.inner_text.side_effect = ['unusual traffic', 'Bing results', 'Bing results']
        page.evaluate.return_value = [{'title':'Example Industries Revenue', 'url':'https://example.com/facts', 'text':'annual revenue'}]
        self.service._page = Mock(return_value=page)
        with patch.dict('os.environ', {'BRAVE_SEARCH_API_KEY':''}):
            first = self.service.search('Example Industries revenue')
            self.service.search('Example Industries SIC code')
        self.assertEqual(first[0]['url'], 'https://example.com/facts')
        self.assertIn('google', self.service.disabled_providers)
        self.assertEqual(sum('google.com' in call.args[0] for call in page.goto.call_args_list), 1)


if __name__ == '__main__':
    unittest.main(verbosity=2)
