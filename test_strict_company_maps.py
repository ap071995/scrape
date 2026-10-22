"""Adversarial company/Maps evidence tests; all inputs and browser APIs are offline."""
import copy
import unittest
from unittest.mock import Mock

from test_scraper import app
from company_public_data import (
    PublicCompanyEnricher, SourceDocument, parse_document, parse_search_result,
    public_url, unwrap_search_url,
)


class StrictCompanyEvidenceTests(unittest.TestCase):
    record = {'Company Name': 'Example Industries', 'Company Website': 'https://example.com/'}
    match = staticmethod(app.SearchWorker._link_company_names_match)

    def parse(self, text, *, title='Example Industries', headings=None, tables=None,
              url='https://example.com/results'):
        return parse_document(SourceDocument(url, title, text, headings or [], [], tables or []),
                              self.record, self.match)

    def revenue(self, facts):
        return [item for item in facts if item['field'] == 'Revenue']

    def test_subsidiary_heading_overrides_parent_brand_without_title_separator(self):
        facts = self.parse('SubsidiaryCo\nA business of Example Industries\nAnnual revenue: $8 million\n'
                           'NAICS Code: 541512', headings=['SubsidiaryCo'])
        self.assertEqual(facts, [])

    def test_explicit_other_company_name_overrides_matching_directory_title(self):
        self.assertEqual(self.parse('Company Name: Other Corporation\nAnnual revenue: $8 million\n'
                                    'NAICS Code: 541512', url='https://naics.com/company-profile-page/?co=8'), [])

    def test_narrative_company_subject_is_not_borrowed_from_parent_page(self):
        self.assertEqual(self.parse('Example Industries\nAnnual results\n'
                                    'Other Corporation reported annual revenue of $8 million in 2025.'), [])
        facts = self.parse('Example Industries\nAnnual results\n'
                           'Example Industries reported annual revenue of $8 million in 2025.')
        self.assertEqual([f['value'] for f in facts], ['$8 million'])

    def test_financial_plan_section_rejects_table_and_snapshot_amounts(self):
        for heading in ('Annual budget 2025', 'Revenue forecast 2025', 'Financial projections 2025'):
            with self.subTest(heading=heading):
                body = heading + '\nTotal Revenue $7 million $8 million\nSIC Code: 7371'
                facts = self.parse(body, tables=[[['', '2024', '2025'],
                                                   ['Total Revenue', '$7 million', '$8 million']]])
                self.assertEqual(self.revenue(facts), [])
                self.assertEqual([f['value'] for f in facts], ['7371'])
                snapshot = self.parse('2025 Financial Results\n' + heading + '\n$8\nSales (millions)')
                self.assertEqual(self.revenue(snapshot), [])

    def test_split_guidance_section_is_not_reported_annual_revenue(self):
        self.assertEqual(self.revenue(self.parse('Example Industries\n2025 Financial Results\n'
                                                'Full-year guidance\nAnnual revenue: $8 million')), [])

    def test_partial_year_period_on_revenue_line_overrides_annual_page_heading(self):
        for period in ('H1 2025', 'half-year 2025', 'year-to-date 2025', 'six-month period 2025'):
            with self.subTest(period=period):
                facts = self.parse('Annual results\n' + period + ' total revenue was $8 million')
                self.assertEqual(self.revenue(facts), [])

    def test_forecast_header_does_not_contaminate_later_actual_annual_section(self):
        facts = self.parse('Outlook\nAnnual revenue forecast: $9 million\n'
                           '2025 Annual Results\nTotal revenue: $8 million')
        self.assertEqual([f['value'] for f in self.revenue(facts)], ['$8 million'])

    def test_partial_and_range_classification_values_are_not_single_codes(self):
        for value in ('541511-541519', '541512.0', '541512/541519', '541512 or 541519'):
            with self.subTest(value=value):
                self.assertEqual(self.parse('NAICS Code: ' + value), [])

    def test_unrelated_named_classification_cannot_be_borrowed(self):
        self.assertEqual(self.parse('Other Corporation NAICS Code: 541512\n'
                                    'Other Corporation SIC Code: 7371'), [])

    def test_valid_primary_classifications_keep_leading_zero_and_provenance(self):
        facts = self.parse('Company Name: Example Industries\nSIC Code: 0111\nNAICS Code: 111110')
        self.assertEqual({f['field']: f['value'] for f in facts},
                         {'SIC Code': '0111', 'NAICS Code': '111110'})
        self.assertTrue(all(f['url'] == 'https://example.com/results' for f in facts))

    def test_legacy_private_ip_spellings_and_non_web_ports_are_not_public(self):
        for url in ('http://127.1/facts', 'http://127.0.1/facts', 'http://0x7f.0.0.1/facts',
                    'http://0177.0.0.1/facts', 'http://LOCALHOST./facts', 'http://local.internal./facts',
                    'https://8.8.8.8:8443/facts'):
            with self.subTest(url=url):
                self.assertFalse(public_url(url))
        self.assertTrue(public_url('https://investors.example.com:443/results'))

    def test_malformed_search_destination_returns_no_facts_instead_of_raising(self):
        row = {'title': 'Example Industries - Revenue', 'text': 'Annual revenue: $8 million',
               'url': 'http://[malformed', 'provider': 'google', 'kind': 'organic',
               'search_url': 'https://www.google.com/search?q=Example+Industries'}
        self.assertEqual(parse_search_result(row, self.record, self.match), [])

    def test_search_redirect_host_suffix_is_not_search_provider(self):
        url = 'https://notgoogle.com/url?q=https%3A%2F%2Fexample.com%2Fresults'
        self.assertEqual(unwrap_search_url(url), url)

    def test_unsafe_redirect_cannot_provide_company_evidence(self):
        page = Mock()
        page.goto.return_value.status = 200
        page.url = 'http://127.0.0.1/company'
        page.evaluate.return_value = {'title': 'Example Industries', 'text': 'Annual revenue: $8 million',
                                      'headings': [], 'links': [], 'tables': []}
        service = PublicCompanyEnricher(Mock(), self.match)
        service._page = Mock(return_value=page)
        self.addCleanup(service.close)
        self.assertIsNone(service.fetch('https://example.com/redirect'))
        page.evaluate.assert_not_called()


class StrictMapsEvidenceTests(unittest.TestCase):
    def setUp(self):
        self.worker = app.SearchWorker({})

    def test_matching_name_cannot_override_explicit_conflicting_website(self):
        self.assertFalse(self.worker._maps_result_matches_company(
            'Example Industries', 'Example Industries', 'https://example.com/', 'https://other.example/'))

    def test_parent_domain_cannot_validate_named_subsidiary(self):
        self.assertFalse(self.worker._maps_result_matches_company(
            'Example Industries', 'Example Industries Ventures',
            'https://example.com/', 'https://example.com/ventures/'))

    def test_valid_same_company_office_allows_missing_or_regional_website(self):
        for website in ('-', 'https://us.example.com/locations/boston'):
            with self.subTest(website=website):
                self.assertTrue(self.worker._maps_result_matches_company(
                    'Example Industries', 'Example Industries Corporate Office',
                    'https://example.com/', website))

    def test_cached_maps_source_link_must_be_safe_and_match_its_phone(self):
        record = {'Address': '123 Main St, Boston, MA 02110, USA'}
        proposed = {'Address': record['Address'], 'Address Link': 'javascript:alert(1)',
                    'Phone Number': '+16175551234', 'Phone Number Link': 'tel:+12125559999'}
        self.worker._merge_google_maps_fields(record, proposed)
        self.assertNotIn('Address Link', record)
        self.assertNotIn('Phone Number Link', record)
        self.assertEqual(record['Phone Number'], '+16175551234')

    def test_locality_only_address_cannot_upgrade_to_wrong_city_office(self):
        record = {'City': 'Boston', 'State': 'Massachusetts', 'Country': 'United States',
                  'Address': 'Boston, MA'}
        before = copy.deepcopy(record)
        applied = self.worker._merge_google_maps_fields(record, {
            'Address': '23 Boston Street, Springfield, MA 01103, USA',
            'Address Link': 'https://www.google.com/maps/place/Example',
            'Phone Number': '+14135551234', 'Zip Code': '01103'})
        self.assertEqual(applied, {})
        self.assertEqual(record, before)

    def test_existing_suite_cannot_receive_source_for_explicit_different_suite(self):
        record = {'Address': '123 Main St, Suite 500, Boston, MA 02110, USA'}
        before = copy.deepcopy(record)
        self.worker._merge_google_maps_fields(record, {
            'Address': '123 Main St, Suite 600, Boston, MA 02110, USA',
            'Address Link': 'https://www.google.com/maps/place/Example',
            'Phone Number': '+16175551234', 'Phone Number Link': 'tel:+16175551234'})
        self.assertEqual(record, before)


if __name__ == '__main__':
    unittest.main(verbosity=2)
