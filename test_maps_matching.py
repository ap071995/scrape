"""False-rejection regressions for Google Maps identity, geography and merges."""
import unittest
from test_scraper import app


class MapsMatchingTests(unittest.TestCase):
    def setUp(self):
        self.worker = app.SearchWorker({})

    def location(self, city='Boston', state='Massachusetts', country='United States', **extra):
        return {'Company Name': 'Example Industries', 'City': city, 'State': state,
                'Country': country, **extra}

    def test_country_omitted_from_matching_us_address(self):
        self.assertTrue(self.worker._maps_address_matches_location(self.location(),
            '123 Main St, Boston, MA 02110'))

    def test_country_omitted_from_matching_indian_address(self):
        self.assertTrue(self.worker._maps_address_matches_location(
            self.location('Bangalore', 'Karnataka', 'India'),
            '23 Main Road, Bengaluru, Karnataka 560040'))

    def test_district_display_variant(self):
        self.assertTrue(self.worker._maps_address_matches_location(
            self.location('Pune District', 'Maharashtra', 'India'),
            'Hinjewadi Phase 3, Pune, Maharashtra 411057, India'))

    def test_saint_city_abbreviation(self):
        self.assertTrue(self.worker._maps_address_matches_location(
            self.location('Saint Paul', 'Minnesota'),
            '3M Center, St Paul, MN 55144, USA'))

    def test_unknown_city_markers_fall_back_to_state(self):
        for missing in ('-', 'N/A', 'not available', 'unknown', None):
            with self.subTest(missing=missing):
                self.assertTrue(self.worker._maps_address_matches_location(
                    self.location(missing), '123 Main St, Boston, MA 02110, USA'))

    def test_country_aliases_are_bidirectional(self):
        for expected, displayed in [('UK', 'United Kingdom'), ('USA', 'United States'),
                                    ('UAE', 'United Arab Emirates'), ('CH', 'Switzerland')]:
            with self.subTest(expected=expected):
                self.assertTrue(self.worker._maps_address_matches_country(
                    self.location(country=expected), '123 Main Road, Example City, ' + displayed))

    def test_country_with_postal_code_in_same_component(self):
        self.assertTrue(self.worker._maps_address_matches_location(
            self.location('Singapore', '-', 'Singapore'),
            '160 Robinson Rd, #14-04, Singapore 068914'))

    def test_explicit_different_country_is_rejected(self):
        for address in ('123 Main St, Boston, MA 02110, Canada',
                        '123 Main St, Boston, MA 02110 Canada'):
            self.assertFalse(self.worker._maps_address_matches_location(self.location(), address))

    def test_country_alone_does_not_validate_omitted_country(self):
        self.assertFalse(self.worker._maps_address_matches_location(
            self.location('-', '-', 'India'), '123 Main St, Boston, MA 02110'))

    def test_wrong_city_is_rejected(self):
        self.assertFalse(self.worker._maps_address_matches_location(self.location(),
            '100 Broadway, New York, NY 10001, USA'))

    def test_city_name_in_street_is_not_locality_evidence(self):
        self.assertFalse(self.worker._maps_address_matches_location(
            self.location('Paris', '-', 'France'), '123 Paris Rd, Lyon 69001, France'))

    def test_wrong_state_with_same_city_is_rejected(self):
        self.assertFalse(self.worker._maps_address_matches_location(
            self.location('Springfield', 'Illinois'), '123 Main St, Springfield, MA 01103, USA'))

    def test_wrong_indian_state_is_rejected(self):
        self.assertFalse(self.worker._maps_address_matches_location(
            self.location('Bengaluru', 'Karnataka', 'India'),
            '123 Main Rd, Bengaluru, Maharashtra 560040, India'))

    def test_ecu_address_abbreviations_are_equivalent(self):
        self.assertTrue(self.worker._maps_addresses_equivalent(
            'East 5th Street, Greenville, NC 27858-4353, US',
            'E 5th St, Greenville, NC 27858, United States'))

    def test_omitted_country_in_same_office_is_compatible(self):
        self.assertTrue(self.worker._maps_address_compatible(
            '123 Main Street, Boston, MA 02110, USA', '123 Main St, Boston, MA 02110'))

    def test_partial_street_matches_same_office(self):
        self.assertTrue(self.worker._maps_address_compatible(
            '123 Main Street', '123 Main St, Boston, MA 02110, USA'))

    def test_city_only_old_address_allows_complete_maps_address(self):
        self.assertTrue(self.worker._maps_address_compatible(
            'St Paul, MN', '3M Center, Saint Paul, MN 55144, United States'))

    def test_city_only_address_is_upgraded_with_its_source(self):
        record = self.location('St Paul', 'Minnesota', Address='St Paul, MN')
        proposed = {'Address':'2510 Conway Ave, St Paul, MN 55144, United States',
                    'Address Link':'https://www.google.com/maps/place/3M', 'Zip Code':'55144'}
        self.assertTrue(self.worker._google_maps_needs_enrichment(record))
        self.worker._merge_google_maps_fields(record, proposed)
        self.assertEqual(record['Address'], proposed['Address'])
        self.assertEqual(record['Address Link'], proposed['Address Link'])
        self.assertEqual(record['Zip Code'], '55144')

    def test_partial_city_is_not_equivalent_to_complete_office(self):
        self.assertFalse(self.worker._maps_addresses_equivalent(
            'St Paul, MN', '3M Center, Saint Paul, MN 55144, United States'))

    def test_different_street_or_unit_is_rejected(self):
        for other in ('125 Main St, Boston, MA 02110, USA',
                      '123 Other St, Boston, MA 02110, USA',
                      '123 Main St, Suite 200, Boston, MA 02110, USA'):
            self.assertFalse(self.worker._maps_address_compatible(
                '123 Main Street, Suite 100, Boston, MA 02110, USA', other))

    def test_existing_hq_address_takes_precedence_over_prospect_city(self):
        record = self.location('London', '-', 'United Kingdom',
                               Address='Postbox, Zurich, 8001, CH')
        self.assertTrue(self.worker._maps_address_matches_location(record,
            'Postbox, Zurich, 8001, Switzerland'))
        self.assertFalse(self.worker._maps_address_matches_location(record,
            '5 Broadgate, London EC2M 2QS, United Kingdom'))

    def test_section_labels_do_not_block_same_street(self):
        old = 'Corporate Office\n28411 Northwestern Highway, Ste 400, Southfield, Michigan 48034, US'
        new = '28411 Northwestern Hwy, Suite 400, Southfield, MI 48034, United States'
        self.assertTrue(self.worker._maps_addresses_equivalent(old, new))
        query = self.worker._google_maps_queries(self.location(Address=old))[0]['query']
        self.assertNotIn('Corporate+Office', query)

    def test_omitted_suite_is_compatible_but_not_claimed_as_verified_suite(self):
        old = 'Corporate Office\n28411 Northwestern Highway, Ste 400, Southfield, Michigan 48034, US'
        new = '28411 Northwestern Hwy, Southfield, MI 48034, United States'
        self.assertTrue(self.worker._maps_address_compatible(old, new))
        self.assertFalse(self.worker._maps_addresses_equivalent(old, new))
        self.assertFalse(self.worker._maps_address_compatible(old, new.replace('Hwy,', 'Hwy, Suite 500,')))

    def test_existing_national_phone_can_receive_international_source_link(self):
        record = self.location(Address='123 Main St, Boston, MA, US', **{'Phone Number':'6175551234'})
        self.worker._merge_google_maps_fields(record, {'Address':record['Address'],
            'Phone Number':'+16175551234', 'Phone Number Link':'tel:+16175551234'})
        self.assertEqual(record['Phone Number'], '6175551234')
        self.assertEqual(record['Phone Number Link'], 'tel:+16175551234')
        self.assertFalse(self.worker._maps_phones_equivalent('6175551234', '+446175551234', record))

    def test_unnumbered_postbox_uses_known_city_and_postal_code(self):
        record = self.location('London', '-', 'United Kingdom', Address='Postbox, Zurich, 8001, CH')
        self.assertTrue(self.worker._maps_address_matches_location(record,
            'Bahnhofstrasse 45, 8001 Zurich, Switzerland'))
        self.assertFalse(self.worker._maps_address_matches_location(record,
            'Bahnhofstrasse 45, 8002 Zurich, Switzerland'))

    def test_numbered_postboxes_are_preserved(self):
        self.assertFalse(self.worker._maps_address_compatible(
            'PO Box 123, Zurich, 8001, CH', 'PO Box 456, Zurich, 8001, CH'))

    def test_old_street_only_still_requires_expected_city(self):
        self.assertFalse(self.worker._maps_address_matches_location(
            self.location(Address='123 Main St'), '123 Main St, New York, NY 10001, USA'))

    def test_formatted_same_office_receives_missing_contact_details(self):
        old = 'East 5th Street, Greenville, NC 27858-4353, US'
        record = {'Address': old}
        proposed = {'Address': 'E 5th St, Greenville, NC 27858, United States',
                    'Zip Code': '27858', 'Phone Number': '+12523286131',
                    'Phone Number Link': 'tel:+12523286131',
                    'Address Link': 'https://www.google.com/maps/place/ECU'}
        applied = self.worker._merge_google_maps_fields(record, proposed)
        self.assertEqual(record['Address'], old)
        self.assertEqual(record['Phone Number'], '+12523286131')
        self.assertEqual(set(applied), set(proposed) - {'Address'})

    def test_different_office_cannot_supply_contact_fields(self):
        record = {'Address': '123 Main St, Boston, MA 02110, USA'}
        self.assertEqual(self.worker._merge_google_maps_fields(record, {
            'Address': '100 Broadway, New York, NY 10001, USA', 'Zip Code': '10001',
            'Phone Number': '+12125551234', 'Phone Number Link': 'tel:+12125551234'}), {})

    def test_conflicting_existing_phone_keeps_its_own_source(self):
        record = {'Address': '123 Main St', 'Phone Number': '+16171111111'}
        self.worker._merge_google_maps_fields(record, {'Address': '123 Main St, Boston, MA',
            'Phone Number': '+16172222222', 'Phone Number Link': 'tel:+16172222222'})
        self.assertEqual(record['Phone Number'], '+16171111111')
        self.assertNotIn('Phone Number Link', record)

    def test_fuzzy_name_can_be_selected_without_exact_token_overlap(self):
        self.assertGreater(self.worker._maps_company_match_score('Acme Industries', 'Acme Industires'), 0)

    def test_corporate_office_descriptor_can_be_selected(self):
        self.assertGreater(self.worker._maps_company_match_score('3M', '3M Corporate Headquarters'), 0)

    def test_different_subsidiary_is_not_an_office_descriptor(self):
        for expected, actual in [('3M', '3M Competitor'), ('Cognizant', 'Cognizant Moment'),
                                 ('Anyscale', 'Scale AI')]:
            self.assertEqual(self.worker._maps_company_match_score(expected, actual), 0)

    def test_shared_city_website_does_not_validate_water_department_phone(self):
        self.assertFalse(self.worker._maps_result_matches_company('City of Glendale AZ',
            'Glendale Water Department', 'https://glendaleaz.com/', 'https://www.glendaleaz.com/water/'))
        self.assertTrue(self.worker._maps_result_matches_company('City of Glendale AZ',
            'Glendale City Hall', 'https://glendaleaz.com/', 'https://www.glendaleaz.com/'))

    def test_university_bookstore_is_not_general_university_contact(self):
        self.assertFalse(self.worker._maps_result_matches_company('East Carolina University',
            'East Carolina University Bookstore', 'https://ecu.edu/', 'https://ecu.edu/bookstore/'))
        self.assertGreater(self.worker._maps_company_match_score('Glendale Water Department',
            'Glendale Water Department'), 0)

    def test_different_non_latin_names_do_not_match_as_empty_ascii(self):
        self.assertEqual(self.worker._maps_company_match_score('東京会社', '大阪企業'), 0)

    def test_queries_include_disambiguating_state_and_country(self):
        query = self.worker._google_maps_queries(self.location('Paris', 'Texas'))[0]['query']
        self.assertIn('Paris+Texas+United+States', query)

    def test_missing_city_keeps_state_country_and_website_fallback(self):
        queries = self.worker._google_maps_queries(self.location('-', Company_Website='unused',
            **{'Company Website': 'https://example.com'}))
        self.assertEqual(len(queries), 2)
        self.assertIn('Massachusetts+United+States', queries[0]['query'])
        self.assertIn('example.com+', queries[1]['query'])

    def test_existing_office_address_is_used_for_query(self):
        query = self.worker._google_maps_queries(self.location('London', '-', 'United Kingdom',
            Address='Postbox, Zurich, 8001, CH'))[0]['query']
        self.assertIn('Zurich', query)
        self.assertNotIn('London', query)

    def test_cache_separates_offices_and_preserves_location_slots(self):
        first = self.location(Address='123 Main St')
        second = self.location(Address='125 Main St')
        self.assertNotEqual(self.worker._google_maps_cache_key(first), self.worker._google_maps_cache_key(second))
        self.assertNotEqual(self.worker._google_maps_cache_key(self.location('Georgia', '-', '-')),
                            self.worker._google_maps_cache_key(self.location('-', 'Georgia', '-')))


if __name__ == '__main__':
    unittest.main(verbosity=2)
