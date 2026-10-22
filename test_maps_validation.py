"""Regressions for contact-data conflicts found during live verification."""
import unittest
from unittest.mock import Mock
from test_scraper import app


class MapsValidationTests(unittest.TestCase):
    def setUp(self):
        self.worker = app.SearchWorker({})

    def test_fresh_and_cached_details_do_not_mix_offices_or_phone_sources(self):
        details = {'company_matched': True, 'address': '100 Broadway, New York, NY 10001, USA',
                   'zip_code': '10001', 'address_link': 'https://maps.google.com/new-york',
                   'phone': '+12125551234', 'phone_link': 'tel:+12125551234'}
        self.worker._get_google_maps_company_details = Mock(return_value=details)
        original = {'Company Name': 'Example', 'City': 'Boston', 'Country': 'United States',
                    'Address': '123 Main St, Boston, MA, USA', 'Phone Number': '+16175551234'}
        for attempt in range(2):
            record = self.worker._google_maps_enrich_record(dict(original), 'test')
            self.assertEqual(record['Address'], original['Address'])
            self.assertEqual(record['Phone Number'], original['Phone Number'])
            for field in ['Zip Code', 'Address Link', 'Phone Number Link']:
                self.assertEqual(record.get(field, '-'), '-', (attempt, field))
        self.assertEqual(self.worker._get_google_maps_company_details.call_count, 1)

    def configure_place(self, address, phone):
        page = Mock()
        page.url = 'https://www.google.com/maps/place/Example'
        self.worker._get_google_maps_page = Mock(return_value=page)
        self.worker._google_maps_queries = Mock(return_value=[{'query': 'Example Boston'}])
        self.worker._dismiss_google_consent = Mock()
        self.worker._wait_and_scroll_google_maps_place = Mock()
        self.worker._extract_google_maps_place = Mock(return_value={
            'Place Name': 'Example', 'Full Address': address, 'Phone Number': phone,
            'Company Website': 'https://example.com/', 'Address Link': page.url})
        return {'Company Name': 'Example', 'City': 'Boston', 'State': 'Massachusetts',
                'Country': 'United States', 'Company Website': 'https://example.com/'}

    def test_matching_website_does_not_override_wrong_city(self):
        record = self.configure_place('100 Broadway, New York, NY 10001, USA', '+12125551234')
        result = self.worker._get_google_maps_company_details(record, 'test')
        self.assertFalse(result['company_matched'])

    def test_explicit_foreign_phone_is_not_used_for_us_place(self):
        record = self.configure_place('123 Main St, Boston, MA 02110, USA', '+44 800 2283957')
        result = self.worker._get_google_maps_company_details(record, 'test')
        self.assertTrue(result['company_matched'])
        self.assertEqual(result['phone'], '-')
        self.assertEqual(result['phone_link'], '-')
        self.assertEqual(result['zip_code'], '02110')

    def test_foreign_phone_rejected_when_domestic_address_omits_country(self):
        record = self.configure_place('123 Main St, Boston, MA 02110', '+44 800 2283957')
        result = self.worker._get_google_maps_company_details(record, 'test')
        self.assertTrue(result['company_matched'])
        self.assertEqual(result['phone'], '-')

    def test_same_address_can_receive_postcode_and_phone(self):
        address = '123 Main St, Boston, MA 02110, USA'
        record = {'Address': address}
        self.worker._merge_google_maps_fields(record, {'Address': address, 'Zip Code': '02110',
            'Phone Number': '+16175551234', 'Phone Number Link': 'tel:+16175551234'})
        self.assertEqual(record['Zip Code'], '02110')
        self.assertEqual(record['Phone Number Link'], 'tel:+16175551234')

    def test_country_name_ending_in_us_is_not_us(self):
        record = self.configure_place('1 Lenin St, Minsk, Belarus', '+375 17 1234567')
        record.update({'City': 'Minsk', 'State': '-', 'Country': 'Belarus'})
        result = self.worker._get_google_maps_company_details(record, 'test')
        self.assertTrue(result['company_matched'])
        self.assertEqual(result['phone'], '+375171234567')


if __name__ == '__main__':
    unittest.main(verbosity=2)
