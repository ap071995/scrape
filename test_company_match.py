import unittest
from unittest.mock import Mock
from test_scraper import app


class CompanyMatchTests(unittest.TestCase):
    def test_accepted_variations(self):
        for left, right in [('Cognizant', 'Cognizent'), ('Cognizant', 'COGNIZANT Pvt. Ltd.'),
                            ('Acme & Sons', 'ACME and Sons LLC'), ('Société Générale', 'Societe Generale'),
                            ('Global Acme Technologies', 'Acme Global Technologies'), ('Tech Mahindra', 'TechMahindra')]:
            with self.subTest(left=left, right=right):
                self.assertTrue(app.SearchWorker._fuzzy_company_match(left, right)[0])

    def test_distinct_companies_rejected(self):
        for left, right in [('Cognizant', 'Cognizant Moment'), ('ABC', 'ABCDEF'), ('IBM', 'IBN'),
                            ('Alpha Global Services', 'Beta Global Services'), ('', ''), ('-', 'Unknown'),
                            ('YPF', 'YPF LUZ'), ('Cognizant', 'Infosys')]:
            with self.subTest(left=left, right=right):
                self.assertFalse(app.SearchWorker._fuzzy_company_match(left, right)[0])

    def test_cached_company_uses_fuzzy_verification(self):
        worker = app.SearchWorker({})
        url = 'https://www.linkedin.com/company/cognizant'
        worker.company_cache['standard|' + url] = {'Company Name': 'Cognizant'}
        self.assertEqual(worker._read_standard_company_about(url, 'Cognizent')['Company Name'], 'Cognizant')

    def test_legal_suffix_and_rebrand_aliases(self):
        self.assertTrue(app.SearchWorker._link_company_names_match('YPF SA', 'YPF'))
        self.assertTrue(app.SearchWorker._link_company_names_match('BHP Billiton', 'BHP'))
        self.assertFalse(app.SearchWorker._link_company_names_match('Cognizant', 'Cognizant Moment'))

    def test_two_part_city_state_location(self):
        worker = app.SearchWorker({})
        self.assertEqual(worker._location_to_fields('Richmond, Victoria'),
                         ('Richmond', 'Victoria', '-'))
        self.assertTrue(worker._is_location_candidate('Richmond, Victoria'))

    def test_regional_linkedin_urls_deduplicate(self):
        self.assertEqual(
            app.normalize_url('http://ar.linkedin.com/in/example/?trk=x'),
            'https://www.linkedin.com/in/example',
        )

    def test_company_slug_fallback(self):
        self.assertEqual(
            app.SearchWorker._company_slug_fallback_url('Cognizant Pvt. Ltd.'),
            'https://www.linkedin.com/company/cognizant-pvt-ltd/about/',
        )
        self.assertEqual(app.SearchWorker._company_slug_fallback_url('-'), '-')
        self.assertEqual(
            app.SearchWorker._company_slug_fallback_urls('East Carolina University'),
            [
                'https://www.linkedin.com/school/east-carolina-university/about/',
                'https://www.linkedin.com/company/east-carolina-university/about/',
            ],
        )

    def test_school_employer_and_grouped_role_are_parsed(self):
        worker = app.SearchWorker({})
        role = worker._parse_standard_current_experience({
            'name': 'Zach Loch',
            'currentText': (
                'East Carolina University\n19 yrs 4 mos\n'
                'Chief Information Officer\nFull-time\n'
                'Dec 2020 - Present · 5 yrs 10 mos\nGreenville, NC'
            ),
            'companyLinks': [{
                'text': 'East Carolina University',
                'href': 'https://www.linkedin.com/school/7522/',
            }],
            'profileLocations': ['Greenville, North Carolina, United States'],
        })
        self.assertEqual(role['First Name'], 'Zach')
        self.assertEqual(role['Last Name'], 'Loch')
        self.assertEqual(role['Company Name'], 'East Carolina University')
        self.assertEqual(role['Company Profile Link'], 'https://www.linkedin.com/school/7522/')
        self.assertEqual(role['Job Title'], 'Chief Information Officer')
        self.assertEqual(role['Location'], 'Greenville, NC')
        self.assertEqual(worker._location_to_fields(role['Location'], role['Location']),
                         ('Greenville', 'NC', '-'))
        self.assertEqual(worker._location_to_fields(role['Location'], role['Profile Location']),
                         ('Greenville', 'NC', 'United States'))

    def test_employee_size_range_is_kept_intact(self):
        match = app.EMPLOYEE_SIZE_RE.search('Company size 5,001-10,000 employees')
        self.assertIsNotNone(match)
        self.assertEqual(match.group(0), '5,001-10,000 employees')

    def test_employment_status_is_not_company(self):
        worker = app.SearchWorker({})
        role = worker._parse_standard_current_experience({
            'name': 'Zach Loch',
            'currentText': (
                'East Carolina University\n19 yrs 4 mos\n'
                'Chief Information Officer\nFull-time\n'
                'Dec 2020 - Present · 5 yrs 10 mos\nGreenville, NC'
            ),
            'companyLinks': [],
            'profileLocations': ['Greenville, North Carolina, United States'],
        })
        self.assertEqual(role['Company Name'], 'East Carolina University')
        self.assertEqual(role['Job Title'], 'Chief Information Officer')

    def test_school_about_fields_are_read_from_definitions(self):
        worker = app.SearchWorker({})
        worker.page = Mock()
        worker.page.evaluate.return_value = {
            'companyName': 'East Carolina University',
            'bodyText': (
                'Overview\nWebsite\necu.edu\nIndustry\nHigher Education\n'
                'Company size\n5,001-10,000 employees'
            ),
            'definitions': {
                'website': 'ecu.edu',
                'industry': 'Higher Education',
                'company size': '5,001-10,000 employees',
            },
            'links': [{
                'text': 'ecu.edu',
                'href': 'https://ecu.edu/',
            }],
            'locationsText': 'Locations\nPrimary\nGreenville, North Carolina, United States\nGet directions',
            'locationLinks': [],
        }
        worker._open_standard_company_about = Mock(
            return_value='https://www.linkedin.com/school/7522/about/'
        )
        worker._scroll_standard_page = Mock()
        fields = worker._read_standard_company_about(
            'https://www.linkedin.com/school/7522/',
            'East Carolina University',
        )
        self.assertEqual(fields['Company Name'], 'East Carolina University')
        self.assertEqual(fields['Company Website'], 'https://ecu.edu/')
        self.assertEqual(fields['Industry'], 'Higher Education')
        self.assertEqual(fields['Company Size'], '5,001-10,000 employees')
        self.assertEqual(fields['City'], 'Greenville')
        self.assertEqual(fields['State'], 'North Carolina')
        self.assertEqual(fields['Country'], 'United States')

    def test_status_only_company_row_is_marked_for_repair(self):
        workbook = app.Workbook()
        sheet = workbook.active
        sheet.append(app.HEADERS)
        record = {header: 'ok' for header in app.HEADERS}
        record.update({
            'Prospect LinkedIn Link': 'https://www.linkedin.com/in/zach-loch-ba07b015',
            'Company Name': 'FULL TIME',
        })
        sheet.append([record[header] for header in app.HEADERS])
        manager_like = type('ManagerLike', (), {'ws': sheet})()
        self.assertEqual(
            app.WorkbookManager.incomplete_link_row(
                manager_like,
                record['Prospect LinkedIn Link'],
            ),
            2,
        )
        workbook.close()

    def test_primary_us_address_separates_postal_code(self):
        worker = app.SearchWorker({})
        worker.page = Mock()
        worker._open_standard_company_about = Mock(return_value='https://www.linkedin.com/school/7522/about/')
        worker._scroll_standard_page = Mock()
        worker.page.evaluate.return_value = {
            'companyName': 'East Carolina University',
            'locationsText': 'Primary\nEast 5th Street, Greenville, NC 27858-4353, US',
        }
        fields = worker._read_standard_company_about('https://www.linkedin.com/school/7522/', 'East Carolina University')
        self.assertEqual(fields['City'], 'Greenville')
        self.assertEqual(fields['State'], 'NC')
        self.assertEqual(fields['Zip Code'], '27858-4353')
        self.assertEqual(fields['Country'], 'United States')
