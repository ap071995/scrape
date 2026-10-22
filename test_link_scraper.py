"""Browser fixtures transcribed from the supplied screenshots; not live data."""
import tempfile
from pathlib import Path
from unittest.mock import Mock, patch
from test_browser import BrowserTests
from test_scraper import app

PROFILE = '''<main><h1>Pranoday Dingare</h1><p>Pune, Maharashtra, India</p>
<section><h2>Experience</h2><ul><li>
<a href="https://www.linkedin.com/company/cognizant/"><img alt="Cognizant logo" width="40" height="40"></a>
<div>Sr. Manager - Projects</div><div>Cognizant · Full-time</div>
<div>Feb 2025 - Present · 1 yr 8 mos</div><div>Pune District, Maharashtra, India</div>
<p>In my current role as a Senior Manager at Cognizant, I lead a highly skilled team.</p>
</li><li><div>Magic Software Enterprises</div><div>Test Automation Architect and Lead DevOps Engineer</div>
<div>Sep 2022 - Feb 2025 · 2 yrs 6 mos</div></li></ul></section></main>'''
ABOUT = '''<main><h1>Cognizant</h1><a href="/company/cognizant/about/">About</a>
<section><h2>Overview</h2><dl><dt>Website</dt><dd><a href="https://www.cognizant.ai">https://www.cognizant.ai</a></dd>
<dt>Industry</dt><dd>IT Services and IT Consulting</dd>
<dt>Company size</dt><dd>10,001+ employees</dd>
<dd><a href="/search/results/people/">350,123 employees</a></dd>
<dt>Headquarters</dt><dd>Teaneck, New Jersey, United States</dd></dl></section>
</main><aside><h2>Cognizant Moment</h2><p>Marketing Services</p>
<a href="https://unrelated.example">https://unrelated.example</a></aside>'''


class LinkScraperTests(BrowserTests):
    def setUp(self):
        super().setUp()
        self.worker.page = self.page
        self.worker._wait_for_manual_login_if_needed = Mock()
        self.worker._launch_browser = Mock()
        self.worker._google_maps_enrich_record = lambda record, **kwargs: record
        self.worker.callbacks = {key: Mock() for key in ['progress', 'record', 'harvest', 'done', 'log']}
        self.requests = []
        def route(request):
            url = request.request.url
            self.requests.append(url)
            request.fulfill(status=200, content_type='text/html; charset=utf-8', body=PROFILE if '/in/' in url else ABOUT)
        self.page.route('**/*', route)

    def test_screenshot_example_through_workbook(self):
        with tempfile.TemporaryDirectory(dir='.') as temporary:
            folder = Path(temporary)
            source = folder / 'profiles.csv'
            source.write_text('URL\nhttps://www.linkedin.com/in/trainerpranoday-dingare/\n')
            destination = folder / 'output.xlsx'
            manager = app.WorkbookManager(str(destination))
            with patch.multiple(app, LINK_PROFILE_SETTLE_MS=1, LINK_COMPANY_SETTLE_MS=1, LINK_COMPANY_MAX_SCROLLS=1):
                self.worker._process_link_file({'input_path': str(source), 'max_records': 1,
                                               'workbook': manager, 'extraction_mode': 'FAST'})
            workbook = app.load_workbook(destination)
            self.assertEqual(workbook.active.max_row, 2, self.worker.callbacks['log'].call_args_list)
            row = dict(zip(app.HEADERS, list(workbook.active.values)[1]))
            expected = {'First Name': 'Pranoday', 'Last Name': 'Dingare', 'Company Name': 'Cognizant',
                        'Job Title': 'Sr. Manager - Projects', 'Tenure': 'Feb 2025 - Present · 1 yr 8 mos',
                        'City': 'Pune District', 'State': 'Maharashtra', 'Country': 'India',
                        'Industry': 'IT Services and IT Consulting', 'Company Website': 'https://www.cognizant.ai/',
                        'Company Size': '10,001+ employees'}
            for key, value in expected.items():
                self.assertEqual(row[key], value, key)
            self.assertTrue(any('/company/cognizant/about/' in url for url in self.requests))
            workbook.close()

    def test_grouped_role_uses_current_title(self):
        grouped = '''<main><h1>Pranoday Dingare</h1><section><h2>Experience</h2><ul><li>
        <a href="/company/cognizant/"><img alt="Cognizant logo" width="40" height="40"></a>
        <div>Cognizant</div><ul><li><div>Sr. Manager - Projects</div>
        <div>Feb 2025 - Present · 1 yr 8 mos</div><div>Pune District, Maharashtra, India · Hybrid</div></li>
        <li><div>Engineer</div><div>Jan 2020 - Jan 2025</div></li></ul></li></ul></section></main>'''
        self.page.unroute('**/*')
        self.page.route('**/*', lambda route: route.fulfill(content_type='text/html; charset=utf-8', body=grouped))
        with patch.object(app, 'LINK_PROFILE_SETTLE_MS', 1):
            data = self.worker._capture_standard_profile('https://www.linkedin.com/in/example/')
        role = self.worker._parse_standard_current_experience(data)
        self.assertEqual(role['Job Title'], 'Sr. Manager - Projects')
        self.assertEqual(role['Company Name'], 'Cognizant')
        self.assertEqual(role['Location'], 'Pune District, Maharashtra, India')

    def test_wrong_company_rejected(self):
        with patch.multiple(app, LINK_COMPANY_SETTLE_MS=1, LINK_COMPANY_MAX_SCROLLS=1):
            with self.assertRaises(app.CompanyPageMismatchError):
                self.worker._read_standard_company_about('https://www.linkedin.com/company/cognizant/', 'Other Company')

    def test_present_in_description_is_not_current_employment(self):
        result = self.worker._parse_standard_current_experience({'name': 'Pranoday Dingare',
            'currentText': 'Engineer\nExample\nJan 2020 - Feb 2025\nI present reports to management.'})
        self.assertEqual(result['Tenure'], 'Dead')

    def test_related_company_is_not_same_employer(self):
        self.worker.company_cache['standard|https://www.linkedin.com/company/cognizant'] = {
            'Company Name': 'Cognizant Moment'}
        with self.assertRaises(app.CompanyPageMismatchError):
            self.worker._read_standard_company_about('https://www.linkedin.com/company/cognizant/', 'Cognizant')

    def test_profile_name_without_h1(self):
        self.page.set_content('<main><section><h2>Santiago Sajaroff</h2></section></main>')
        self.assertEqual(self.worker._standard_profile_name(500), 'Santiago Sajaroff')

    def test_delayed_name_rendering(self):
        self.page.set_content('''<main><section></section></main><script>
        setTimeout(() => document.querySelector('section').innerHTML = '<h2>Walter Ganser</h2>', 300);
        </script>''')
        self.assertEqual(self.worker._standard_profile_name(2000), 'Walter Ganser')

    def test_title_fallback_ignores_experience_heading(self):
        self.page.goto('https://www.linkedin.com/in/example/')
        self.page.set_content('<title>(3) Walter Ganser - Example | LinkedIn</title><main><section><h2>Experience</h2></section></main>')
        self.assertEqual(self.worker._standard_profile_name(500), 'Walter Ganser')

    def test_authwall_does_not_become_person_name(self):
        self.page.set_content('<main><h1>Sign in</h1><input type="password"></main>')
        with self.assertRaises(app.ProfileAccessError):
            self.worker._standard_profile_name(500)

    def test_regional_profile_url_normalization(self):
        self.assertEqual(app.canonical_profile_url('http://ar.linkedin.com/in/santiago-sajaroff-4b29161b'),
                         'https://www.linkedin.com/in/santiago-sajaroff-4b29161b')
        self.assertEqual(app.canonical_profile_url('http://at.linkedin.com/in/walter-ganser-b99b84b5'),
                         'https://www.linkedin.com/in/walter-ganser-b99b84b5')
        with self.assertRaises(ValueError):
            app.canonical_profile_url('https://notlinkedin.com/in/fake')

    def test_experience_div_layout(self):
        self.page.unroute('**/*')
        html = PROFILE.replace('<section>', '<div>').replace('</section>', '</div>')
        self.page.route('**/*', lambda route: route.fulfill(content_type='text/html; charset=utf-8', body=html))
        with patch.multiple(app, LINK_PROFILE_SETTLE_MS=1, LINK_PROFILE_MAX_SCROLLS=1):
            data = self.worker._capture_standard_profile('https://www.linkedin.com/in/example/')
        current = self.worker._parse_standard_current_experience(data)
        self.assertEqual(current['Company Name'], 'Cognizant')
        self.assertEqual(current['Job Title'], 'Sr. Manager - Projects')

    def test_experience_detail_page_fallback(self):
        self.page.unroute('**/*')
        def route(request):
            html = PROFILE if '/details/experience/' in request.request.url else '<main><h1>Pranoday Dingare</h1></main>'
            request.fulfill(content_type='text/html; charset=utf-8', body=html)
        self.page.route('**/*', route)
        with patch.multiple(app, LINK_PROFILE_SETTLE_MS=1, LINK_PROFILE_MAX_SCROLLS=1):
            data = self.worker._capture_standard_profile('https://www.linkedin.com/in/example/')
        self.assertIn('/details/experience/', self.page.url)
        self.assertEqual(self.worker._parse_standard_current_experience(data)['Company Name'], 'Cognizant')

    def test_missing_experience_is_failure_not_dead_row(self):
        self.page.unroute('**/*')
        self.page.route('**/*', lambda route: route.fulfill(content_type='text/html', body='<main><h1>Test Person</h1></main>'))
        with patch.multiple(app, LINK_PROFILE_SETTLE_MS=1, LINK_PROFILE_MAX_SCROLLS=1):
            with self.assertRaisesRegex(RuntimeError, 'not saved as Dead'):
                self.worker._capture_standard_profile('https://www.linkedin.com/in/example/')

    def test_incomplete_row_is_repaired_without_duplicate(self):
        with tempfile.TemporaryDirectory(dir='.') as temporary:
            folder = Path(temporary)
            source = folder / 'profiles.csv'
            url = 'https://www.linkedin.com/in/example'
            source.write_text('URL\n' + url + '\n')
            manager = app.WorkbookManager(str(folder / 'output.xlsx'))
            manager.add_record({'First Name': 'Pranoday', 'Tenure': 'Dead', 'Prospect LinkedIn Link': url,
                                'Campaign Name': 'Keep campaign'})
            with patch.multiple(app, LINK_PROFILE_SETTLE_MS=1, LINK_PROFILE_MAX_SCROLLS=1,
                                LINK_COMPANY_SETTLE_MS=1, LINK_COMPANY_MAX_SCROLLS=1):
                self.worker._process_link_file({'input_path': str(source), 'max_records': 1,
                                               'workbook': manager, 'extraction_mode': 'FAST'})
            reopened = app.load_workbook(manager.output_path)
            self.assertEqual(reopened.active.max_row, 2)
            row = dict(zip(app.HEADERS, list(reopened.active.values)[1]))
            self.assertEqual(row['Company Name'], 'Cognizant')
            self.assertEqual(row['Campaign Name'], 'Keep campaign')
            self.assertNotEqual(row['Tenure'], 'Dead')
            reopened.close()

    def test_parser_infers_employer_without_company_link(self):
        role = self.worker._parse_standard_current_experience({
            'name': 'Andrew Rovers',
            'currentText': 'Construction Manager\nCivil Infrastructure Group\nMay 2016 - Present · 10 yrs 5 mos',
            'companyLinks': [],
            'profileLocations': ['Melbourne, Victoria, Australia'],
        })
        self.assertEqual(role['Company Name'], 'Civil Infrastructure Group')
        self.assertEqual(role['Job Title'], 'Construction Manager')
        self.assertEqual(role['Company Profile Link'], '-')
        self.assertEqual(role['Location'], 'Melbourne, Victoria, Australia')

    def test_same_employer_retry_does_not_restore_unverified_contact_data(self):
        with tempfile.TemporaryDirectory(dir='.') as temporary:
            folder = Path(temporary)
            source = folder / 'profiles.csv'
            url = 'https://www.linkedin.com/in/example'
            source.write_text('URL\n' + url + '\n')
            manager = app.WorkbookManager(str(folder / 'output.xlsx'))
            manager.add_record({'Company Name': 'Cognizant', 'Prospect LinkedIn Link': url,
                                'Phone Number': '+919999999999', 'Phone Number Link': 'tel:+919999999999',
                                'Zip Code': '411001', 'Address Link': 'https://wrong.example/',
                                'Notes': 'Google Maps enriched: Phone Number',
                                'Campaign Name': 'Keep campaign', 'Source File Name': 'original.xlsx'})
            with patch.multiple(app, LINK_PROFILE_SETTLE_MS=1, LINK_COMPANY_SETTLE_MS=1,
                                LINK_COMPANY_MAX_SCROLLS=1):
                self.worker._process_link_file({'input_path': str(source), 'max_records': 1,
                                               'workbook': manager, 'extraction_mode': 'FAST'})
            reopened = app.load_workbook(manager.output_path)
            self.assertEqual(reopened.active.max_row, 2)
            row = dict(zip(app.HEADERS, list(reopened.active.values)[1]))
            for field in ['Phone Number', 'Phone Number Link', 'Zip Code', 'Address Link', 'Notes']:
                self.assertEqual(row[field], '-', field)
            self.assertEqual(row['Campaign Name'], 'Keep campaign')
            self.assertEqual(row['Source File Name'], 'original.xlsx')
            reopened.close()

    def test_university_profile_repairs_wrong_employer_in_excel(self):
        profile = '''<main><h1>Zach Loch</h1>
        <span data-anonymize="location">Greenville, North Carolina, United States</span>
        <section><h2>Experience</h2><ul><li>
        <a href="/school/7522/"><img alt="East Carolina University logo" width="40" height="40"></a>
        <a href="/school/7522/">East Carolina University</a><div>19 yrs 4 mos</div>
        <ul><li><div>Chief Information Officer</div><div>Full-time</div>
        <div>Dec 2020 - Present · 5 yrs 10 mos</div></li>
        <li><div>Interim CIO</div><div>Apr 2020 - Dec 2020 · 9 mos</div></li>
        </ul></li></ul></section></main>'''
        about = '''<main><h1>East Carolina University</h1>
        <a href="/school/7522/about/">About</a><section><h2>Overview</h2><dl>
        <dt>Website</dt><dd><a href="https://ecu.edu/">ecu.edu</a></dd>
        <dt>Industry</dt><dd>Higher Education</dd>
        <dt>Company size</dt><dd>5,001-10,000 employees</dd>
        <dd><a href="/search/results/people/">8,765 associated members</a></dd>
        </dl></section></main>'''
        self.page.unroute('**/*')
        requests = []
        def route(request):
            url = request.request.url
            requests.append(url)
            request.fulfill(content_type='text/html; charset=utf-8',
                            body=profile if '/in/' in url else about)
        self.page.route('**/*', route)
        with tempfile.TemporaryDirectory(dir='.') as temporary:
            folder = Path(temporary)
            source = folder / 'profiles.csv'
            url = 'https://www.linkedin.com/in/zach-loch-ba07b015'
            source.write_text('URL\n' + url + '\n')
            manager = app.WorkbookManager(str(folder / 'output.xlsx'))
            manager.add_record({'Prospect LinkedIn Link': url, 'Company Name': 'FULL TIME',
                                'Company Website': 'https://wrong.example/',
                                'Phone Number': 'wrong phone', 'Zip Code': 'wrong zip',
                                'Campaign Name': 'Keep campaign'})
            with patch.multiple(app, LINK_PROFILE_SETTLE_MS=1, LINK_COMPANY_SETTLE_MS=1,
                                LINK_COMPANY_MAX_SCROLLS=1):
                self.worker._process_link_file({'input_path': str(source), 'max_records': 1,
                                               'workbook': manager, 'extraction_mode': 'FAST'})
            workbook = app.load_workbook(manager.output_path)
            self.assertEqual(workbook.active.max_row, 2)
            row = dict(zip(app.HEADERS, list(workbook.active.values)[1]))
            expected = {'First Name': 'Zach', 'Last Name': 'Loch',
                        'Company Name': 'East Carolina University',
                        'Job Title': 'Chief Information Officer',
                        'Tenure': 'Dec 2020 - Present · 5 yrs 10 mos',
                        'Company Website': 'https://ecu.edu/', 'Industry': 'Higher Education',
                        'Company Size': '5,001-10,000 employees', 'City': 'Greenville',
                        'State': 'North Carolina', 'Country': 'United States',
                        'Phone Number': '-', 'Zip Code': '-', 'Campaign Name': 'Keep campaign'}
            for key, value in expected.items():
                self.assertEqual(row[key], value, key)
            self.assertTrue(any('/school/7522/about/' in item for item in requests))
            workbook.close()

    def test_primary_office_excludes_adjacent_branch_and_description(self):
        self.page.set_content('''<main><h1>Example Bank</h1><section><h3>Locations (2)</h3><ul>
        <li><span>Primary</span><h4>We operate in many countries.</h4>
        <p>Postbox, Zurich, 8001, CH</p><a href="https://maps.example/primary">Get directions</a></li>
        <li><p>5 Broadgate, London, EC2M 2QS, GB</p>
        <a href="https://maps.example/branch">Get directions</a></li></ul></section></main>''')
        self.worker._open_standard_company_about = Mock(return_value='https://www.linkedin.com/company/example/about/')
        self.worker._scroll_standard_page = Mock()
        fields = self.worker._read_standard_company_about('https://www.linkedin.com/company/example/', 'Example Bank')
        self.assertEqual(fields['Address'], 'Postbox, Zurich, 8001, CH')
        self.assertEqual(fields['City'], 'Zurich')
        self.assertEqual(fields['State'], '-')
        self.assertEqual(fields['Country'], 'Switzerland')
        self.assertEqual(fields['Zip Code'], '8001')
        self.assertEqual(fields['Address Link'], 'https://maps.example/primary')

    def test_about_phone_uses_number_from_tel_link(self):
        self.page.set_content('''<main><h1>Example</h1><section><dl><dt>Website</dt><dd>example.com</dd>
        <dt>Phone</dt><dd><a href="tel:2488519600">2488519600 <span>Phone number is 2488519600</span></a></dd>
        </dl></section></main>''')
        self.worker._open_standard_company_about = Mock(return_value='https://www.linkedin.com/company/example/about/')
        self.worker._scroll_standard_page = Mock()
        fields = self.worker._read_standard_company_about('https://www.linkedin.com/company/example/', 'Example')
        self.assertEqual(fields['Phone Number'], '2488519600')

