"""Offline regressions for failures found in real LinkedIn profile markup.

The Zach fixture preserves the nested employer/role anchor shape captured on
2026-09-17. All browser requests are aborted; these tests never access LinkedIn.
"""
import unittest
from unittest.mock import Mock

from test_scraper import app


ZACH_ROLE_ANCHORS = '''<main><h1>Zach Loch</h1>
<section><h2>Experience</h2><div class="employer-group">
  <a href="https://www.linkedin.com/school/7522/">
    <img alt="East Carolina University logo" width="40" height="40">
  </a>
  <div><div><a href="https://www.linkedin.com/school/7522/">
    <div><p>East Carolina University</p><p>19 yrs 4 mos</p></div>
  </a></div></div>
  <ul>
    <li><div><a href="https://www.linkedin.com/school/7522/">
      <div><p>Chief Information Officer</p><p>Full-time</p>
      <p>Dec 2020 - Present · 5 yrs 10 mos</p></div>
    </a><div></div></div></li>
    <li><div><a href="https://www.linkedin.com/school/7522/">
      <div><p>Interim CIO</p><p>Apr 2020 - Dec 2020 · 9 mos</p></div>
    </a></div></li>
  </ul>
</div></section></main>'''


class LiveLayoutRegressionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.playwright = app.sync_playwright().start()
        cls.addClassCleanup(cls.playwright.stop)
        cls.browser = cls.playwright.chromium.launch(channel='chrome', headless=True, chromium_sandbox=True)
        cls.addClassCleanup(cls.browser.close)

    def setUp(self):
        self.context = self.browser.new_context(service_workers='block')
        self.context.route('**/*', lambda route: route.abort())
        self.page = self.context.new_page()
        self.worker = app.SearchWorker({})
        self.worker.page = self.page

    def tearDown(self):
        self.context.close()

    def parse_page(self, html):
        self.page.set_content(html)
        data = self.worker._read_standard_experience_dom()
        data['name'] = 'Zach Loch'
        return data, self.worker._parse_standard_current_experience(data)

    def test_role_link_does_not_become_employer_label(self):
        data, role = self.parse_page(ZACH_ROLE_ANCHORS)
        self.assertEqual(role['Company Name'], 'East Carolina University')
        self.assertEqual(role['Company Profile Link'], 'https://www.linkedin.com/school/7522/')
        self.assertEqual(role['Job Title'], 'Chief Information Officer')
        self.assertEqual(role['Tenure'], 'Dec 2020 - Present · 5 yrs 10 mos')
        self.assertTrue(data['companyLinksScoped'])
        self.assertFalse(any('Present' in item['text'] for item in data['companyLinks']))

    def test_nested_metadata_retains_sibling_title(self):
        _, role = self.parse_page('''<main><section><h2>Experience</h2><ul><li>
        <a href="https://www.linkedin.com/school/7522/">East Carolina University</a>
        <div>19 yrs 4 mos</div><ul><li><div>Chief Information Officer</div>
        <div><div>Full-time</div><div>Dec 2020 - Present · 5 yrs 10 mos</div>
        <div>Greenville, NC</div></div></li></ul></li></ul></section></main>''')
        self.assertEqual(role['Job Title'], 'Chief Information Officer')
        self.assertEqual(role['Company Name'], 'East Carolina University')
        self.assertEqual(role['Location'], 'Greenville, NC')

    def test_first_current_role_is_not_replaced_by_shorter_later_role(self):
        _, role = self.parse_page('''<main><section><h2>Experience</h2><ul>
        <li><a href="https://www.linkedin.com/company/main/">Main Company</a>
        <div>Chief Information Officer</div><div>Jan 2020 - Present</div>
        <div>Greenville, NC</div></li>
        <li><a href="https://www.linkedin.com/company/side/">Side</a>
        <div>Advisor</div><div>2024 - Present</div></li>
        </ul></section></main>''')
        self.assertEqual(role['Company Name'], 'Main Company')
        self.assertEqual(role['Job Title'], 'Chief Information Officer')

    def test_company_title_notification_count_is_ignored_without_h1(self):
        self.page.set_content('''<title>(3) East Carolina University: About | LinkedIn</title>
        <main><section><h2>Overview</h2><dl>
        <dt>Website</dt><dd><a href="https://ecu.edu/">ecu.edu</a></dd>
        <dt>Industry</dt><dd>Higher Education</dd>
        <dt>Company size</dt><dd>5,001-10,000 employees</dd>
        </dl></section></main>''')
        self.worker._open_standard_company_about = Mock(
            return_value='https://www.linkedin.com/school/7522/about/')
        self.worker._scroll_standard_page = Mock()
        result = self.worker._read_standard_company_about(
            'https://www.linkedin.com/school/7522/', 'East Carolina University')
        self.assertEqual(result['Company Name'], 'East Carolina University')
        self.assertEqual(result['Company Website'], 'https://ecu.edu/')
        self.assertEqual(result['Industry'], 'Higher Education')
        self.assertEqual(result['Company Size'], '5,001-10,000 employees')


class EmployerIdentityRegressionTests(unittest.TestCase):
    def setUp(self):
        self.worker = app.SearchWorker({})

    def test_unscoped_school_link_cannot_supply_unknown_employer(self):
        result = self.worker._parse_standard_current_experience({
            'name': 'Example Person',
            'currentText': 'Manager\nFull-time\nJan 2025 - Present\nBoston, MA',
            'companyLinksScoped': False,
            'companyLinks': [{'text': 'Unrelated University',
                              'href': 'https://www.linkedin.com/school/unrelated/'}],
        })
        self.assertEqual(result['Company Name'], '-')
        self.assertEqual(result['Company Profile Link'], '-')

    def test_subsidiary_suffix_cannot_bypass_fuzzy_rejection(self):
        self.assertFalse(self.worker._link_company_names_match(
            'Alpha Health Services', 'Alpha Health Services International'))
        # Legal suffix changes and explicitly supported aliases remain valid.
        self.assertTrue(self.worker._link_company_names_match('YPF SA', 'YPF'))
        self.assertTrue(self.worker._link_company_names_match('BHP Billiton', 'BHP'))


if __name__ == '__main__':
    unittest.main(verbosity=2)
