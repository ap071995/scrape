"""Adversarial email ownership and evidence-precedence regressions (offline)."""
import copy
import unittest
from unittest.mock import Mock

from business_email import BusinessEmailEnricher, FOUND, GENERATED, pattern_evidence, published_emails
from company_public_data import SourceDocument
from test_business_email import company_match, document, record, search_row


class StrictEmailOwnershipTests(unittest.TestCase):
    def test_uppercase_next_staff_heading_does_not_belong_to_previous_person(self):
        for heading in ('JOHN SMITH', 'JOHN SMITH, Manager', 'John SMITH'):
            with self.subTest(heading=heading):
                self.assertEqual(published_emails(document('Jane Doe\nDirector\n' + heading +
                    '\nEmail: john.smith@example.com'), record()), [])

    def test_inline_uppercase_next_staff_entry_is_not_borrowed(self):
        self.assertEqual(published_emails(document('Jane Doe | JOHN SMITH | john.smith@example.com'), record()), [])

    def test_previous_inline_staff_email_does_not_belong_to_later_name(self):
        self.assertEqual(published_emails(document('JOHN SMITH | john.smith@example.com | Jane Doe'), record()), [])

    def test_matching_address_shape_does_not_override_other_person_ownership(self):
        for text in ('Jane Doe | JOHN SMITH | jane.doe@example.com',
                     'JOHN SMITH | jane.doe@example.com | Jane Doe'):
            for kind in ('page', 'search'):
                with self.subTest(text=text, kind=kind):
                    self.assertEqual(published_emails(document(text), record(), kind), [])

    def test_uppercase_target_name_remains_accepted(self):
        facts = published_emails(document('JANE DOE\nDIRECTOR\nEmail: jd742@example.com'), record())
        self.assertEqual([item['email'] for item in facts], ['jd742@example.com'])

    def test_unicode_masked_fragments_are_not_treated_as_full_addresses(self):
        for address in ('j••doe@example.com', 'j…doe@example.com', '••jane.doe@example.com',
                        'j●●doe@example.com'):
            with self.subTest(address=address):
                self.assertEqual(published_emails(document('Jane Doe\nEmail: ' + address), record()), [])

    def test_bullet_list_with_whitespace_can_publish_complete_address(self):
        facts = published_emails(document('Jane Doe\n• jane.doe@example.com'), record())
        self.assertEqual([item['email'] for item in facts], ['jane.doe@example.com'])

    def test_prose_delegation_is_not_person_ownership(self):
        for text in ('Jane Doe\nFor appointments email her assistant at sam.adams@example.com',
                     'Jane Doe\nPlease contact the executive assistant: sam.adams@example.com',
                     'Jane Doe\nEmail on behalf of Jane: sam.adams@example.com'):
            with self.subTest(text=text):
                self.assertEqual(published_emails(document(text), record()), [])

    def test_delegated_mailto_link_retains_the_delegation_context(self):
        doc = document('Jane Doe\nDirector')
        doc.links = [{'url': 'mailto:sam.adams@example.com', 'text': 'Contact',
                      'context': 'Jane Doe\nEmail her assistant at this link'}]
        self.assertEqual(published_emails(doc, record()), [])

    def test_mailto_query_header_is_not_part_of_the_recipient(self):
        doc = document('Jane Doe\nDirector')
        doc.links = [{'url': 'mailto:jd742%40example.com?subject=Hello&cc=other@example.com',
                      'text': 'Email', 'context': 'Jane Doe\nDirector\nEmail'}]
        facts = published_emails(doc, record())
        self.assertEqual([item['email'] for item in facts], ['jd742@example.com'])

    def test_multiple_mailto_recipients_are_not_assigned_as_one_person_email(self):
        doc = document('Jane Doe\nDirector')
        doc.links = [{'url': 'mailto:jd742@example.com,other@example.com',
                      'text': 'Email', 'context': 'Jane Doe\nEmail'}]
        self.assertEqual(published_emails(doc, record()), [])

    def test_person_specific_search_title_cannot_borrow_another_named_entry(self):
        doc = document('JOHN SMITH - jane.doe@example.com', 'Jane Doe | Example Industries',
                       'https://directory.example/jane-doe')
        self.assertEqual(published_emails(doc, record(), 'search'), [])

    def test_person_specific_search_title_can_supply_a_missing_name(self):
        facts = published_emails(document('Email: jane.doe@example.com'), record(), 'search')
        self.assertEqual([item['email'] for item in facts], ['jane.doe@example.com'])

    def test_nonpublic_source_cannot_supply_a_published_fact(self):
        for url in ('javascript:alert(1)', 'file:///tmp/staff.html', 'http://127.0.0.1/staff'):
            with self.subTest(url=url):
                self.assertEqual(published_emails(document('Jane Doe\nEmail: jane.doe@example.com', url=url), record()), [])

    def test_nonpublic_source_cannot_supply_a_company_pattern(self):
        doc = document('Example Industries email format: {first}.{last}@example.com',
                       'Example Industries Email Format', 'file:///tmp/format.html')
        self.assertEqual(pattern_evidence(doc, record(), company_match), [])

    def test_malformed_company_website_cannot_establish_published_ownership(self):
        self.assertEqual(published_emails(document('Jane Doe\nEmail: jane.doe@example.com'),
                         record(**{'Company Website': 'https://[invalid'})), [])

    def test_null_or_malformed_links_do_not_discard_visible_evidence(self):
        for links in (None, [None, {}, {'url': None}, {'url': 42}]):
            with self.subTest(links=links):
                doc = document('Jane Doe\nEmail: jane.doe@example.com')
                doc.links = links
                self.assertEqual([item['email'] for item in published_emails(doc, record())],
                                 ['jane.doe@example.com'])

    def test_lookalike_domain_is_not_authoritative_for_other_company_mailbox(self):
        doc = document('Jane Doe\nEmail: jane.doe@othercorp.com', url='https://example.com.evil.example/staff')
        self.assertEqual(published_emails(doc, record()), [])

    def test_long_unrelated_section_separates_name_from_email(self):
        doc = document('Jane Doe\n' + ('Company information ' * 50) + '\nEmail: someone@example.com')
        self.assertEqual(published_emails(doc, record()), [])


class StrictEmailWorkflowTests(unittest.TestCase):
    def setUp(self):
        self.web = Mock()
        self.web.search.return_value = []
        self.web.fetch.return_value = None
        self.web.documents = {}
        self.web.source_index = {}
        self.service = BusinessEmailEnricher(self.web, company_match)

    def test_cached_official_address_precedes_lower_quality_search_fact(self):
        official = document('Jane Doe\nEmail: jd742@example.com')
        snippet = document('Jane Doe\nEmail: jane.doe@example.com', url='https://directory.example/jane')
        self.web.documents = {official.url: official}
        self.web.search.return_value = [search_row(snippet)]
        result = self.service.enrich(record())
        self.assertEqual(result['Business Email ID'], 'jd742@example.com')
        self.assertEqual(result['_email_evidence'][0]['url'], official.url)

    def test_later_official_search_can_correct_an_earlier_third_party_snippet(self):
        snippet = document('Jane Doe\nEmail: jane.doe@example.com', url='https://directory.example/jane')
        official = document('Jane Doe\nEmail: jd742@example.com')
        self.web.search.side_effect = lambda query: [search_row(official if query.startswith('site:') else snippet)]
        self.web.fetch.side_effect = lambda url: official if url == official.url else None
        result = self.service.enrich(record())
        self.assertEqual(result['Business Email ID'], 'jd742@example.com')

    def test_official_page_gets_fetch_budget_after_six_third_party_results(self):
        snippets = [document('Jane Doe\nEmail: jane.doe@example.com',
                             url=f'https://directory.example/person/{index}') for index in range(6)]
        official_snippet = document('Jane Doe\nDirector\nContact')
        official = document('Jane Doe\nEmail: jd742@example.com')
        self.web.search.side_effect = lambda query: ([search_row(official_snippet)] if query.startswith('site:')
                                                     else [search_row(doc) for doc in snippets])
        self.web.fetch.side_effect = lambda url: official if url == official.url else None
        self.assertEqual(self.service.enrich(record())['Business Email ID'], 'jd742@example.com')

    def test_malformed_search_rows_do_not_discard_valid_official_result(self):
        official = document('Jane Doe\nEmail: jd742@example.com')
        self.web.search.return_value = [None, 'unusable row', {}, search_row(official)]
        self.web.fetch.return_value = official
        self.assertEqual(self.service.enrich(record())['Business Email ID'], 'jd742@example.com')

    def test_invalid_search_provenance_falls_back_to_public_result_url(self):
        snippet = document('Jane Doe\nEmail: jane.doe@example.com', url='https://directory.example/jane')
        self.web.search.return_value = [{**search_row(snippet), 'search_url': 'javascript:alert(1)'}]
        self.assertEqual(self.service.enrich(record())['_email_evidence'][0]['url'], snippet.url)

    def test_conflicting_snippets_do_not_prevent_official_directory_resolution(self):
        first = document('Jane Doe\nEmail: jane.doe@example.com', url='https://directory.example/one')
        second = document('Jane Doe\nEmail: jdoe@example.com', url='https://directory.example/two')
        official = document('Jane Doe\nEmail: jd742@example.com')
        self.web.search.return_value = [search_row(first), search_row(second)]
        self.web.source_index = {'example.com': {'company': 'Example Industries', 'email_urls': [official.url]}}
        self.web.fetch.side_effect = lambda url: official if url == official.url else None
        result = self.service.enrich(record())
        self.assertEqual(result['Business Email ID'], 'jd742@example.com')

    def test_conflicting_official_addresses_remain_unresolved(self):
        for index, address in enumerate(('jdoe@example.com', 'jd742@example.com')):
            doc = document('Jane Doe\nEmail: ' + address, url=f'https://example.com/staff/{index}')
            self.web.documents[doc.url] = doc
        result = self.service.enrich(record())
        self.assertEqual(result['Email Status'], 'Ambiguous (unverified)')
        self.assertEqual(result['Business Email ID'], '-')

    def test_official_search_must_not_bypass_known_conflicting_directory(self):
        searched = document('Jane Doe\nEmail: jdoe@example.com')
        seeded = document('Jane Doe\nEmail: jd742@example.com', url='https://example.com/team')
        self.web.search.return_value = [search_row(searched)]
        self.web.source_index = {'example.com': {'company': 'Example Industries', 'email_urls': [seeded.url]}}
        self.web.fetch.side_effect = lambda url: {searched.url: searched, seeded.url: seeded}.get(url)
        result = self.service.enrich(record())
        self.assertEqual(result['Email Status'], 'Ambiguous (unverified)')
        self.assertEqual(result['Business Email ID'], '-')

    def test_generated_candidate_is_rechecked_in_same_session(self):
        pattern = document('Example Industries email format: {first}.{last}@example.com',
                           'Example Industries Email Format', 'https://directory.example/company')
        self.web.search.side_effect = lambda query: [search_row(pattern)] if 'format' in query else []
        self.web.fetch.side_effect = lambda url: pattern if url == pattern.url else None
        first = self.service.enrich(record())
        self.assertEqual(first['Email Status'], GENERATED)
        official = document('Jane Doe\nEmail: jd742@example.com')
        self.web.documents[official.url] = official
        second = self.service.enrich(first)
        self.assertEqual(second['Business Email ID'], 'jd742@example.com')
        self.assertEqual(second['Email Status'], FOUND)

    def test_pattern_search_fetched_page_prefers_published_person_address(self):
        source = document('Example Industries email format: {first}.{last}@example.com\n'
                          'Jane Doe\nEmail: jd742@example.com',
                          'Example Industries Email Format', 'https://example.com/email-format')
        self.web.search.side_effect = lambda query: (
            [search_row(source, text='Example Industries email format')] if 'email format' in query else [])
        self.web.fetch.return_value = source
        result = self.service.enrich(record())
        self.assertEqual(result['Business Email ID'], 'jd742@example.com')
        self.assertEqual(result['Email Status'], FOUND)
        self.assertEqual(result['_email_evidence'][0]['url'], source.url)

    def test_pattern_search_snippet_prefers_published_person_address(self):
        source = document('Example Industries email format: {first}.{last}@example.com\n'
                          'Jane Doe\nEmail: jd742@example.com',
                          'Example Industries Email Format', 'https://example.com/email-format')
        self.web.search.side_effect = lambda query: [search_row(source)] if 'email format' in query else []
        result = self.service.enrich(record())
        self.assertEqual(result['Business Email ID'], 'jd742@example.com')
        self.assertEqual(result['Email Status'], FOUND)
        self.assertIn('source page not verified', result['_email_evidence'][0]['method'])
        self.assertEqual(result['_email_evidence'][0]['url'], search_row(source)['search_url'])

    def test_pattern_search_conflicting_published_addresses_do_not_generate(self):
        sources = [document('Example Industries email format: {first}.{last}@example.com\n'
                            'Jane Doe\nEmail: ' + address,
                            'Example Industries Email Format', f'https://example.com/email-format/{index}')
                   for index, address in enumerate(('jd742@example.com', 'jd915@example.com'))]
        self.web.search.side_effect = lambda query: ([search_row(source, text='Example Industries email format')
                                                     for source in sources] if 'email format' in query else [])
        self.web.fetch.side_effect = lambda url: next(source for source in sources if source.url == url)
        result = self.service.enrich(record())
        self.assertEqual(result['Business Email ID'], '-')
        self.assertEqual(result['Email Status'], 'Ambiguous (unverified)')

    def test_pattern_cache_never_skips_published_search_for_second_employee(self):
        self.service.pattern_cache[('example industries', 'example.com')] = {
            'pattern': '{first}.{last}', 'domain': 'example.com', 'url': 'https://directory.example/pattern',
            'excerpt': 'Company pattern', 'method': 'Published company email pattern'}
        official = document('Jane Doe\nEmail: jd742@example.com')
        self.web.search.return_value = [search_row(official)]
        self.web.fetch.return_value = official
        self.assertEqual(self.service.enrich(record())['Business Email ID'], 'jd742@example.com')

    def test_invalid_website_fails_closed_without_search(self):
        for website in ('https://[invalid', 'http://127.0.0.1', 'javascript:alert(1)'):
            with self.subTest(website=website):
                result = self.service.enrich(record(**{'Company Website': website}))
                self.assertEqual(result['Email Status'], 'Not found')
        self.web.search.assert_not_called()

    def test_lookup_failure_does_not_mutate_manual_record(self):
        before = record(**{'Business Email ID': 'manual@example.com', 'Email Status': 'Reviewed',
                           'Bounce Email Status': 'Hard bounce', 'Notes': 'Manual review'})
        snapshot = copy.deepcopy(before)
        result = self.service.enrich(before)
        self.assertEqual(result, snapshot)
        self.assertEqual(before, snapshot)
        self.web.search.assert_not_called()


if __name__ == '__main__':
    unittest.main()
