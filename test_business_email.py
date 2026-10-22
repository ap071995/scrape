"""Offline regressions for evidence-first, person-specific work email enrichment."""
import copy
import re
import unittest
from unittest.mock import Mock

from company_public_data import SourceDocument
from business_email import (
    BusinessEmailEnricher, email_candidates, pattern_evidence,
    published_emails, select_pattern,
)
from business_email import valid_email, inferred_patterns


def company_match(expected, actual):
    def normalized(value):
        value = re.sub(r"\b(?:incorporated|inc|limited|ltd|corporation|corp)\b", "", value.lower())
        return re.sub(r"[^a-z0-9]", "", value)
    return bool(normalized(expected)) and normalized(expected) == normalized(actual)


def record(**updates):
    row = {
        'First Name': 'Jane', 'Last Name': 'Doe',
        'Company Name': 'Example Industries', 'Company Website': 'https://example.com',
        'Prospect LinkedIn Link': 'https://www.linkedin.com/in/jane-doe',
        'Business Email ID': '-', 'Email Status': '-', 'Notes': 'Keep manual note',
    }
    row.update(updates)
    return row


def document(text, title='Jane Doe | Example Industries', url='https://example.com/staff/jane-doe'):
    return SourceDocument(url, title, text, headings=[title.split(' | ')[0]])


def search_row(doc, text=None):
    return {
        'url': doc.url, 'title': doc.title, 'text': text if text is not None else doc.text,
        'provider': 'google', 'kind': 'organic',
        'search_url': 'https://www.google.com/search?q=Jane+Doe+Example+Industries+email',
    }


class PublishedEmailTests(unittest.TestCase):
    def test_negated_and_example_only_email_is_not_published(self):
        for text in ('Jane Doe email is not jane.doe@example.com.',
                     'Jane Doe\nEmail: jane.doe@example.com (example only)',
                     'Jane Doe\nDo not contact jane.doe@example.com'):
            for kind in ('page', 'search'):
                with self.subTest(text=text, kind=kind):
                    self.assertEqual(published_emails(document(text), record(), kind), [])

    def test_assistant_contact_is_not_assigned_to_target(self):
        for text in ('Jane Doe\nAssistant contact: Bob Smith\nEmail: bob.smith@example.com',
                     'Jane Doe\nPA: bob.smith@example.com',
                     'Executive assistant: jane.doe@example.com'):
            for kind in ('page', 'search'):
                with self.subTest(text=text, kind=kind):
                    self.assertEqual(published_emails(document(text), record(), kind), [])

    def test_person_whose_own_role_is_assistant_is_accepted(self):
        for title in ('Assistant Professor', 'Assistant Director', 'Executive Assistant'):
            facts = published_emails(document('Jane Doe\n' + title + '\nEmail: jane.doe@example.com'), record())
            self.assertEqual(facts[0]['email'], 'jane.doe@example.com')

    def test_named_next_employee_with_role_is_not_target(self):
        doc = document('Jane Doe\nDirector\nJohn Smith, Manager\nEmail: john.smith@example.com')
        self.assertEqual(published_emails(doc, record()), [])

    def test_generated_example_is_not_a_published_person_email(self):
        doc = document('Jane Doe email is generated from the format first.last@example.com. Example: jane.doe@example.com')
        self.assertEqual(published_emails(doc, record()), [])
        self.assertEqual(published_emails(doc, record(), 'search'), [])

    def test_malformed_email_is_rejected(self):
        for value in ('jane.@example.com', 'jane@a.-example.com', 'jane@example-.com'):
            self.assertFalse(valid_email(value))

    def test_named_official_contact_can_establish_alternate_corporate_domain(self):
        doc = document('Jane Doe\nEmail: jd742@examplemail.com')
        self.assertEqual(published_emails(doc, record())[0]['email'], 'jd742@examplemail.com')

    def test_exact_person_email_on_official_profile(self):
        facts = published_emails(document('Jane Doe\nDirector\nEmail: jane.doe@example.com'), record())
        self.assertEqual({fact['email'] for fact in facts}, {'jane.doe@example.com'})
        self.assertEqual(facts[0]['url'], 'https://example.com/staff/jane-doe')
        self.assertIn('jane.doe@example.com', facts[0]['excerpt'])

    def test_email_local_part_need_not_be_guessed_name_pattern(self):
        facts = published_emails(document('Jane Doe\nEmail: jd742@example.com'), record())
        self.assertEqual({fact['email'] for fact in facts}, {'jd742@example.com'})

    def test_other_person_profile_is_rejected(self):
        doc = document('John Doe\nEmail: john.doe@example.com', 'John Doe | Example Industries')
        self.assertEqual(published_emails(doc, record()), [])

    def test_team_listing_does_not_assign_next_person_email(self):
        doc = document('Jane Doe\nDirector\n\nJohn Smith\nManager\nEmail: john.smith@example.com',
                       'Our Team | Example Industries', 'https://example.com/team')
        facts = published_emails(doc, record())
        self.assertNotIn('john.smith@example.com', {fact['email'] for fact in facts})

    def test_generic_footer_email_is_not_a_person_email(self):
        doc = document('Jane Doe\nDirector\nContact our team\ninfo@example.com\nsupport@example.com')
        self.assertEqual(published_emails(doc, record()), [])

    def test_personal_mailbox_is_not_business_email(self):
        doc = document('Jane Doe\nEmail: jane.doe@gmail.com\nEmail: jane.doe@hotmail.com')
        self.assertEqual(published_emails(doc, record()), [])

    def test_conflicting_company_domain_is_rejected(self):
        doc = document('Jane Doe\nEmail: jane.doe@othercompany.com',
                       'Jane Doe | Example Industries', 'https://directory.example/jane-doe')
        self.assertEqual(published_emails(doc, record()), [])

    def test_obfuscated_public_address_is_decoded(self):
        for email_text in ['jane.doe [at] example [dot] com', 'jane.doe (at) example (dot) com']:
            with self.subTest(email_text=email_text):
                facts = published_emails(document('Jane Doe\nEmail: ' + email_text), record())
                self.assertEqual({fact['email'] for fact in facts}, {'jane.doe@example.com'})

    def test_masked_address_is_not_repaired_by_guessing(self):
        doc = document('Jane Doe\nEmail: j***@example.com\nEmail: jane.***@example.com')
        self.assertEqual(published_emails(doc, record()), [])

    def test_near_name_in_article_does_not_assign_media_contact(self):
        doc = document('Jane Doe leads the new project.\n\nMedia contact\nJohn Smith\nEmail: john.smith@example.com',
                       'Project announcement | Example Industries', 'https://example.com/news/project')
        self.assertEqual(published_emails(doc, record()), [])


class CompanyPatternTests(unittest.TestCase):
    def test_official_blog_cannot_supply_another_company_pattern(self):
        doc = document('Other Industries email format: {first}.{last}@othercompany.com',
                       'Blog | Example Industries', 'https://example.com/blog')
        self.assertEqual(pattern_evidence(doc, record(), company_match), [])

    def test_two_named_official_employees_establish_unique_pattern(self):
        doc = document('John Smith\nEmail: john.smith@example.com\nAlice Brown\nEmail: alice.brown@example.com',
                       'Our Team | Example Industries', 'https://example.com/team')
        self.assertEqual(inferred_patterns(doc, record())[0]['pattern'], '{first}.{last}')

    def test_conflicting_employee_patterns_are_not_inferred(self):
        doc = document('John Smith\nEmail: john.smith@example.com\nAlice Brown\nEmail: abrown@example.com',
                       'Our Team | Example Industries', 'https://example.com/team')
        self.assertEqual(inferred_patterns(doc, record()), [])

    def test_candidate_formats_include_requested_examples(self):
        candidates = email_candidates('Jane', 'Doe', 'example.com')
        self.assertEqual(candidates['{first}.{last}'], 'jane.doe@example.com')
        self.assertEqual(candidates['{last}_{f}'], 'doe_j@example.com')

    def test_missing_name_does_not_create_address(self):
        self.assertEqual(email_candidates('-', 'Doe', 'example.com'), {})
        self.assertEqual(email_candidates('Jane', '-', 'example.com'), {})

    def test_free_mail_domain_does_not_create_address(self):
        self.assertEqual(email_candidates('Jane', 'Doe', 'gmail.com'), {})

    def test_explicit_company_pattern_is_accepted(self):
        doc = document('Example Industries email format: {first}.{last}@example.com',
                       'Example Industries Email Format', 'https://directory.example/example-industries')
        facts = pattern_evidence(doc, record(), company_match)
        self.assertIn('{first}.{last}', {fact['pattern'] for fact in facts})

    def test_wrong_company_pattern_is_rejected(self):
        doc = document('Other Industries email format: {first}.{last}@othercompany.com',
                       'Other Industries Email Format', 'https://directory.example/other-industries')
        self.assertEqual(pattern_evidence(doc, record(), company_match), [])

    def test_unsubstantiated_example_is_not_company_pattern(self):
        doc = document('Contact us. For example, use firstname.lastname@example.com.',
                       'How to write a professional email', 'https://writing.example/email-guide')
        self.assertEqual(pattern_evidence(doc, record(), company_match), [])

    def test_conflicting_patterns_are_not_arbitrarily_selected(self):
        evidence = [
            {'pattern': pattern, 'domain': 'example.com', 'url': f'https://directory.example/{index}',
             'excerpt': 'Example Industries email format: ' + pattern + '@example.com'}
            for index, pattern in enumerate(['{first}.{last}', '{f}{last}'])
        ]
        self.assertIsNone(select_pattern(evidence))

    def test_agreeing_pattern_sources_return_one_supported_pattern(self):
        evidence = [
            {'pattern': '{last}_{f}', 'domain': 'example.com', 'url': f'https://directory.example/{index}',
             'excerpt': 'Example Industries email format: {last}_{f}@example.com'}
            for index in (1, 2)
        ]
        selected = select_pattern(evidence)
        self.assertIsNotNone(selected)
        self.assertEqual(selected['pattern'], '{last}_{f}')


class BusinessEmailWorkflowTests(unittest.TestCase):
    def setUp(self):
        self.web = Mock()
        self.web.search.return_value = []
        self.web.fetch.return_value = None
        self.service = BusinessEmailEnricher(self.web, company_match)

    def use_document(self, doc):
        self.web.search.return_value = [search_row(doc)]
        self.web.fetch.side_effect = lambda url: doc if url == doc.url else None

    def test_bare_website_is_usable_without_changing_original_cell(self):
        self.use_document(document('Jane Doe\nEmail: jane.doe@example.com'))
        result = self.service.enrich(record(**{'Company Website':'example.com'}))
        self.assertEqual(result['Business Email ID'], 'jane.doe@example.com')
        self.assertEqual(result['Company Website'], 'example.com')

    def test_same_address_promotion_preserves_bounce_status(self):
        self.use_document(document('Jane Doe\nEmail: jane.doe@example.com'))
        result = self.service.enrich(record(**{'Business Email ID':'jane.doe@example.com',
            'Email Status':'Generated (unverified)', 'Bounce Email Status':'Hard bounce'}))
        self.assertEqual(result['Email Status'], 'Found online (unverified)')
        self.assertEqual(result['Bounce Email Status'], 'Hard bounce')

    def test_official_directory_fallback_finds_address_before_generating(self):
        doc = document('Jane Doe\nEmail: jd742@example.com')
        self.web.source_index = {'example.com': {'company': 'Example Industries', 'email_urls': [doc.url]}}
        self.web.fetch.side_effect = lambda url: doc if url == doc.url else None
        result = self.service.enrich(record())
        self.assertEqual(result['Business Email ID'], 'jd742@example.com')
        self.assertEqual(result['Email Status'], 'Found online (unverified)')
        self.assertEqual(result['_email_evidence'][0]['url'], doc.url)
        self.assertFalse(any('format' in call.args[0] for call in self.web.search.call_args_list))

    def test_directory_fallback_checks_company_and_official_domain(self):
        for company, url in [('Other Company', 'https://example.com/staff'),
                             ('Example Industries', 'https://othercompany.com/staff')]:
            with self.subTest(company=company, url=url):
                self.web.source_index = {'example.com': {'company': company, 'email_urls': [url]}}
                self.service.cache.clear()
                self.service.enrich(record())
                self.web.fetch.assert_not_called()

    def test_cached_official_page_is_reused_for_person_lookup(self):
        doc = document('Jane Doe\nEmail: jd742@example.com')
        self.web.documents = {doc.url: doc}
        result = self.service.enrich(record())
        self.assertEqual(result['Business Email ID'], 'jd742@example.com')
        self.web.fetch.assert_not_called()

    def test_known_company_source_can_supply_pattern_when_search_misses_it(self):
        doc = document('Example Industries email format: {last}_{f}@example.com',
                       'Example Industries Email Format', 'https://directory.example/company')
        self.web.source_index = {'example.com': {'company': 'Example Industries', 'urls': [doc.url]}}
        self.web.fetch.side_effect = lambda url: doc if url == doc.url else None
        result = self.service.enrich(record())
        self.assertEqual(result['Business Email ID'], 'doe_j@example.com')
        self.assertEqual(result['Email Status'], 'Generated (unverified)')

    def test_person_email_on_seeded_company_page_precedes_its_pattern(self):
        doc = document('Jane Doe\nEmail: jd742@example.com\n'
                       'Example Industries email format: {first}.{last}@example.com',
                       'Example Industries Email Format', 'https://directory.example/company')
        self.web.source_index = {'example.com': {'company': 'Example Industries', 'urls': [doc.url]}}
        self.web.fetch.side_effect = lambda url: doc if url == doc.url else None
        result = self.service.enrich(record())
        self.assertEqual(result['Business Email ID'], 'jd742@example.com')
        self.assertEqual(result['Email Status'], 'Found online (unverified)')

    def test_found_address_precedes_generation_and_records_evidence(self):
        self.use_document(document('Jane Doe\nEmail: jane.doe@example.com'))
        before = record()
        snapshot = copy.deepcopy(before)
        result = self.service.enrich(before)
        self.assertEqual(result['Business Email ID'], 'jane.doe@example.com')
        self.assertEqual(result['Email Status'], 'Found online (unverified)')
        self.assertEqual(before, snapshot)
        self.assertIn('Keep manual note', result['Notes'])
        self.assertTrue(result.get('_email_evidence'))
        evidence = result['_email_evidence'][0]
        for key in ('person', 'profile', 'company', 'email', 'method', 'pattern', 'url', 'excerpt', 'checked_at'):
            self.assertIn(key, evidence)
        queries = [call.args[0].lower() for call in self.web.search.call_args_list]
        self.assertFalse(any('pattern' in query or 'format' in query for query in queries))

    def test_visible_search_address_survives_unavailable_destination(self):
        doc = document('Jane Doe - Example Industries\nEmail: jane.doe@example.com')
        self.web.search.return_value = [search_row(doc)]
        result = self.service.enrich(record())
        self.assertEqual(result['Business Email ID'], 'jane.doe@example.com')
        self.assertEqual(result['Email Status'], 'Found online (unverified)')

    def test_generated_address_requires_published_company_pattern(self):
        doc = document('Example Industries email format: {last}_{f}@example.com',
                       'Example Industries Email Format', 'https://directory.example/example-industries')
        self.web.search.side_effect = lambda query: [search_row(doc)] if re.search(r'pattern|format', query, re.I) else []
        self.web.fetch.side_effect = lambda url: doc if url == doc.url else None
        result = self.service.enrich(record())
        self.assertEqual(result['Business Email ID'], 'doe_j@example.com')
        self.assertEqual(result['Email Status'], 'Generated (unverified)')
        self.assertEqual(result['_email_evidence'][0]['pattern'], '{last}_{f}')

    def test_no_source_no_guessed_default(self):
        result = self.service.enrich(record())
        self.assertIn(result['Business Email ID'], ('', '-'))
        self.assertEqual(result['Email Status'], 'Not found')

    def test_existing_valid_address_and_status_are_preserved(self):
        row = record(**{'Business Email ID': 'manual@example.com', 'Email Status': 'Manually reviewed'})
        result = self.service.enrich(row)
        self.assertEqual(result['Business Email ID'], 'manual@example.com')
        self.assertEqual(result['Email Status'], 'Manually reviewed')
        self.web.search.assert_not_called()

    def test_generated_address_is_retried_and_promoted_when_published(self):
        self.use_document(document('Jane Doe\nEmail: jd742@example.com'))
        result = self.service.enrich(record(**{
            'Business Email ID': 'jane.doe@example.com', 'Email Status': 'Generated (unverified)',
        }))
        self.assertEqual(result['Business Email ID'], 'jd742@example.com')
        self.assertEqual(result['Email Status'], 'Found online (unverified)')

    def test_two_person_addresses_are_not_silently_selected(self):
        self.use_document(document('Jane Doe\nEmail: jane.doe@example.com\nEmail: jdoe@example.com'))
        result = self.service.enrich(record())
        self.assertIn(result['Business Email ID'], ('', '-'))
        self.assertEqual(result['Email Status'], 'Ambiguous (unverified)')

    def test_person_cache_does_not_reuse_a_colleagues_email(self):
        jane = document('Jane Doe\nEmail: jane.doe@example.com')
        john = document('John Smith\nEmail: john.smith@example.com',
                        'John Smith | Example Industries', 'https://example.com/staff/john-smith')
        self.web.search.side_effect = lambda query: [search_row(john if 'John' in query else jane)]
        self.web.fetch.side_effect = lambda url: {jane.url: jane, john.url: john}.get(url)
        first = self.service.enrich(record())
        second = self.service.enrich(record(**{
            'First Name': 'John', 'Last Name': 'Smith',
            'Prospect LinkedIn Link': 'https://www.linkedin.com/in/john-smith',
        }))
        self.assertEqual(first['Business Email ID'], 'jane.doe@example.com')
        self.assertEqual(second['Business Email ID'], 'john.smith@example.com')


if __name__ == '__main__':
    unittest.main()
