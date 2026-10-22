"""Directory wording and table regressions for evidenced email formats."""
import unittest

from business_email import pattern_evidence, select_pattern
from company_public_data import SourceDocument
from test_business_email import company_match


class PublishedPatternWordingTests(unittest.TestCase):
    def setUp(self):
        self.record = {'Company Name': 'East Carolina University', 'Company Website': 'https://ecu.edu'}

    def parse(self, text, title='East Carolina University Email Format', url='https://directory.example/ecu'):
        return pattern_evidence(SourceDocument(url, title, text), self.record, company_match)

    def test_possessive_intro_prioritizes_primary_not_also_uses(self):
        facts = self.parse("The most accurate and popular East Carolina University's email format is LastF@ecu.edu "
                           "(ex. DoeJ@ecu.edu). East Carolina University also uses FirstL@ecu.edu and FLast@ecu.edu.")
        self.assertEqual({item['pattern'] for item in facts}, {'{last}{f}', '{first}{l}', '{f}{last}'})
        self.assertEqual(select_pattern(facts)['pattern'], '{last}{f}')
        self.assertEqual(sum(item['preferred'] for item in facts), 1)

    def test_possessive_company_title_is_accepted(self):
        facts = self.parse("East Carolina University's primary email format is LastF@ecu.edu.",
                           "East Carolina University's Email Format")
        self.assertEqual(select_pattern(facts)['pattern'], '{last}{f}')

    def test_most_common_percentage_table_prefers_largest_share(self):
        facts = self.parse('Most Common Email Formats at East Carolina University\n'
                           'LastF@ecu.edu 78%\nFirstL@ecu.edu 15%\nFLast@ecu.edu 4%',
                           'Most Common Email Formats at East Carolina University')
        self.assertEqual(len(facts), 3)
        self.assertEqual(select_pattern(facts)['pattern'], '{last}{f}')
        self.assertIn('78%', select_pattern(facts)['excerpt'])

    def test_table_split_percentage_line(self):
        facts = self.parse('Most Common Email Formats at East Carolina University\n'
                           'LastF@ecu.edu\n78%\nFirstL@ecu.edu\n15%')
        self.assertEqual(select_pattern(facts)['pattern'], '{last}{f}')

    def test_percentage_tie_remains_ambiguous(self):
        facts = self.parse('Most Common Email Formats at East Carolina University\n'
                           'LastF@ecu.edu 50%\nFirstL@ecu.edu 50%')
        self.assertIsNone(select_pattern(facts))

    def test_unlabelled_percentages_do_not_establish_primary(self):
        facts = self.parse('LastF@ecu.edu 78%\nFirstL@ecu.edu 15%')
        self.assertIsNone(select_pattern(facts))

    def test_most_common_company_possessive_initial_alias(self):
        self.record = {'Company Name': '3M', 'Company Website': 'https://mmm.com'}
        facts = self.parse("3M's most common email format is [first_initial][last] (ex. jdoe@mmm.com)",
                           '3M Email Format')
        self.assertEqual(select_pattern(facts)['pattern'], '{f}{last}')
        self.assertTrue(select_pattern(facts)['preferred'])

    def test_curly_possessive(self):
        facts = self.parse('East Carolina University\u2019s standard email format is LastF@ecu.edu.')
        self.assertEqual(select_pattern(facts)['pattern'], '{last}{f}')

    def test_dom_spacing_between_explicit_tokens(self):
        for value, expected in (('[first_initial] [last]', '{f}{last}'), ('[first]. [last]', '{first}.{last}')):
            with self.subTest(value=value):
                facts = self.parse('East Carolina University email format is ' + value + ' (ex. jdoe@ecu.edu).')
                self.assertEqual(select_pattern(facts)['pattern'], expected)

    def test_table_subheader_preserves_ranking_context(self):
        facts = self.parse('Most Common Email Formats at East Carolina University\n'
                           'East Carolina University Email Formats\tExample\tPercentage\n'
                           'LastF@ecu.edu\tDoeJ@ecu.edu\n78%\nFirstL@ecu.edu\tJohnD@ecu.edu\n15%')
        self.assertEqual(select_pattern(facts)['pattern'], '{last}{f}')

    def test_conflicting_explicit_primaries_remain_ambiguous(self):
        facts = self.parse("East Carolina University's primary email format is LastF@ecu.edu. "
                           "East Carolina University's most common email format is FirstL@ecu.edu.")
        self.assertEqual(sum(item['preferred'] for item in facts), 2)
        self.assertIsNone(select_pattern(facts))

    def test_example_address_alone_is_not_pattern(self):
        self.assertEqual(self.parse('The most common email format is shown by DoeJ@ecu.edu.'), [])

    def test_third_party_pattern_must_match_company_domain(self):
        self.assertEqual(self.parse('East Carolina University email format is LastF@other.edu.'), [])

    def test_other_company_subject_rejected_on_official_page(self):
        self.assertEqual(self.parse("Other University\u2019s most common email format is LastF@other.edu.",
                                   'Our Blog', 'https://ecu.edu/blog'), [])

    def test_other_company_table_rejected_on_official_page(self):
        self.assertEqual(self.parse('Most Common Email Formats at Other University\nLastF@other.edu 80%',
                                   'Our Blog', 'https://ecu.edu/blog'), [])

    def test_other_company_also_uses_statement_rejected(self):
        facts = self.parse('East Carolina University email format: LastF@ecu.edu. '
                           'Other University also uses First.Last@other.edu.',
                           'Our Blog', 'https://ecu.edu/blog')
        self.assertEqual({item['domain'] for item in facts}, {'ecu.edu'})

    def test_explicitly_invalid_pattern_is_rejected(self):
        self.assertEqual(self.parse('East Carolina University email format is not First.Last@ecu.edu.'), [])
        self.assertEqual(self.parse('East Carolina University no longer uses First.Last@ecu.edu.'), [])

    def test_bracketed_row_cannot_escape_foreign_heading_scope(self):
        text = 'Most Common Email Formats at Other University\n[first].[last]@other.edu'
        self.assertEqual(self.parse(text, url='https://ecu.edu/blog'), [])
        facts = self.parse(text + '\nEast Carolina University email format: LastF@ecu.edu\nFirstL@ecu.edu',
                           url='https://ecu.edu/blog')
        self.assertEqual({item['domain'] for item in facts}, {'ecu.edu'})
        self.assertEqual({item['pattern'] for item in facts}, {'{last}{f}', '{first}{l}'})

    def test_parenthetical_foreign_heading_clears_target_scope(self):
        facts = self.parse('Most Common Email Formats at East Carolina University\n'
                           'Other University Email Formats (2026)\nFirst.Last@other.edu\n'
                           'East Carolina University Email Formats\nLastF@ecu.edu',
                           url='https://ecu.edu/blog')
        self.assertEqual({item['domain'] for item in facts}, {'ecu.edu'})
        self.assertEqual({item['pattern'] for item in facts}, {'{last}{f}'})


if __name__ == '__main__':
    unittest.main()
