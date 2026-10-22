"""Offline domain/DNS checks: no mailbox probing and no live resolver traffic."""
import copy
import json
import unittest
from unittest.mock import Mock, patch

import requests

from business_email import build_email_details, BusinessEmailEnricher, FOUND, GENERATED
from email_domain_check import DomainMXChecker, normalized_domain, RESOLVER_URL
from test_business_email import record, company_match


def answer(kind=15, data='10 mail.example.com.', ttl=120):
    return {'Status': 0, 'Answer': [{'type': kind, 'data': data, 'TTL': ttl}]}


class EmailDetailsTests(unittest.TestCase):
    def test_published_and_generated_are_separate_and_do_not_mutate_lead(self):
        for status, field in ((FOUND, 'published_email'), (GENERATED, 'generated_email'),
                              ('Verified by user', 'existing_email'), ('-', 'existing_email')):
            with self.subTest(status=status):
                lead = record(**{'Business Email ID': 'Jane.Doe@example.com', 'Email Status': status,
                                 'Bounce Email Status': 'Keep this value'})
                snapshot = copy.deepcopy(lead)
                details = build_email_details(lead)
                self.assertEqual(details[field], 'Jane.Doe@example.com')
                self.assertEqual(sum(bool(details[key]) for key in
                                     ('published_email', 'generated_email', 'existing_email')), 1)
                self.assertEqual(details['mailbox_status'], 'Not verified')
                self.assertEqual(details['profile'], lead['Prospect LinkedIn Link'])
                self.assertEqual(details['company'], lead['Company Name'])
                self.assertEqual(lead, snapshot)

    def test_disabled_means_no_network_and_no_checked_date(self):
        with patch('email_domain_check.requests.get') as network:
            details = build_email_details(record(**{'Business Email ID': 'jane@example.com'}))
        network.assert_not_called()
        self.assertEqual(details['domain_status'], 'Not checked (disabled)')
        self.assertEqual(details['checked_at'], '')

    def test_no_email_means_no_dns_even_when_enabled(self):
        checker = Mock()
        for value in ('-', '', None, 'jane@', 'bad address@example.com'):
            details = build_email_details(record(**{'Business Email ID': value}), checker)
            self.assertEqual(details['domain_status'], 'No valid email')
            self.assertEqual(details['email_domain'], '')
        checker.check.assert_not_called()

    def test_checks_actual_email_domain_without_name_or_local_part(self):
        query = Mock(return_value=answer(data='10 mx.other.org.'))
        checker = DomainMXChecker(query=query)
        lead = record(**{'Business Email ID': 'jane@staff.other.org', 'Email Status': FOUND})
        details = build_email_details(lead, checker)
        self.assertEqual(query.call_args.args[:2], ('staff.other.org', 'MX'))
        self.assertEqual(details['domain_status'], 'MX records found')
        self.assertTrue(details['checked_at'])
        self.assertEqual(details['mailbox_status'], 'Not verified')
        self.assertEqual(details['published_email'], lead['Business Email ID'])

    def test_negative_dns_never_clears_or_reclassifies_email(self):
        for payload in ({'Status': 3}, answer(data='0 .')):
            for status, key in ((FOUND, 'published_email'), (GENERATED, 'generated_email')):
                details = build_email_details(record(**{'Business Email ID': 'jane@example.com',
                                                         'Email Status': status}),
                                              DomainMXChecker(query=Mock(return_value=payload)))
                self.assertEqual(details[key], 'jane@example.com')
                self.assertEqual(details['mailbox_status'], 'Not verified')

    def test_invalid_public_domain_is_reported_without_network(self):
        query = Mock()
        details = build_email_details(record(**{'Business Email ID': 'jane@company.local'}),
                                      DomainMXChecker(query=query))
        self.assertEqual(details['domain_status'], 'Invalid domain')
        query.assert_not_called()

    def test_optional_checker_failure_or_bad_result_does_not_lose_lead(self):
        for bad in (None, {}, {'domain': 'unrelated.org', 'status': 'MX records found',
                              'checked_at': 'now', 'detail': 'Wrong domain'},
                    {'domain': 'example.com', 'status': None, 'checked_at': 'now', 'detail': 'bad'}):
            details = build_email_details(record(**{'Business Email ID': 'jane@example.com'}),
                                          Mock(check=Mock(return_value=bad)))
            self.assertEqual(details['domain_status'], 'Inconclusive (DNS lookup failed)')
            self.assertEqual(details['existing_email'], 'jane@example.com')
        details = build_email_details(record(**{'Business Email ID': 'jane@example.com'}),
                                      Mock(check=Mock(side_effect=RuntimeError('resolver failed'))))
        self.assertEqual(details['mailbox_status'], 'Not verified')
        self.assertEqual(details['existing_email'], 'jane@example.com')

    def test_business_enricher_preserves_existing_public_return_shape(self):
        lead = record(**{'Business Email ID': 'jane@example.com', 'Email Status': FOUND})
        enricher = BusinessEmailEnricher(Mock(), Mock(), company_match)
        self.assertEqual(enricher.enrich(lead), lead)
        self.assertNotIn('_email_details', enricher.enrich(record(**{'First Name': '-'})))


class DomainMXTests(unittest.TestCase):
    def test_public_domain_normalization_and_idna(self):
        self.assertEqual(normalized_domain('EXAMPLE.COM.'), 'example.com')
        self.assertEqual(normalized_domain('bücher.de'), 'xn--bcher-kva.de')
        for invalid in ('', None, 3, 'user@example.com', 'https://example.com',
                        'example.com/path', ' example.com', 'example.com ', 'example..com',
                        '.example.com', 'example.com..', 'bad_.com', '-bad.com', 'bad-.com',
                        'a' * 64 + '.com', '127.0.0.1', '::1', 'corp.local', 'corp.internal',
                        'localhost', 'corp.home.arpa', 'corp.invalid', 'corp.test', 'corp.onion',
                        '127.1', '0x7f.0.0.1', 'hello.123', 'example.com\n'):
            with self.subTest(value=invalid):
                self.assertEqual(normalized_domain(invalid), '')

    def test_invalid_domain_avoids_resolver(self):
        query = Mock()
        checker = DomainMXChecker(query=query)
        self.assertEqual(checker.check('localhost')['status'], 'Invalid domain')
        query.assert_not_called()
        self.assertFalse(checker.cache)

    def test_mx_stops_after_one_query_and_never_probes_exchange(self):
        query = Mock(return_value=answer())
        result = DomainMXChecker(query=query).check('example.com')
        self.assertEqual(result['status'], 'MX records found')
        self.assertEqual(query.call_count, 1)
        self.assertEqual(query.call_args.args[:2], ('example.com', 'MX'))
        self.assertIn('not verified', result['detail'])

    def test_nxdomain_is_separate_from_timeout_and_does_not_fallback(self):
        query = Mock(return_value={'Status': 3})
        self.assertEqual(DomainMXChecker(query=query).check('example.com')['status'],
                         'Domain does not exist (NXDOMAIN)')
        self.assertEqual(query.call_count, 1)

    def test_null_mx_does_not_fallback_to_web_server(self):
        query = Mock(return_value=answer(data='0 .'))
        self.assertEqual(DomainMXChecker(query=query).check('example.com')['status'],
                         'Does not accept mail (null MX)')
        self.assertEqual(query.call_count, 1)

    def test_conflicting_and_nonzero_null_mx_are_inconclusive(self):
        mixed = answer(data='0 .')
        mixed['Answer'].extend(answer()['Answer'])
        for payload in (mixed, answer(data='10 .')):
            self.assertEqual(DomainMXChecker(query=Mock(return_value=payload)).check('example.com')['status'],
                             'Inconclusive (invalid MX configuration)')

    def test_no_mx_with_a_uses_implicit_fallback_in_two_queries(self):
        query = Mock(side_effect=[{'Status': 0}, answer(1, '93.184.216.34')])
        result = DomainMXChecker(query=query).check('example.com')
        self.assertEqual(result['status'], 'No MX; A/AAAA fallback available')
        self.assertEqual([call.args[1] for call in query.call_args_list], ['MX', 'A'])

    def test_ipv6_only_implicit_fallback(self):
        query = Mock(side_effect=[{'Status': 0}, {'Status': 0}, answer(28, '2606:4700::1111')])
        self.assertEqual(DomainMXChecker(query=query).check('example.com')['status'],
                         'No MX; A/AAAA fallback available')
        self.assertEqual([call.args[1] for call in query.call_args_list], ['MX', 'A', 'AAAA'])

    def test_no_dns_route_is_distinct_from_failed_lookup(self):
        query = Mock(return_value={'Status': 0})
        self.assertEqual(DomainMXChecker(query=query).check('example.com')['status'],
                         'No MX or A/AAAA records')
        self.assertEqual(query.call_count, 3)

    def test_temporary_dns_errors_never_mean_missing_domain(self):
        for payload in ({'Status': 2}, {'Status': 5}, {'Status': True}, {'Status': 0, 'TC': True},
                        {'Status': 0, 'Answer': None}, {'Status': 0, 'Answer': ['bad']},
                        {'Status': 3, 'Answer': answer()['Answer']}, [], {}):
            with self.subTest(payload=payload):
                query = Mock(return_value=payload)
                result = DomainMXChecker(query=query).check('example.com')
                self.assertEqual(result['status'], 'Inconclusive (DNS lookup failed)')
                self.assertEqual(query.call_count, 1)

    def test_malformed_mx_never_gets_positive_verdict(self):
        for value in ('mail.example.com', '65536 mail.example.com.', '1 localhost',
                      '1 127.0.0.1', '1 https://mail.example.com', 'xx mail.example.com'):
            result = DomainMXChecker(query=Mock(return_value=answer(data=value))).check('example.com')
            self.assertEqual(result['status'], 'Inconclusive (DNS lookup failed)')

    def test_timeout_is_separate_in_mx_and_fallback_paths(self):
        for exception in (TimeoutError('deadline'), requests.Timeout('network timeout')):
            result = DomainMXChecker(query=Mock(side_effect=exception)).check('example.com')
            self.assertEqual(result['status'], 'Inconclusive (DNS timeout)')
            query = Mock(side_effect=[{'Status': 0}, exception, {'Status': 0}])
            self.assertEqual(DomainMXChecker(query=query).check('example.com')['status'],
                             'Inconclusive (DNS timeout)')

    def test_fallback_can_succeed_when_other_address_family_fails(self):
        query = Mock(side_effect=[{'Status': 0}, requests.Timeout(), answer(28, '2606:4700::1111')])
        self.assertEqual(DomainMXChecker(query=query).check('example.com')['status'],
                         'No MX; A/AAAA fallback available')

    def test_inconsistent_nxdomain_and_malformed_address_are_inconclusive(self):
        for middle in ({'Status': 3}, answer(1, 'not-an-address'), answer(1, '::1')):
            query = Mock(side_effect=[{'Status': 0}, middle, {'Status': 0}])
            self.assertEqual(DomainMXChecker(query=query).check('example.com')['status'],
                             'Inconclusive (DNS lookup failed)')

    def test_connection_failure_does_not_escape_or_retry(self):
        query = Mock(side_effect=requests.ConnectionError('offline'))
        self.assertEqual(DomainMXChecker(query=query).check('example.com')['status'],
                         'Inconclusive (DNS lookup failed)')
        self.assertEqual(query.call_count, 1)

    def test_cache_shared_by_normalized_domain_and_returns_copies(self):
        query = Mock(return_value=answer())
        checker = DomainMXChecker(query=query)
        first = checker.check('EXAMPLE.COM.')
        date = first['checked_at']
        first['status'] = 'poisoned'
        second = checker.check('example.com')
        self.assertEqual(second['status'], 'MX records found')
        self.assertEqual(second['checked_at'], date)
        self.assertEqual(query.call_count, 1)

    def test_positive_cache_expires_at_ttl_and_caps_at_one_hour(self):
        for ttl, lifetime in ((120, 120), (7200, 3600), (0, 0)):
            now = [10.0]
            query = Mock(return_value=answer(ttl=ttl))
            checker = DomainMXChecker(query=query, clock=lambda: now[0])
            checker.check('example.com')
            self.assertEqual(checker.cache['example.com'][0], 10.0 + lifetime)
            now[0] += lifetime
            checker.check('example.com')
            self.assertEqual(query.call_count, 2)

    def test_negative_and_transient_cache_expire_for_retry(self):
        for payload, lifetime in (({'Status': 3}, 300), ({'Status': 2}, 60),
                                  (answer(data='0 .', ttl=180), 180)):
            now = [0.0]
            query = Mock(return_value=payload)
            checker = DomainMXChecker(query=query, clock=lambda: now[0])
            checker.check('example.com')
            self.assertEqual(checker.cache['example.com'][0], lifetime)
            now[0] = lifetime - 0.1
            checker.check('example.com')
            self.assertEqual(query.call_count, 1)
            now[0] = lifetime
            checker.check('example.com')
            self.assertEqual(query.call_count, 2)

    def test_cache_is_bounded_and_evicts_least_recently_used(self):
        checker = DomainMXChecker(query=Mock(return_value=answer()), max_cache=2)
        for domain in ('one.com', 'two.com', 'one.com', 'three.com'):
            checker.check(domain)
        self.assertEqual(list(checker.cache), ['one.com', 'three.com'])

    def test_deadline_bounds_fallback_requests_without_retries(self):
        now = [0.0]
        calls = []
        def query(domain, kind, timeout):
            calls.append((domain, kind, timeout))
            now[0] += timeout
            return {'Status': 0}
        checker = DomainMXChecker(query=query, timeout=3, total_timeout=4,
                                  clock=lambda: now[0])
        self.assertEqual(checker.check('example.com')['status'], 'Inconclusive (DNS timeout)')
        self.assertEqual([item[1:] for item in calls], [('MX', 3.0), ('A', 1.0)])

    def test_nonfinite_timeout_rejected_and_large_limits_clamped(self):
        for value in (float('nan'), float('inf'), -float('inf')):
            with self.assertRaises(ValueError):
                DomainMXChecker(timeout=value)
        checker = DomainMXChecker(timeout=999, total_timeout=999, max_cache=999999)
        self.assertEqual((checker.timeout, checker.total_timeout, checker.max_cache), (10, 30, 10000))


class GoogleTransportTests(unittest.TestCase):
    def response(self, payload=None, chunks=None, status=200):
        response = Mock()
        response.status_code = status
        response.iter_content.return_value = chunks if chunks is not None else [json.dumps(payload).encode()]
        response.__enter__ = Mock(return_value=response)
        response.__exit__ = Mock(return_value=False)
        return response

    def test_only_domain_sent_to_fixed_https_resolver_without_redirects(self):
        with patch('email_domain_check.requests.get', return_value=self.response(answer())) as get:
            result = DomainMXChecker().check('example.com')
        self.assertEqual(result['status'], 'MX records found')
        args, kwargs = get.call_args
        self.assertEqual(args, (RESOLVER_URL,))
        self.assertEqual(kwargs['params'], {'name': 'example.com', 'type': 'MX',
                                           'edns_client_subnet': '0.0.0.0/0'})
        self.assertFalse(kwargs['allow_redirects'])
        self.assertTrue(kwargs['stream'])
        self.assertLessEqual(sum(kwargs['timeout']), 3.0)
        self.assertNotIn('cookies', kwargs)
        self.assertNotIn('auth', kwargs)

    def test_http_failure_redirect_oversize_and_invalid_json_fail_closed(self):
        for response in (self.response({}, status=500), self.response({}, status=302),
                          self.response(chunks=[b'x' * 65537]), self.response(chunks=[b'not json'])):
            with patch('email_domain_check.requests.get', return_value=response) as get:
                result = DomainMXChecker().check('example.com')
            self.assertEqual(result['status'], 'Inconclusive (DNS lookup failed)')
            self.assertEqual(get.call_count, 1)
            response.__exit__.assert_called_once()

    def test_response_stream_deadline_is_reported_as_timeout(self):
        with patch('email_domain_check.requests.get', return_value=self.response(answer())), \
             patch('email_domain_check.time.monotonic', side_effect=[0, 4]):
            with self.assertRaises(TimeoutError):
                DomainMXChecker._query_google('example.com', 'MX', 3)


if __name__ == '__main__':
    unittest.main()
