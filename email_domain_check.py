"""Optional DNS-only email-domain checks; never SMTP or mailbox verification.

The default resolver is Google Public DNS's JSON HTTPS endpoint. Only the domain
is submitted, never the email local part or a person's name. No checks occur
until ``check`` is called. A caller can inject ``query(domain, type, timeout)``
to use a different resolver or test without network access.

Semantics: RFC 5321 section 5 (MX and implicit A/AAAA fallback), RFC 7505
(a sole ``0 .`` MX means no mail). A listed MX is not proof its server works.
https://www.rfc-editor.org/rfc/rfc5321#section-5
https://www.rfc-editor.org/rfc/rfc7505#section-3
https://developers.google.com/speed/public-dns/docs/doh/json
"""
from __future__ import annotations

import copy
from collections import OrderedDict
from datetime import datetime, timezone
import ipaddress
import json
import math
import re
import time

import requests


RESOLVER_URL = 'https://dns.google/resolve'
NOT_VERIFIED = 'Individual mailbox not verified; DNS does not prove delivery or ownership.'


def normalized_domain(value):
    """Accept public DNS names only; never interpret an email or URL as a host."""
    if not isinstance(value, str) or value != value.strip() or value.endswith('..'):
        return ''
    try:
        domain = value.rstrip('.').encode('idna').decode('ascii').lower()
    except (UnicodeError, ValueError):
        return ''
    if (len(domain) > 253 or '.' not in domain or
            not all(re.fullmatch(r'[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?', label)
                    for label in domain.split('.'))):
        return ''
    try:
        ipaddress.ip_address(domain)
        return ''
    except ValueError:
        pass
    if domain.split('.')[-1] in {'localhost', 'local', 'internal', 'lan', 'home', 'invalid', 'test', 'onion'}:
        return ''
    if domain.endswith('.home.arpa') or domain.split('.')[-1].isdigit():
        return ''
    return domain


class DomainMXChecker:
    """Small, cached, best-effort lookup with a deadline and no retries.

    At most three queries (MX, then A/AAAA only if MX is absent) per uncached
    domain. HTTP connect/read timeouts and a total lookup budget bound normal
    network failures. Cache lifetime respects positive DNS TTLs up to one hour;
    inconclusive/error results expire after one minute, other negatives after
    five minutes. Returned dictionaries are copies so callers cannot poison it.
    """
    def __init__(self, query=None, timeout=3.0, total_timeout=8.0, clock=None,
                 max_cache=1024):
        self.query = query or self._query_google
        if not all(math.isfinite(float(value)) for value in (timeout, total_timeout)):
            raise ValueError('DNS timeouts must be finite')
        self.timeout = min(10.0, max(0.1, float(timeout)))
        self.total_timeout = min(30.0, max(0.1, float(total_timeout)))
        self.clock = clock or time.monotonic
        self.max_cache = min(10000, max(1, int(max_cache)))
        self.cache = OrderedDict()

    @staticmethod
    def _query_google(domain, record_type, timeout):
        # No redirects, credentials, account cookies, or automatic retries.
        started = time.monotonic()
        with requests.get(RESOLVER_URL,
                          params={'name': domain, 'type': record_type,
                                  'edns_client_subnet': '0.0.0.0/0'},
                          headers={'Accept': 'application/dns-json'},
                          timeout=(min(1.5, timeout / 2), max(0.05, timeout / 2)),
                          allow_redirects=False, stream=True) as response:
            if response.status_code != 200:
                raise ValueError('DNS resolver HTTP response was not successful')
            body = bytearray()
            for chunk in response.iter_content(chunk_size=2048):
                body.extend(chunk)
                if time.monotonic() - started >= timeout:
                    raise TimeoutError('DNS resolver deadline reached')
                if len(body) > 65536:
                    raise ValueError('DNS resolver response exceeded its limit')
            return json.loads(body)

    @staticmethod
    def _answers(payload, record_type):
        if (not isinstance(payload, dict) or type(payload.get('Status')) is not int
                or payload.get('TC') is True):
            raise ValueError('Malformed or truncated DNS response')
        status = payload['Status']
        if status not in (0, 3):
            raise ValueError('DNS lookup temporarily unavailable')
        answers = payload.get('Answer', [])
        if not isinstance(answers, list) or not all(isinstance(item, dict) for item in answers):
            raise ValueError('Malformed DNS answers')
        records = [item for item in answers if type(item.get('type')) is int
                   and item.get('type') == record_type]
        if status == 3 and records:
            raise ValueError('DNS response contradicts NXDOMAIN')
        return status, records

    def check(self, domain):
        domain = normalized_domain(domain)
        checked_at = datetime.now(timezone.utc).isoformat()
        if not domain:
            return {'domain': '', 'status': 'Invalid domain', 'checked_at': checked_at,
                    'detail': 'A public DNS domain is required. ' + NOT_VERIFIED}
        cached = self.cache.get(domain)
        if cached and cached[0] > self.clock():
            self.cache.move_to_end(domain)
            return copy.deepcopy(cached[1])
        deadline = self.clock() + self.total_timeout
        ttl = 300

        def query(kind):
            remaining = deadline - self.clock()
            if remaining <= 0:
                raise TimeoutError('DNS lookup deadline reached')
            payload = self.query(domain, kind, min(self.timeout, remaining))
            if self.clock() >= deadline:
                raise TimeoutError('DNS lookup deadline reached')
            return self._answers(payload, {'MX': 15, 'A': 1, 'AAAA': 28}[kind])

        def positive_ttl(records):
            values = [entry.get('TTL') for entry in records]
            values = [value for value in values if type(value) is int and value >= 0]
            return min([3600] + values) if values else 300

        try:
            status, records = query('MX')
            if status == 3:
                result_status, detail = 'Domain does not exist (NXDOMAIN)', 'Resolver reports a non-existent domain.'
            elif records:
                mx = []
                for entry in records:
                    match = re.fullmatch(r'(\d{1,5})\s+(\S+)', str(entry.get('data', '')))
                    if not match or int(match[1]) > 65535:
                        raise ValueError('Malformed MX record')
                    preference, exchange = int(match[1]), match[2]
                    if exchange != '.' and not normalized_domain(exchange):
                        raise ValueError('Malformed MX exchange')
                    mx.append((preference, exchange))
                if mx == [(0, '.')]:
                    result_status, detail = 'Does not accept mail (null MX)', 'Domain explicitly publishes null MX.'
                    ttl = min(300, positive_ttl(records))
                elif any(exchange == '.' for _, exchange in mx):
                    result_status, detail = 'Inconclusive (invalid MX configuration)', 'Null MX conflicts with other MX data.'
                    ttl = 60
                else:
                    result_status, detail = 'MX records found', 'DNS lists mail exchangers; their availability was not tested.'
                    ttl = positive_ttl(records)
            else:
                found, temporary_error, timed_out, negative_answers = [], False, False, 0
                for kind, version in (('A', 4), ('AAAA', 6)):
                    try:
                        address_status, addresses = query(kind)
                        if address_status == 3:
                            negative_answers += 1
                            continue
                        for entry in addresses:
                            address = ipaddress.ip_address(entry.get('data', ''))
                            if address.version != version:
                                raise ValueError('DNS address has wrong type')
                        if addresses:
                            found.extend(addresses)
                            break
                    except (TimeoutError, requests.Timeout):
                        timed_out = True
                    except (ValueError, TypeError, requests.RequestException):
                        temporary_error = True
                if found:
                    result_status = 'No MX; A/AAAA fallback available'
                    detail = 'Implicit mail routing is possible under RFC 5321; no mail server was contacted.'
                    ttl = positive_ttl(found)
                elif timed_out:
                    result_status, detail = 'Inconclusive (DNS timeout)', 'DNS lookup timed out; retry later.'
                    ttl = 60
                elif temporary_error or negative_answers:
                    result_status, detail = 'Inconclusive (DNS lookup failed)', 'DNS responses were incomplete or inconsistent; retry later.'
                    ttl = 60
                else:
                    result_status, detail = 'No MX or A/AAAA records', 'No DNS mail route was found.'
        except (TimeoutError, requests.Timeout):
            result_status, detail = 'Inconclusive (DNS timeout)', 'DNS lookup timed out; retry later.'
            ttl = 60
        except (ValueError, TypeError, requests.RequestException):
            result_status, detail = 'Inconclusive (DNS lookup failed)', 'DNS lookup failed or returned incomplete data; retry later.'
            ttl = 60
        result = {'domain': domain, 'status': result_status, 'checked_at': checked_at,
                  'detail': detail + ' ' + NOT_VERIFIED}
        self.cache[domain] = (self.clock() + ttl, copy.deepcopy(result))
        self.cache.move_to_end(domain)
        while len(self.cache) > self.max_cache:
            self.cache.popitem(last=False)
        return result
