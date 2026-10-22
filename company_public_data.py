"""Source-backed company revenue and US industry classification enrichment.

Visible search results can supply explicitly labelled fallback facts. Fetched
company pages take precedence; search-only evidence retains its own provenance.
No LLM inference, industry-to-code guesses, CAPTCHA bypass, or paid data access.
"""
from __future__ import annotations

import base64
from collections import OrderedDict
import copy
from dataclasses import dataclass, field
from datetime import datetime, timezone
import html
import ipaddress
import json
import os
from pathlib import Path
import re
import time
from contextlib import contextmanager
from urllib.parse import parse_qs, quote, unquote, urlsplit

import requests

from company_search_dom import SEARCH_RESULT_SCRIPT


TARGETS = ('Revenue', 'SIC Code', 'NAICS Code')
TIMEOUT_MS = 18_000
MAX_SOURCE_PAGES = 9
DIRECTORY_HOSTS = {'naics.com', 'siccode.com', 'leadiq.com', 'zoominfo.com',
                   'buzzfile.com', 'dnb.com', 'stockanalysis.com', 'prospeo.io'}
BLOCK_RE = re.compile(r'unusual traffic|verify (?:that )?you are human|verify you.*human|'
                      r'checking your browser|access denied|complete the.*challenge|'
                      r'just a moment|captcha', re.I)
NON_ANNUAL_RE = re.compile(r'quarter|\bQ[1-4]\b|three[- ]months?|six[- ]months?|'
                           r'nine[- ]months?|half[- ]year|\bH[12]\b|interim|year[- ]to[- ]date|\bYTD\b', re.I)
ANNUAL_HEADER_RE = re.compile(r'annual results|full.year results|year ended|'
                              r'20\d{2} financial results', re.I)
PROJECTED_RE = re.compile(r'\b(?:budget|forecast|projection|guidance|outlook)s?\b', re.I)


def clean(value):
    return re.sub(r'\s+', ' ', str(value or '')).strip()


def unknown(value):
    return clean(value).lower() in {'', '-', 'unknown', 'none', 'n/a', 'na', 'null',
                                   'nan', 'not available', 'unavailable', 'not found', 'not provided'}


def host(url):
    try:
        return (urlsplit(url if '://' in url else 'https://' + url).hostname or '').lower().rstrip('.').removeprefix('www.')
    except (TypeError, ValueError):
        return ''


def public_url(url):
    """Do not follow local/private URLs returned in an untrusted search page."""
    try:
        parts = urlsplit(url)
        name = (parts.hostname or '').lower().rstrip('.')
        if parts.scheme not in {'http', 'https'} or parts.username or parts.password:
            return False
        if parts.port not in (None, 80, 443) or re.search(r'[\x00-\x20\\]', url):
            return False
        if not name or '.' not in name or name.endswith(('.local', '.localhost', '.internal')):
            return False
        try:
            return ipaddress.ip_address(name).is_global
        except ValueError:
            # Browsers also accept shortened, octal and hexadecimal IPv4
            # spellings. Never pass those ambiguous hosts through as DNS names.
            if re.fullmatch(r'(?:0x[0-9a-f]+|\d+)(?:\.(?:0x[0-9a-f]+|\d+))*', name, re.I):
                return False
            return bool(re.fullmatch(r'[a-z0-9-]+(?:\.[a-z0-9-]+)+', name))
    except (TypeError, ValueError):
        return False


def unwrap_search_url(url):
    try:
        parts = urlsplit(url)
    except (TypeError, ValueError):
        return ''
    query = parse_qs(parts.query)
    source_host = host(url)
    if (source_host == 'bing.com' or source_host.endswith('.bing.com')) and parts.path == '/ck/a':
        value = query.get('u', [''])[0]
        if value.startswith('a1'):
            try:
                return base64.urlsafe_b64decode(value[2:] + '=' * (-len(value[2:]) % 4)).decode()
            except (ValueError, UnicodeError):
                return ''
    if (source_host == 'google.com' or source_host.endswith('.google.com')) and parts.path == '/url':
        return query.get('q', query.get('url', ['']))[0]
    return url


@dataclass
class SourceDocument:
    url: str
    title: str
    text: str
    headings: list[str] = field(default_factory=list)
    links: list[dict] = field(default_factory=list)
    tables: list[list[list[str]]] = field(default_factory=list)


def company_identity(document, record, company_match):
    expected = clean(record.get('Company Name'))
    website_value = record.get('Company Website', '')
    website = '' if unknown(website_value) else host(website_value)
    source_host = host(document.url)
    official = bool(website and (source_host == website or source_host.endswith('.' + website)))
    acronym = ''.join(word[0] for word in expected.split() if word).upper()
    title_parts = re.split(r'\s+[|–—]\s+|\s+-\s+', clean(document.title))
    leading_title = title_parts[0]
    first_heading = clean(document.headings[0]) if document.headings else next(
        (clean(line) for line in document.text.splitlines() if clean(line)), '')
    named = re.search(r'Company Name:\s*([^\n\t|]+)', document.text, re.I)
    if named and not company_match(expected, clean(named.group(1))):
        return False, False
    # A page's parent brand and URL do not make its named subsidiary the parent.
    if (official and first_heading and not company_match(expected, first_heading)
            and re.search(r'\b(?:a business of|subsidiary of|part of)\s+' + re.escape(expected) + r'(?!\w)',
                          document.text[:2500], re.I)):
        return False, False
    generic_heading = re.fullmatch(
        r'(?:(?:FY\s*)?20\d{2}\s+)?(?:investors?|investor relations|financial (?:results|statements|reports?)|'
        r'annual (?:results|reports?)|full.year results|revenues?|overview|about(?: us)?|'
        r'company (?:overview|profile)|financials|fact sheets?|news(?:room)?|press releases?)',
        leading_title, re.I)
    # A parent-domain page may be about a named subsidiary. Its explicit page
    # title and matching first heading override a parent name in the site title.
    if (official and len(title_parts) > 1 and first_heading.casefold() == leading_title.casefold()
            and not generic_heading and not company_match(expected, leading_title)
            and not (3 <= len(acronym) <= 6 and leading_title == acronym)):
        return False, False
    # On an official domain, require the company name as a complete token.
    # Regional subdomains are not used to infer a parent's industry codes.
    if official and re.search(r'(?<!\w)' + re.escape(expected) + r'(?!\w)',
                              document.title + '\n' + document.text[:2500], re.I):
        return True, True
    if official and 3 <= len(acronym) <= 6 and re.search(r'\b' + re.escape(acronym) + r'\b', document.title):
        return True, True
    candidates = list(document.headings[:5]) + [document.title]
    if named:
        candidates.insert(0, named.group(1))
    suffix = r'\s+(?:Company (?:Overview|Profile)|Overview|Revenue|Contact Details|Information|Financials)\b.*'
    for candidate in candidates:
        candidate = re.split(r'\s+[|–—]\s+|\s+-\s+', candidate)[0]
        candidate = re.sub(suffix, '', candidate, flags=re.I)
        candidate = re.sub(r'\s*\([A-Z.]{1,8}\)\s*$', '', candidate)
        if company_match(expected, clean(candidate)):
            # Reject a profile explicitly naming a different website.
            explicit = re.search(r'(?:Website|URL)\s*:?\s*\n?\s*((?:https?://)?(?:www\.)?[a-z0-9.-]+\.[a-z]{2,}(?:/\S*)?)',
                                 document.text, re.I)
            if explicit and website and host(explicit.group(1)) != website:
                return False, False
            return True, False
    return False, False


def _money(text):
    currency = r'(?:US\$|USD|EUR|GBP|JPY|INR|AUD|CAD|CHF|CNY|[$€£¥₹])'
    magnitude = r'(?:trillion|billion|million|thousand|[KMBT])(?![A-Za-z])'
    pattern = rf'{currency}\s*\d[\d,]*(?:\.\d+)?\s*(?:{magnitude})?'
    matches = list(re.finditer(pattern, text, re.I))
    if not matches:
        return ''
    result = matches[0].group().strip()
    if len(matches) > 1:
        gap = text[matches[0].end():matches[1].start()].strip()
        if gap in {'', '-', '–', '—', 'to'}:
            result += ' - ' + matches[1].group().strip()
    # Ranges whose second endpoint omits currency, such as $5 - 10 million.
    tail = text[matches[0].end():]
    continuation = re.match(rf'\s*(?:[-–—]|to)\s*(\d[\d,.]*\s*(?:{magnitude}))', tail, re.I)
    if continuation and len(matches) == 1:
        end = continuation.group(1)
        unit = re.search(magnitude + r'$', end, re.I).group()
        if not re.search(magnitude + r'$', result, re.I):
            result += ' ' + unit
        result += ' - ' + re.match(currency, result, re.I).group() + ' ' + end
    return clean(result)


def _non_annual_section(text):
    """The most recent period heading controls a nearby financial table."""
    partial = list(NON_ANNUAL_RE.finditer(text)) + list(PROJECTED_RE.finditer(text))
    annual = list(ANNUAL_HEADER_RE.finditer(text))
    return bool(partial and (not annual or max(match.start() for match in partial) > annual[-1].start()))


def _other_company_subject(line, company, company_match):
    """Reject an explicit different subject even inside a matched company page."""
    subject = re.match(r"^(.+?)(?:'s|’s)?\s+(?:(?:estimated|annual|total|net)\s+)*"
                       r'(?:revenues?|SIC|NAICS)\b', clean(line), re.I)
    if not subject:
        return False
    name = re.sub(r'\s+(?:reports?|reported|generates?|generated)$', '',
                  clean(subject.group(1)), flags=re.I)
    generic = {'annual', 'estimated', 'total', 'net', 'the', 'the company', 'company',
               'its', 'our', 'consolidated', 'fiscal year', 'primary'}
    dated = re.fullmatch(r'(?:(?:in|for|FY|fiscal year)\s*)?20\d{2}', name, re.I)
    return name.lower() not in generic and not dated and not company_match(company, name)


def _standalone_money(text):
    """A value below a Revenue label must be a value, not another metric."""
    currency = r'(?:US\$|USD|EUR|GBP|JPY|INR|AUD|CAD|CHF|CNY|[$€£¥₹])'
    number = r'\d[\d,]*(?:\.\d+)?\s*(?:trillion|billion|million|thousand|[KMBT])?'
    endpoint = rf'(?:\s*(?:[-–—]|to)\s*(?:{currency})?\s*{number}|\s*{currency}\s*{number})'
    if re.fullmatch(rf'{currency}\s*{number}(?:{endpoint})?',
                    clean(text), re.I):
        return _money(text)
    return ''


def _valid_existing_value(field_name, value):
    """Preserve existing facts only when their structure is usable."""
    if unknown(value):
        return False
    value = clean(value)
    if field_name in {'SIC Code', 'NAICS Code'}:
        digits = 4 if field_name == 'SIC Code' else 6
        return bool(re.fullmatch(rf'\d{{{digits}}}', value))
    if field_name == 'Revenue':
        if _standalone_money(value):
            return True
        # Legacy rows can contain an amount and magnitude without currency.
        number = r'\d[\d,.]*\s*(?:K|M|B|T|thousand|million|billion|trillion)'
        return bool(re.fullmatch(rf'{number}(?:\s*[-–—]\s*{number})?', value, re.I))
    return True


def parse_document(document, record, company_match, *, search_snippet=False):
    """Return facts with exact source excerpts, period and basis; never guesses."""
    if not public_url(document.url):
        return []
    matched, official = company_identity(document, record, company_match)
    if not matched:
        return []
    text = document.text.replace('\xa0', ' ')
    # Secondary-company sections are never evidence for the profile above them.
    text = re.split(r'\n\s*(?:Other companies|Similar companies|Competitors|Target customers)\b',
                    text, maxsplit=1, flags=re.I)[0]
    lines = [clean(line) for line in text.splitlines() if clean(line)]
    company = record['Company Name']
    source_host = host(document.url)
    directory = (search_snippet or source_host in DIRECTORY_HOSTS or
                 any(source_host.endswith('.' + h) for h in DIRECTORY_HOSTS))
    checked_at = datetime.now(timezone.utc).isoformat(timespec='seconds')
    facts = []

    def fact(field_name, value, excerpt, period='Not stated', basis='Company-specific listing', rank=0):
        report_date = re.search(r'Date of Report:\s*(\d{4}-\d{2}-\d{2})', text, re.I)
        if report_date:
            basis += '; source report dated ' + report_date.group(1)
        facts.append({'company': company, 'website': record.get('Company Website', '-'),
                      'field': field_name, 'value': value, 'url': document.url,
                      'period': period, 'basis': basis, 'excerpt': excerpt[:500],
                      'checked_at': checked_at, '_rank': rank})

    # Numeric code *labels*, not sector names, recommendation lists or a code definition.
    # NAICS accepts six-digit US industries only. Eight-digit commercial SIC
    # extensions are reduced to their standard four-digit SIC parent explicitly.
    for label, field_name, digits in [('SIC', 'SIC Code', 4), ('NAICS', 'NAICS Code', 6)]:
        candidates = []
        code_pattern = re.compile(rf'\b{label}(?:\s+(?:Code(?:\(s\))?|1))?\s*(?:is\s*)?[:#\-]?\s*(\d{{{digits}}}(?:\d{{4}})?)\b', re.I)
        for match in code_pattern.finditer(text):
            value = match.group(1)
            if len(value) != digits and not (label == 'SIC' and len(value) == 8):
                continue
            suffix = text[match.end():].split('\n', 1)[0]
            if re.match(r'\s*(?:[.,/–—-]|and\b|or\b|to\b)\s*\d', suffix, re.I):
                continue
            line_start = text.rfind('\n', 0, match.start()) + 1
            if _other_company_subject(text[line_start:match.end()], company, company_match):
                continue
            nearby = text[max(0, match.start()-90):match.end()+100]
            if re.search(r'target|competitor|similar companies|crosswalk|cross-reference', nearby, re.I):
                continue
            candidates.append((value[:digits], nearby, value))
        if candidates and len({c[0] for c in candidates}) == 1:
            value, excerpt, raw = candidates[0]
            basis = 'Published company classification'
            if len(raw) == 8:
                basis += '; four-digit SIC from listed extension ' + raw
            fact(field_name, value, excerpt, basis=basis, rank=4 if official else 2)

    # A quarterly/interim or projected page cannot establish annual revenue.
    # This guard covers tables, snapshots and prose consistently; codes above
    # remain usable on the same company-specific page.
    if NON_ANNUAL_RE.search(document.title) or PROJECTED_RE.search(document.title):
        return facts

    # Map annual table columns explicitly so an older amount cannot acquire the
    # latest year. Only total revenue rows with their own currency are eligible.
    if official:
        for table in document.tables:
            table_text = '\n'.join(' '.join(row) for row in table)
            if NON_ANNUAL_RE.search(table_text) or PROJECTED_RE.search(table_text):
                continue
            # Rendered table labels also occur in body text. Check the period
            # heading before the total row rather than treating year columns
            # alone as proof of an annual reporting period.
            total_row = next((row for row in table if row and
                              re.fullmatch(r'Total (?:Revenues?|Net Sales)', clean(row[0]), re.I)), [])
            flattened_text = clean(text)
            total_at = flattened_text.find(clean(' '.join(total_row))) if total_row else -1
            before_total = flattened_text[:total_at] if total_at >= 0 else flattened_text
            if _non_annual_section(before_total):
                continue
            year_columns = {}
            for row in table:
                years = {i:re.sub(r'[^0-9]', '', value) for i,value in enumerate(row)
                         if re.fullmatch(r'\s*(?:FY\s*)?20\d{2}\s*\*{0,2}\s*', value)}
                if len(years) >= 2:
                    year_columns = years
                    continue
                if not year_columns or not row or not re.fullmatch(r'Total (?:Revenues?|Net Sales)', clean(row[0]), re.I):
                    continue
                column, year = max(year_columns.items(), key=lambda item:int(item[1]))
                if column < len(row) and int(year) <= datetime.now().year:
                    value = _money(row[column])
                    if value:
                        fact('Revenue', value, f'{row[0]} | {year}: {row[column]}', year,
                             'Reported annual total revenue', rank=5)

    # Financial snapshots with amount preceding Sales (billions), common on
    # investor sites. Restrict to the annual Financial Results block.
    annual_block = re.search(r'\b(20\d{2}) Financial Results\b(.{0,650})', text, re.S | re.I)
    if official and annual_block:
        match = re.search(r'([$€£]\s*\d[\d,.]*)\s*\n+\s*(?:Revenue|Sales)\s*\((millions|billions)\)',
                          annual_block.group(2), re.I)
        if match and not _non_annual_section(annual_block.group(2)[:match.end()]):
            unit = match.group(2).removesuffix('s')
            fact('Revenue', clean(match.group(1)) + ' ' + unit, match.group(),
                 annual_block.group(1), 'Reported annual revenue/sales', rank=5)

    for index, line in enumerate(lines):
        if not re.search(r'\b(?:revenues?|sales volume|annual sales|net sales|total sales)\b', line, re.I):
            continue
        if _other_company_subject(line, company, company_match) or NON_ANNUAL_RE.search(line):
            continue
        # Avoid extracting another company from news or competitor lists.
        subject = re.match(r"(.+?)(?:'s|’s)\s+(?:annual\s+)?revenues?\b", line, re.I)
        if subject and not company_match(company, subject.group(1)):
            continue
        context = ' '.join(lines[max(0, index-1):index+3])
        section = ' '.join(lines[max(0, index-6):index+1])
        if _non_annual_section(section):
            continue
        annual_headers = list(re.finditer(r'annual results|full.year results|20\d{2} financial results', section, re.I))
        quarter_headers = list(re.finditer(r'quarter|\bQ[1-4]\b|three months|six months|nine months', section, re.I))
        if quarter_headers and (not annual_headers or quarter_headers[-1].start() > annual_headers[-1].start()):
            continue
        other_subject = re.search(r'(?:revenue|sales)\s+(?:for|of|from)\s+([^:$€£]+):', line, re.I)
        if other_subject and not company_match(company, clean(other_subject.group(1))):
            continue
        if re.search(r'quarter|\bQ[1-4]\b|three months|nine months|six months|'
                     r'budget|forecast|guidance|projected|target|adjusted|segment|expects?|anticipat|outlook|'
                     r'revenue growth|increase in revenue|decrease in revenue|'
                     r'net income|net profit|market cap|per employee|per capita|per share', line, re.I):
            continue
        label_match = re.search(r'\b(?:revenues?|sales volume|annual sales|net sales|total sales)\b', line, re.I)
        amount_text = line[label_match.end():]
        # Tie the amount to the revenue label. A later sentence about taxes,
        # assets or market value is not the company's revenue.
        currency_at = re.search(r'(?:US\$|USD|EUR|GBP|JPY|INR|AUD|CAD|CHF|CNY|[$€£¥₹])\s*\d', amount_text, re.I)
        if currency_at:
            connector = amount_text[:currency_at.start()]
            if len(connector) > 100 or re.search(r'[.!?;]|tax|asset|capital|profit|income', connector, re.I):
                continue
        amount = _money(amount_text)
        # Do not attach the newest year to the first of two annual amounts.
        if len(set(re.findall(r'\b20\d{2}\b', line))) > 1:
            continue
        if len(re.findall(r'(?:[$€£¥₹]|USD|EUR|GBP|CHF)\s*\d', line)) > 1 and ' - ' not in amount:
            continue
        if not amount and re.fullmatch(r'(?:Annual |Estimated |Total |Net )?(?:Revenues?|Sales Volume|Sales):?', line, re.I):
            # Never skip an unavailable value or consume the next metric's
            # currency merely because it is within two lines of Revenue.
            amount = _standalone_money(lines[index+1]) if index+1 < len(lines) else ''
        if not amount:
            continue
        years = re.findall(r'\b(?:FY\s*)?(20\d{2})\b', line)
        if not years:
            years = re.findall(r'\b(?:FY\s*)?(20\d{2})\b', ' '.join(lines[max(0,index-1):index]))
        period = max(years) if years else 'Not stated'
        annual = bool(re.search(r'annual|full.year|year ended|fiscal|\bFY\b', context, re.I))
        # A year + total/net revenue is also explicit annual context unless
        # quarterly wording appears in the immediately preceding header.
        if official and period != 'Not stated' and re.search(r'total revenue|net sales', line, re.I):
            annual = True
        if re.search(r'quarter|\bQ[1-4]\b|three months', ' '.join(lines[max(0,index-1):index]), re.I):
            continue
        if not annual and not directory:
            continue
        estimated = bool(re.search(r'estimat|range', context, re.I)) or not official
        if not annual and directory:
            basis = 'Third-party sales/revenue listing; period not stated'
        else:
            basis = 'Estimated annual revenue' if estimated else 'Reported annual revenue'
        fact('Revenue', amount, context, period, basis, rank=5 if official and annual else 1)
    return facts


def _search_entity_title(title):
    """Strip result descriptors without treating a query as company evidence."""
    title = clean(html.unescape(re.sub(r'<[^>]+>', ' ', title or '')))
    title = re.split(r'\s+[|–—]\s+|\s+-\s+', title)[0]
    title = re.sub(r'\s+(?:(?:FY\s*)?20\d{2}\s+)?(?:(?:annual|estimated|total)\s+)?(?:revenue|sales|results|reports|announces|SIC(?:\s+code)?|'
                   r'NAICS(?:\s+code)?|company (?:profile|overview)|overview|financials|'
                   r'financial (?:results|statements)|contact details)\b.*', '', title, flags=re.I)
    title = re.sub(r"(?:'s|’s)$", '', title)
    return re.sub(r'\s*\([A-Z.]{1,8}\)\s*$', '', title).strip()


def parse_search_result(row, record, company_match):
    """Read a matched company's visible result below fetched-page priority."""
    if row.get('kind', 'organic') not in {'organic', 'featured', 'knowledge'}:
        return []
    provider = clean(row.get('provider')).lower()
    search_url = row.get('search_url', '')
    if provider not in {'google', 'bing', 'brave'} or not public_url(search_url):
        return []
    search_host = host(search_url)
    search_domain = (bool(re.fullmatch(r'google\.(?:com|[a-z]{2,3}|co\.[a-z]{2}|com\.[a-z]{2})', search_host))
                     if provider == 'google' else search_host == {'bing':'bing.com', 'brave':'search.brave.com'}[provider])
    if not search_domain or urlsplit(search_url).path != '/search':
        return []
    title = clean(row.get('title'))
    text = html.unescape(re.sub(r'<[^>]+>', ' ', row.get('text', '') or ''))
    entity = title if company_match(record.get('Company Name', ''), title) else _search_entity_title(title)
    if unknown(record.get('Company Name')) or not company_match(record['Company Name'], entity):
        return []
    if re.search(r'people also ask|related questions|crosswalk|cross.reference|industry definitions|'
                 r'AI overview|AI.generated', title + '\n' + text, re.I):
        return []
    raw_target = row.get('url', '') or ''
    target = unwrap_search_url(raw_target)
    if (raw_target and not target) or (target and not public_url(target)):
        return []
    # Fact cards can put currency after the number. Never infer currency.
    currency = r'USD|EUR|GBP|JPY|INR|AUD|CAD|CHF|CNY'
    # Inline label/value elements can render without a separating text node.
    normalized = re.sub(r'(\b(?:revenues?|(?:SIC|NAICS)(?:\s*Codes?)?))(?=[\d$€£¥₹])',
                        r'\1: ', text, flags=re.I)
    normalized = re.sub(rf'(?<![\w.])(\d[\d,]*(?:\.\d+)?\s*(?:trillion|billion|million|thousand|[KMBT])?)'
                         rf'\s+({currency})\b', r'\2 \1', normalized, flags=re.I)
    lines = []
    for line in re.split(r'\n|;\s*|(?<=[.!?])\s+(?=[A-Z])', normalized):
        line = clean(line)
        if not line or line.endswith('?'):
            continue
        named_code = re.search(r'\b(SIC|NAICS)(?:\s+codes?)?\s+(?:for|of)\s+(.+?)\s+(?:is|are)\s+', line, re.I)
        if named_code:
            if not company_match(record['Company Name'], clean(named_code.group(2))):
                continue
            line = line[:named_code.start()] + named_code.group(1) + ' Code: ' + line[named_code.end():]
        # Even a matched result can mention a competitor. Check an explicit
        # sentence subject before accepting its figures under the result title.
        subject = re.match(r"^(.+?)(?:'s|’s)?\s+(?:(?:estimated|annual|total|net)\s+)*"
                           r'(?:revenues?|SIC|NAICS)\b', line, re.I)
        if subject:
            name = re.sub(r'\s+(?:reports?|reported|generates?|generated)$', '',
                          clean(subject.group(1)), flags=re.I)
            generic = {'annual', 'estimated', 'total', 'net', 'the', 'the company',
                       'company', 'its', 'our', 'consolidated', 'fiscal year'}
            dated = re.fullmatch(r'(?:(?:in|for|FY|fiscal year)\s*)?20\d{2}', name, re.I)
            if name.lower() not in generic and not dated and not company_match(record['Company Name'], name):
                continue
        value_without_year = re.sub(r'\s*\((?:FY\s*)?20\d{2}\)\s*$', '', line)
        if (lines and re.fullmatch(r'(?:Annual |Estimated |Total )?Revenue:?', lines[-1], re.I)
                and _standalone_money(value_without_year)):
            lines[-1] += ': ' + line
        else:
            lines.append(line)
    document = SourceDocument(target or search_url, entity, '\n'.join(lines), [entity])
    facts = parse_document(document, record, company_match, search_snippet=True)
    skip_revenue = bool(NON_ANNUAL_RE.search(title) or PROJECTED_RE.search(title))
    source_is_official = (not unknown(record.get('Company Website')) and bool(target) and
                         (host(target) == host(record['Company Website']) or
                          host(target).endswith('.' + host(record['Company Website']))))
    accepted = []
    for fact in facts:
        if fact['field'] == 'Revenue' and skip_revenue:
            continue
        fact['url'] = search_url
        fact['basis'] = (f"{provider.title()} search result; source page not verified" +
                         ('; estimated revenue' if fact['field']=='Revenue' and
                          re.search(r'estimat|range', text, re.I) else ''))
        fact['excerpt'] = (f"Result: {target or '(no linked source)'}\n{title}\n" + fact['excerpt'])[:500]
        fact['_rank'] = -1 if source_is_official else -2
        accepted.append(fact)
    return accepted


def select_facts(facts):
    selected = []
    for target in TARGETS:
        options = [x for x in facts if x['field'] == target]
        if not options:
            continue
        # Prefer explicit reporting period, primary source, then latest period.
        options.sort(key=lambda x: (x['_rank'], x['period'] if x['period'] != 'Not stated' else ''), reverse=True)
        best = options[0]
        peers = [x for x in options if (x['_rank'], x['period']) == (best['_rank'], best['period'])]
        if len({clean(x['value']).lower() for x in peers}) > 1:
            continue  # conflicting equally ranked sources require review
        selected.append(best)
    return selected


class PublicCompanyEnricher:
    def __init__(self, context, company_match, log=None, trace=None,
                 stop_event=None, pause_check=None, sources_path=None):
        self.context = context
        self.company_match = company_match
        self.log = log or (lambda message: None)
        self.trace = trace or (lambda event, **payload: None)
        self.stop_event = stop_event
        self.pause_check = pause_check
        self.page = None
        self.cache = {}
        self.documents = OrderedDict()
        self._discovery_facts = []
        self.disabled_providers = set()
        self.deadline = None
        path = Path(sources_path) if sources_path else Path(__file__).with_name('company_sources.json')
        try:
            self.source_index = json.loads(path.read_text(encoding='utf-8')) if path.exists() else {}
        except (OSError, ValueError):
            self.source_index = {}

    def close(self):
        if self.page is not None:
            try:
                self.page.close()
            except Exception:
                pass
            self.page = None

    def _active(self):
        return (self.deadline is None or time.monotonic() < self.deadline) and not (
            self.stop_event and self.stop_event.is_set()) and (
            self.pause_check is None or self.pause_check())

    @contextmanager
    def time_budget(self, seconds):
        previous = self.deadline
        self.deadline = min(previous or float('inf'), time.monotonic() + seconds)
        try:
            yield
        finally:
            self.deadline = previous

    def _timeout(self, milliseconds):
        if self.deadline is None:
            return milliseconds
        remaining = int((self.deadline - time.monotonic()) * 1000)
        if remaining <= 0:
            raise TimeoutError('Email lookup time budget reached')
        return max(1, min(milliseconds, remaining))

    def _page(self):
        if self.page is None or self.page.is_closed():
            self.page = self.context.new_page()
        return self.page

    def fetch(self, url):
        if url in self.documents:
            return self.documents[url]
        if not public_url(url) or not self._active() or urlsplit(url).path.lower().endswith('.pdf'):
            return None
        self.documents[url] = None
        # Retain a bounded diagnostic cache for long prospect lists.
        while len(self.documents) > 64:
            self.documents.popitem(last=False)
        try:
            page = self._page()
            response = page.goto(url, wait_until='domcontentloaded',
                                 timeout=self._timeout(5000 if self.deadline else TIMEOUT_MS))
            if not public_url(page.url):
                self.trace('source_unavailable', url=url, reason='unsafe redirect')
                return None
            if response and response.status >= 400:
                self.trace('source_unavailable', url=url, status=response.status)
                return None
            page.wait_for_timeout(self._timeout(250 if self.deadline else 900))
            self._timeout(1)
            data = page.evaluate('''() => {
                const root = document.querySelector('main') || document.body;
                return {title:document.title, text:root.innerText,
                    headings:Array.from(root.querySelectorAll('h1,h2')).slice(0,8).map(e=>e.innerText),
                    links:Array.from(root.querySelectorAll('a[href]')).map(e=>{
                        let context = '';
                        if (e.href.startsWith('mailto:')) {
                            let parent = e.parentElement;
                            for (let depth=0; parent && depth<5; depth++,parent=parent.parentElement) {
                                const text = (parent.innerText || '').trim();
                                if (text.length>800 || parent.querySelectorAll('a[href^="mailto:"]').length>1) break;
                                context = text;
                                if (parent.matches('article, li, tr')) break;
                            }
                        }
                        return {url:e.href,text:e.innerText,context};
                    }).slice(0,350),
                    tables:Array.from(root.querySelectorAll('table')).slice(0,20).map(t=>
                      Array.from(t.querySelectorAll('tr')).slice(0,120).map(r=>
                        Array.from(r.querySelectorAll('th,td')).map(c=>c.innerText.trim())))};
            }''')
            if BLOCK_RE.search(data['title'] + '\n' + data['text'][:1200]):
                self.trace('source_blocked', url=url)
                return None
            doc = SourceDocument(page.url, data['title'], data['text'][:180000], data['headings'], data['links'], data['tables'])
            self.documents[url] = doc
            return doc
        except Exception as exc:
            if self.deadline is not None:
                self.documents.pop(url, None)  # A short-budget failure must be retryable.
            self.trace('source_error', url=url, error=type(exc).__name__, detail=str(exc)[:800])
            # Failed navigations can leave a pending chrome-error page which
            # interrupts the next unrelated source. Recreate only our own tab.
            self.close()
            return None

    def search(self, query):
        # Optional API avoids fragile search-result DOM. Keys never enter logs.
        api_key = os.environ.get('BRAVE_SEARCH_API_KEY', '')
        if api_key and 'brave' not in self.disabled_providers and self._active():
            try:
                if self.deadline is not None:
                    # Playwright's request timeout bounds the whole API response.
                    response = self.context.request.get('https://api.search.brave.com/res/v1/web/search',
                        params={'q': query, 'count': 6}, headers={'X-Subscription-Token': api_key},
                        timeout=self._timeout(5000))
                    try:
                        if not response.ok:
                            raise ValueError('Search API unavailable')
                        payload = response.json()
                    finally:
                        response.dispose()
                else:
                    response = requests.get('https://api.search.brave.com/res/v1/web/search',
                        params={'q':query, 'count':6}, headers={'X-Subscription-Token':api_key}, timeout=15)
                    response.raise_for_status()
                    payload = response.json()
                rows = [{'url':x['url'], 'title':x.get('title',''),
                         'text':html.unescape(re.sub(r'<[^>]+>', ' ', x.get('description',''))),
                         'provider':'brave', 'kind':'organic',
                         'search_url':'https://search.brave.com/search?q=' + quote(query, safe='')}
                        for x in payload.get('web',{}).get('results',[])]
                if rows:
                    self.trace('search_results', provider='brave', query=query, results=rows)
                    return rows
            except Exception:
                if self.deadline is None:
                    self.disabled_providers.add('brave')
                self.log('[COMPANY WEB] Search API unavailable; using browser search.')
        for provider, base in [('google','https://www.google.com/search?q='),
                               ('bing','https://www.bing.com/search?q=')]:
            if provider in self.disabled_providers or not self._active():
                continue
            try:
                page = self._page()
                page.goto(base + quote(query, safe=''), wait_until='domcontentloaded',
                          timeout=self._timeout(5000 if self.deadline else TIMEOUT_MS))
                page.wait_for_timeout(self._timeout(400 if self.deadline else 1200))
                body = page.locator('body').inner_text(timeout=self._timeout(3000))
                if '/sorry/' in page.url or BLOCK_RE.search(body[:2500]):
                    self.disabled_providers.add(provider)
                    self.log(f'[COMPANY WEB] {provider} requires verification; provider stopped for this run.')
                    continue
                self._timeout(1)
                rows = page.evaluate(SEARCH_RESULT_SCRIPT, provider)
                for row in rows:
                    row['url'] = unwrap_search_url(row['url'])
                    row.update(provider=provider, search_url=page.url)
                rows = [x for x in rows if clean(x['title']) and
                        (public_url(x['url']) or (x.get('kind')=='knowledge' and not x['url']))]
                self.trace('search_results', provider=provider, query=query, results=rows)
                if rows:
                    return rows
            except Exception as exc:
                self.trace('search_error', provider=provider, error=type(exc).__name__, detail=str(exc)[:800])
                self.close()
        return []

    def _discover(self, record):
        company = record['Company Name']
        website = record.get('Company Website', '')
        urls = []
        self._discovery_facts = []
        index = self.source_index.get(host(website), {})
        if index and self.company_match(company, index.get('company', '')):
            urls.extend(index.get('urls', []))
        for suffix in ['annual revenue', 'SIC code', 'NAICS code']:
            if not self._active():
                break
            for result in self.search(f'"{company}" {suffix}'):
                snippet_facts = parse_search_result(result, record, self.company_match)
                self._discovery_facts.extend(snippet_facts)
                if snippet_facts:
                    self.trace('search_facts', company=company, facts=snippet_facts)
                title = result.get('title', '')
                relevant = (suffix in {'SIC code', 'NAICS code'} or
                            re.search(r'revenue|sales|SIC|NAICS|annual|investor|financial', title, re.I))
                directory = host(result['url']) in DIRECTORY_HOSTS
                if public_url(result['url']) and (relevant or directory) and (
                        self.company_match(company, _search_entity_title(title)) or
                        re.search(re.escape(company), title, re.I) or host(result['url']) == host(website)):
                    urls.append(result['url'])
        if public_url(website):
            urls.append(website)
        return list(dict.fromkeys(url.split('#', 1)[0] for url in urls))[:MAX_SOURCE_PAGES]

    def enrich(self, record):
        output = dict(record)
        if unknown(record.get('Company Name')) or not self._active():
            return output
        key = (clean(record['Company Name']).casefold(), host(record.get('Company Website', '')))
        if key not in self.cache:
            facts = []
            pending = self._discover(record)
            facts.extend(self._discovery_facts)
            seen = set()
            while pending and len(seen) < MAX_SOURCE_PAGES:
                url = pending.pop(0).split('#', 1)[0]
                if url in seen:
                    continue
                seen.add(url)
                if not self._active():
                    break
                doc = self.fetch(url)
                if doc:
                    found = parse_document(doc, record, self.company_match)
                    facts.extend(found)
                    self.trace('source_facts', company=record['Company Name'], url=url, facts=found)
                    selected = select_facts(facts)
                    if len(selected) == 3 and any(x['field']=='Revenue' and x['_rank']==5 for x in selected):
                        break
                    # Investor homepages often link to a readable annual-results
                    # page. Follow a bounded number of same-company HTML links.
                    _, official = company_identity(doc, record, self.company_match)
                    if official:
                        candidates = [link['url'].split('#', 1)[0] for link in doc.links if
                            re.search(r'annual (?:report|result)|financial results|revenue', link.get('text',''), re.I)
                            and (host(link['url']) == host(record.get('Company Website','')) or
                                 host(link['url']).endswith('.' + host(record.get('Company Website',''))))
                            and not urlsplit(link['url']).path.lower().endswith('.pdf')
                            and link['url'].split('#', 1)[0] not in seen and link['url'].split('#', 1)[0] not in pending]
                        pending.extend(list(dict.fromkeys(candidates))[:2])
            self.cache[key] = select_facts(facts)
        evidence = []
        for item in copy.deepcopy(self.cache[key]):
            field_name = item['field']
            if _valid_existing_value(field_name, output.get(field_name)) and output[field_name] != item['value']:
                continue
            output[field_name] = item['value']
            item.pop('_rank', None)
            evidence.append(item)
            if field_name == 'Revenue':
                output['Revenue Link'] = item['url']
                note = f"Revenue period: {item['period']}; {item['basis']}"
                old = '' if unknown(output.get('Notes')) else output['Notes']
                output['Notes'] = old + (' | ' if old and note not in old else '') + (note if note not in old else '')
        code_evidence = [x for x in evidence if x['field'] in {'SIC Code', 'NAICS Code'}]
        if code_evidence:
            output['SIC/NAICS Code Link'] = code_evidence[-1]['url']
        output['_company_evidence'] = evidence
        missing = [name for name in TARGETS if not _valid_existing_value(name, output.get(name))]
        search_count = sum('search result;' in item['basis'] for item in evidence)
        self.log(f"[COMPANY WEB] {record['Company Name']} | sourced={len(evidence)} | "
                 f"search_results={search_count} | unresolved={', '.join(missing) or 'none'}")
        return output
