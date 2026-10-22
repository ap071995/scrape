"""Published professional email first; evidence-based, unverified patterns second.

This module never sends mail or probes an SMTP server. Public visibility and
matching a company pattern do not verify ownership or mailbox deliverability.
"""
from __future__ import annotations

import copy
import time
from contextlib import nullcontext
from datetime import datetime, timezone
import html
import re
import unicodedata
from urllib.parse import unquote, urlsplit

from company_public_data import SourceDocument, clean, host, public_url, unknown


EMAIL_RE = re.compile(r"(?<![\w.*%+_\-\u2022\u2026\u25cf\u2219\u25cb])([a-z0-9][a-z0-9._%+'-]*@[a-z0-9](?:[a-z0-9.-]*[a-z0-9])?\.[a-z]{2,24})(?![\w*\-\u2022\u2026\u25cf\u2219\u25cb])", re.I)
FREE_DOMAINS = {'gmail.com', 'googlemail.com', 'outlook.com', 'hotmail.com', 'live.com',
                'yahoo.com', 'yahoo.co.in', 'ymail.com', 'aol.com', 'icloud.com', 'me.com',
                'proton.me', 'protonmail.com', 'mail.com', 'gmx.com', 'msn.com'}
GENERIC = {'info', 'contact', 'hello', 'support', 'sales', 'admin', 'office', 'hr', 'careers',
           'jobs', 'recruitment', 'press', 'media', 'privacy', 'legal', 'enquiries', 'inquiries',
           'help', 'service', 'customerservice', 'billing', 'noreply', 'no-reply', 'webmaster',
           'admissions', 'marketing', 'communications', 'first', 'last', 'firstname', 'lastname'}
PATTERNS = ('{first}.{last}', '{first}_{last}', '{first}-{last}', '{first}{last}',
            '{f}{last}', '{f}.{last}', '{f}_{last}', '{last}.{first}', '{last}_{first}',
            '{last}{first}', '{last}{f}', '{last}_{f}', '{last}.{f}', '{first}{l}',
            '{first}.{l}', '{first}_{l}', '{first}', '{last}', '{f}{l}')
FOUND = 'Found online (unverified)'
GENERATED = 'Generated (unverified)'


def build_email_details(record, checker=None):
    """Classify the selected email without turning unknown origin into evidence.

    DNS is opt-in: pass a DomainMXChecker to check the selected address's actual
    domain. This does not change the address, its publication status, or the
    bounce status. Domain/MX records never establish an individual mailbox.
    """
    address = clean(record.get('Business Email ID'))
    status = clean(record.get('Email Status')).casefold()
    valid = valid_email(address)
    details = {'profile': record.get('Prospect LinkedIn Link', ''),
               'company': record.get('Company Name', ''),
               'published_email': '', 'generated_email': '', 'existing_email': '',
               'email_domain': address.rsplit('@', 1)[-1].lower() if valid else '',
               'domain_status': 'Not checked (disabled)' if valid else 'No valid email',
               'checked_at': '', 'mailbox_status': 'Not verified',
               'detail': 'Domain/MX checks do not verify mailbox ownership or delivery.'}
    if valid:
        category = ('published_email' if status == FOUND.casefold() else
                    'generated_email' if status == GENERATED.casefold() else 'existing_email')
        details[category] = address
        if category == 'existing_email':
            details['detail'] = 'Existing address; publication or generation source is unknown. ' + details['detail']
        if checker is not None:
            try:
                checked = checker.check(details['email_domain'])
                if (not isinstance(checked, dict) or
                        (checked.get('domain') != details['email_domain'] and
                         not (checked.get('domain') == '' and checked.get('status') == 'Invalid domain')) or
                        not all(isinstance(checked.get(key), str) and checked[key]
                                for key in ('status', 'checked_at', 'detail'))):
                    raise ValueError('Invalid domain-check result')
                details.update(domain_status=checked['status'], checked_at=checked['checked_at'],
                               detail=details['detail'] + ' ' + checked['detail'])
            except Exception:
                # An optional resolver failure must never lose the lead or
                # turn an existing address into a negative mailbox verdict.
                details.update(domain_status='Inconclusive (DNS lookup failed)',
                               checked_at=datetime.now(timezone.utc).isoformat(),
                               detail=details['detail'] + ' Optional domain check failed; retry later.')
    return details


def normalized(value):
    value = unicodedata.normalize('NFKD', str(value or '')).casefold()
    return ''.join(ch for ch in value if not unicodedata.combining(ch))


def valid_email(value):
    value = clean(value)
    match = EMAIL_RE.fullmatch(value)
    if not match or '..' in value or value.startswith('.') or len(value) > 254:
        return False
    local, domain = value.rsplit('@', 1)
    return bool(not local.endswith('.') and len(local) <= 64 and all(
        re.fullmatch(r'[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?', label, re.I)
        for label in domain.split('.')))


def business_address(value):
    if not valid_email(value):
        return False
    local, domain = value.casefold().split('@')
    return (domain not in FREE_DOMAINS and local not in GENERIC
            and not re.fullmatch(r'(?:first|last|firstname|lastname)[._-](?:first|last|firstname|lastname)', local)
            and not re.search(r'\.(?:png|jpg|jpeg|gif|svg|webp)$', domain))


def company_domain(record):
    website = clean(record.get('Company Website'))
    if unknown(website):
        return ''
    if '://' not in website:
        website = 'https://' + website
    return host(website) if public_url(website) else ''


def official(url, record):
    if not public_url(url):
        return False
    domain = company_domain(record)
    return bool(domain and (host(url) == domain or host(url).endswith('.' + domain)))


def allowed_domain(email, record, authoritative=False):
    if not business_address(email):
        return False
    domain = email.casefold().split('@')[-1]
    expected = company_domain(record)
    return bool(expected and (domain == expected or domain.endswith('.' + expected) or authoritative))


def email_candidates(first, last, domain):
    if unknown(first) or unknown(last) or unknown(domain):
        return {}
    first = re.sub(r'[^a-z0-9]', '', normalized(first))
    last = re.sub(r'[^a-z0-9]', '', normalized(last))
    domain = clean(domain).lower().removeprefix('@')
    if not first or not last or domain in FREE_DOMAINS or not valid_email('test@' + domain):
        return {}
    values = {'first':first, 'last':last, 'f':first[0], 'l':last[0]}
    return {pattern:pattern.format(**values) + '@' + domain for pattern in PATTERNS}


def decode_text(text):
    text = html.unescape(str(text or ''))
    text = re.sub(r'\s*[\[({]\s*at\s*[\])}]\s*', '@', text, flags=re.I)
    return re.sub(r'\s*[\[({]\s*dot\s*[\])}]\s*', '.', text, flags=re.I)


def name_re(record):
    first, last = (clean(record.get(key, '')) for key in ('First Name', 'Last Name'))
    if unknown(first) or unknown(last):
        return None
    return re.compile(r'(?<!\w)' + re.escape(normalized(first)) + r'(?:\s+[a-z]+){0,2}\s+'
                      + re.escape(normalized(last)) + r'(?!\w)')


def other_name(line, record):
    """Stop at a different staff heading before borrowing its email address."""
    value = clean(line)
    target = name_re(record)
    if target and target.search(normalized(value)):
        return False
    if normalized(value) == normalized(record.get('Company Name', '')):
        return False
    heading = re.split(r'[,|:–—]|\s+-\s+', value, maxsplit=1)[0].strip()
    if heading != value and other_name(heading, record):
        return True
    if re.search(r'director|manager|officer|president|engineer|professor|chair|head|chief|'
                 r'email|phone|contact|department|university|industries|company|executive', value, re.I):
        return False
    return bool(re.fullmatch(r"[A-Z][A-Za-z'-]+(?:\s+[A-Z][A-Za-z'-]+){1,3}", value))


NOT_PUBLISHED = re.compile(r'\b(?:generated|guessed|guess|predicted|possible email|potential email|'
                           r'format|pattern|incorrect|invalid|no longer|do not (?:use|contact|email))\b|'
                           r'\bfor example\b|\bexample\s*:|\bexample (?:only|address|email)\b|'
                           r"\bnot\s+[\w.%+'-]+@|"
                           r'\bnot\s+(?:(?:his|her|their|the)\s+)?(?:email|address)\b', re.I)
DELEGATED_CONTACT = re.compile(
    r"\b(?:assistant|delegate)(?:'s)?(?:[ \t]+(?:email|contact|address))?[ \t]*:|"
    r'\b(?:PA|EA)[ \t]*:|\b(?:contact|email|reach)\b[^\n]{0,60}\b(?:assistant|delegate)\b|'
    r'\bon behalf of\b', re.I)


def owned_context(text, email, record):
    """Require the name and address in one short staff/contact block."""
    lines = [clean(line) for line in decode_text(text).splitlines() if clean(line)]
    target = name_re(record)
    if target is None:
        return ''
    for index, line in enumerate(lines):
        if email.casefold() not in line.casefold():
            continue
        for start in range(max(0, index - 3), index + 1):
            match = target.search(normalized(lines[start]))
            if not match:
                continue
            if start == index and match.end() > normalized(line).find(email.casefold()):
                continue
            block = lines[start:index + 1]
            if len(' '.join(block)) > 550 or any(other_name(item, record) for item in block[1:]):
                continue
            joined = '\n'.join(block)
            if NOT_PUBLISHED.search(joined) or DELEGATED_CONTACT.search(joined):
                continue
            if re.search(r'\b(?:media contact|press contact|for interviews|contact our team)\b', joined, re.I):
                continue
            # Inline directory entries must not borrow a later person's email.
            tail = lines[start][match.end():]
            if any(other_name(pair.group(), record) for pair in
                   re.finditer(r"\b[A-Z][A-Za-z'-]+\s+[A-Z][A-Za-z'-]+\b", tail)):
                continue
            return joined
    return ''


def published_emails(doc, record, source_kind='page'):
    target = name_re(record)
    if target is None or not public_url(doc.url):
        return []
    text = decode_text(doc.text)
    candidates = [(match.group(1), text) for match in EMAIL_RE.finditer(text)]
    for link in doc.links or []:
        if (isinstance(link, dict) and isinstance(link.get('url'), str)
                and link['url'].lower().startswith('mailto:')):
            address = unquote(link['url'][7:].split('?', 1)[0])
            context = str(link.get('context') or link.get('text') or '')
            candidates.append((address, context + '\n' + address))
    facts = []
    for email, context in candidates:
        email = email.lower()
        if not allowed_domain(email, record, authoritative=source_kind == 'page' and official(doc.url, record)):
            continue
        excerpt = owned_context(context, email, record)
        if (not excerpt and source_kind == 'search' and target.search(normalized(doc.title))
                and not target.search(normalized(text))
                and not NOT_PUBLISHED.search(doc.title + '\n' + text)
                and not DELEGATED_CONTACT.search(doc.title + '\n' + text)):
            # A title can supply a missing name, but cannot override a rejected
            # ownership association when the body already names that person.
            expected = email_candidates(record['First Name'], record['Last Name'], email.split('@')[1])
            if email in expected.values() and not any(other_name(line, record) for line in text.splitlines()):
                excerpt = doc.title + '\n' + text
        if not excerpt:
            continue
        facts.append({'email':email, 'url':doc.url, 'method':'Published on source page' if source_kind == 'page'
                      else 'Published in search result; source page not verified',
                      'pattern':'-', 'excerpt':excerpt[:1000],
                      '_rank':3 if source_kind == 'page' and official(doc.url, record) else 2 if source_kind == 'page' else 1})
    return list({(fact['email'], fact['url']):fact for fact in facts}.values())


def company_context(doc, record, company_match):
    candidates = [doc.title] + list(doc.headings)
    explicit = []
    for value in candidates:
        value = re.split(r'\s+[|–—]\s+|\s+-\s+', value)[0]
        base = re.sub(r'\s+(?:Email (?:Format|Pattern|Address|Contact)s?|Contact Information)\b.*', '', value, flags=re.I)
        if base != value:
            explicit.append(base)
        value = base
        if company_match(record.get('Company Name', ''), clean(value)):
            return True
    return bool(official(doc.url, record) and not explicit)


def pattern_evidence(doc, record, company_match):
    if not public_url(doc.url):
        return []
    company = record.get('Company Name', '')
    format_words = r'email\s+(?:formats?|patterns?)'
    qualifiers = r'(?:the\s+)?(?:most\s+)?(?:accurate\s+and\s+)?(?:common|popular|accurate|primary|standard)'

    def subject(value):
        """Read the named company in a heading or explicit format statement."""
        value = clean(value).replace('\u2019', "'")
        value = re.split(r'\s+[|\u2013\u2014]\s+|\s+-\s+', value)[0]
        after = re.search(format_words + r'\s+(?:at|for|of)\s+(.+?)(?:[.:]|$)', value, re.I)
        if after:
            return after[1].strip().removesuffix("'s")
        before = re.match(r'^(.{1,160}?)\s+' + format_words, value, re.I)
        if not before:
            uses = re.match(r'^(.{1,100}?)\s+(?:also\s+)?uses\s+(?=[\[{]|first|last|flast|lfirst)', value, re.I)
            return uses[1].strip().removesuffix("'s") if uses else None
        candidate = re.sub(r'^' + qualifiers + r'\s+', '', before[1], flags=re.I)
        candidate = re.sub(r'\s+' + qualifiers + r'$', '', candidate, flags=re.I)
        candidate = candidate.strip().removesuffix("'s")
        return None if candidate.lower() in {'', 'our', 'the', 'company', 'most common'} else candidate

    headings = [doc.title] + list(doc.headings)
    heading_subjects = [(value, subject(value)) for value in headings]
    matching_headings = [value for value, named in heading_subjects if named and company_match(company, named)]
    if not (matching_headings or company_context(doc, record, company_match)):
        return []
    evidence = []
    text = decode_text(doc.text)
    lines = [clean(line) for line in text.splitlines() if clean(line)]
    active_heading = matching_headings[0] if matching_headings else ''
    foreign_section = False
    table_common = bool(re.search(r'most common', active_heading, re.I))
    table_id = 0
    percentages = {}
    for line_index, raw_line in enumerate(lines):
        named = subject(raw_line)
        has_format = bool(re.search(format_words, raw_line, re.I))
        if named and has_format and '@' not in raw_line:
            foreign_section = not company_match(company, named)
            active_heading = '' if foreign_section else raw_line
            table_common = bool(active_heading and (table_common or re.search(r'most common', raw_line, re.I)))
            table_id += 1
        # Scope preference and company ownership to each sentence. A secondary
        # "also uses" statement must not inherit the primary statement's label.
        clauses = re.split(r'(?<=[.;])\s+(?=[A-Z])', raw_line)
        for clause in clauses:
            named = subject(clause)
            if named and not company_match(company, named):
                # A foreign section remains out of scope even when its rows
                # use explicit brackets or omit the company name entirely.
                if not foreign_section:
                    table_id += 1
                foreign_section, active_heading, table_common = True, '', False
                continue
            if named and foreign_section:
                # An explicit statement about the target can resume its scope.
                foreign_section, active_heading, table_common = False, clause, False
                table_id += 1
            if foreign_section:
                continue
            if re.search(r'\b(?:is|are)\s+not\b|\b(?:do not|no longer|deprecated|retired|incorrect|invalid)\b', clause, re.I):
                continue
            explicit = bool(re.search(format_words, clause, re.I))
            if not (explicit or active_heading or re.search(r'\[(?:first|last)|\{(?:first|last)', clause, re.I)):
                continue
            line = clause.lower()
            for term, replacement in [(r'first(?:[ _-]?name)?[ _-]?initial', 'f'), (r'last(?:[ _-]?name)?[ _-]?initial', 'l'),
                                      (r'first[ _-]?name', 'first'), (r'last[ _-]?name|surname', 'last')]:
                line = re.sub(term, replacement, line)
            line = re.sub(r'[\[\]{}]', '', line)
            token = r'(?:first|last|f|l)'
            previous_end = 0
            for match in re.finditer(r'(?<!\w)(' + token + r'(?:\s*[._-]?\s*' + token + r')?)\s*(?:@\s*([a-z0-9.-]+\.[a-z]{2,24})|(?=\s*\())', line):
                local, domain = match.group(1), match.group(2)
                if not domain:
                    example = EMAIL_RE.search(line[match.end():match.end() + 160])
                    domain = example.group(1).split('@')[-1] if example else ''
                pattern = re.sub(token, lambda part:'{' + part.group() + '}', re.sub(r'\s+', '', local))
                prefix = line[previous_end:match.start()]
                previous_end = match.end()
                if pattern not in PATTERNS or not allowed_domain('employee@' + domain, record, official(doc.url, record)):
                    continue
                preferred = explicit and bool(re.search(r'most common|most accurate(?: and popular)?|most popular|primary|standard', prefix))
                excerpt = (active_heading + '\n' if active_heading and active_heading != raw_line else '') + raw_line
                item = {'pattern':pattern, 'domain':domain, 'url':doc.url, 'excerpt':excerpt[:1000],
                        'preferred':preferred, 'method':'Published company email pattern'}
                evidence.append(item)
                # Percentages are meaningful only under a company-specific
                # "most common" table heading, never from an unrelated statistic.
                if table_common and not explicit:
                    remainder = line[match.end():]
                    percentage = re.search(r'\b(\d{1,3}(?:\.\d+)?)\s*%', remainder)
                    if not percentage and line_index + 1 < len(lines):
                        percentage = re.fullmatch(r'(\d{1,3}(?:\.\d+)?)\s*%', lines[line_index + 1])
                        if percentage:
                            item['excerpt'] = (excerpt + '\n' + lines[line_index + 1])[:1000]
                    if percentage and float(percentage[1]) <= 100:
                        percentages.setdefault(table_id, []).append((float(percentage[1]), item))
    for rows in percentages.values():
        greatest = max(percent for percent, _ in rows)
        leaders = [item for percent, item in rows if percent == greatest]
        if len({(item['pattern'], item['domain']) for item in leaders}) == 1:
            for item in leaders:
                item['preferred'] = True
    return evidence


def inferred_patterns(doc, record):
    """Infer only from multiple named employees on an official company page."""
    if not official(doc.url, record):
        return []
    lines = [clean(line) for line in doc.text.splitlines() if clean(line)]
    by_domain = {}
    for index, line in enumerate(lines):
        person = re.fullmatch(r"([A-Z][a-z'-]+)\s+([A-Z][a-z'-]+)", line)
        if not person or not other_name(line, {'First Name':'Unused', 'Last Name':'Person', 'Company Name':record.get('Company Name')}):
            continue
        employee = {**record, 'First Name':person[1], 'Last Name':person[2]}
        block = '\n'.join(lines[index:index+4])
        for fact in published_emails(SourceDocument(doc.url, doc.title, block), employee):
            email = fact['email']
            domain = email.split('@')[1]
            matches = {pattern for pattern, value in email_candidates(person[1], person[2], domain).items() if value == email}
            if matches:
                by_domain.setdefault(domain, {})[line] = (matches, fact['excerpt'])
    evidence = []
    for domain, people in by_domain.items():
        if len(people) < 2:
            continue
        patterns = set.intersection(*(entry[0] for entry in people.values()))
        if len(patterns) == 1:
            evidence.append({'pattern':patterns.pop(), 'domain':domain, 'url':doc.url,
                             'excerpt':'\n'.join(item[1] for item in people.values())[:1000],
                             'method':'Inferred from multiple named employees on official site'})
    return evidence


def select_pattern(evidence):
    if not evidence:
        return None
    preferred = [item for item in evidence if item.get('preferred')]
    choices = preferred or evidence
    identities = {(item['pattern'], item['domain']) for item in choices}
    return dict(choices[0]) if len(identities) == 1 else None


class BusinessEmailEnricher:
    def __init__(self, web, company_match, log=None, time_budget=15.0, clock=time.monotonic):
        self.web = web
        self.company_match = company_match
        self.log = log or (lambda message: None)
        self.cache = {}
        self.pattern_cache = {}
        self.time_budget = max(0.1, float(time_budget))
        self.clock = clock
        self.deadline = None

    def _expired(self):
        return self.deadline is not None and self.clock() >= self.deadline

    def _fetch(self, url):
        return None if self._expired() else self.web.fetch(url)

    def _search(self, query):
        if self._expired():
            return []
        rows = self.web.search(query)
        if not isinstance(rows, (list, tuple)):
            return []
        valid = []
        for row in rows:
            if not isinstance(row, dict) or not isinstance(row.get('url'), str) or not public_url(row['url']):
                continue
            search_url = row.get('search_url')
            valid.append({**row, 'title':str(row.get('title') or ''), 'text':str(row.get('text') or ''),
                          'search_url':search_url if isinstance(search_url, str) and public_url(search_url) else row['url']})
        return valid

    @staticmethod
    def needs_enrichment(record):
        return (not valid_email(record.get('Business Email ID', ''))
                or clean(record.get('Email Status')).lower().startswith('generated'))

    def _discover(self, record):
        name = clean(record['First Name']) + ' ' + clean(record['Last Name'])
        company = clean(record['Company Name'])
        domain = host(record['Company Website'])
        documents, seen, facts = [], set(), []
        source_index = getattr(self.web, 'source_index', {})
        seed = source_index.get(domain, {}) if isinstance(source_index, dict) else {}
        if not isinstance(seed, dict) or not self.company_match(company, seed.get('company', '')):
            seed = {}
        cached = getattr(self.web, 'documents', {})
        if isinstance(cached, dict):
            documents.extend(doc for doc in cached.values() if isinstance(doc, SourceDocument)
                             and (official(doc.url, record) or company_context(doc, record, self.company_match)))
        # Read every already-known official source before choosing an address.
        # A search snippet cannot outrank a cached profile or staff directory.
        official_urls = seed.get('email_urls', [])
        for url in (official_urls[:3] if isinstance(official_urls, list) else []):
            if not isinstance(url, str) or not official(url, record) or url in seen:
                continue
            seen.add(url)
            fetched = self._fetch(url)
            if fetched:
                documents.append(fetched)
        for doc in documents:
            facts.extend(published_emails(doc, record))
        queries = [f'"{name}" "{company}" email', f'site:{domain} "{name}" email',
                   f'"{name}" "@{domain}"']
        for query in queries:
            for row in self._search(query)[:6]:
                doc = SourceDocument(row.get('url',''), row.get('title',''), row.get('text',''))
                if not name_re(record).search(normalized(doc.title + '\n' + doc.text)):
                    continue
                snippet = published_emails(doc, record, 'search')
                for fact in snippet:
                    fact['url'] = row.get('search_url') or doc.url
                    fact['excerpt'] = f'Result: {doc.url}\n' + fact['excerpt']
                facts.extend(snippet)
                # Reserve a small official-page allowance so directory results
                # cannot consume every fetch before the company-site query.
                can_fetch = len(seen) < 6 or (official(doc.url, record)
                            and sum(official(url, record) for url in seen) < 3)
                if public_url(doc.url) and doc.url not in seen and can_fetch:
                    seen.add(doc.url)
                    fetched = self._fetch(doc.url)
                    if fetched:
                        documents.append(fetched)
                        facts.extend(published_emails(fetched, record))
            # Fetched official evidence can finish the person lookup. Lower
            # ranks must wait for the remaining queries to find that evidence.
            if any(fact['_rank'] == 3 for fact in facts):
                break
        if facts:
            rank = max(fact['_rank'] for fact in facts)
            best = [fact for fact in facts if fact['_rank'] == rank]
            if len({fact['email'] for fact in best}) != 1:
                return None, 'Ambiguous (unverified)'
            return best[0], FOUND
        # Only now, after the person searches, consider generating a candidate.
        if self._expired():
            return None, 'Time limit reached (unverified)'
        key = (company.casefold(), domain)
        if key not in self.pattern_cache:
            evidence = []
            source_urls = seed.get('urls', [])
            for url in (source_urls[:3] if isinstance(source_urls, list) else []):
                if not isinstance(url, str) or not public_url(url) or url in seen or len(seen) >= 10:
                    continue
                seen.add(url)
                fetched = self._fetch(url)
                if fetched:
                    documents.append(fetched)
                    facts.extend(published_emails(fetched, record))
            if facts:
                rank = max(fact['_rank'] for fact in facts)
                best = [fact for fact in facts if fact['_rank'] == rank]
                if len({fact['email'] for fact in best}) != 1:
                    return None, 'Ambiguous (unverified)'
                return best[0], FOUND
            for doc in documents:
                evidence.extend(pattern_evidence(doc, record, self.company_match))
                evidence.extend(inferred_patterns(doc, record))
            for suffix in ('email format', 'email pattern'):
                for row in self._search(f'"{company}" "{domain}" {suffix}')[:5]:
                    doc = SourceDocument(row.get('url',''), row.get('title',''), row.get('text',''))
                    if not company_context(doc, record, self.company_match):
                        continue
                    for fact in published_emails(doc, record, 'search'):
                        fact['url'] = row.get('search_url') or doc.url
                        fact['excerpt'] = f'Result: {doc.url}\n' + fact['excerpt']
                        facts.append(fact)
                    for item in pattern_evidence(doc, record, self.company_match):
                        item['url'] = row.get('search_url') or doc.url
                        item['excerpt'] = f'Result: {doc.url}\n' + item['excerpt']
                        item['method'] += '; search result, source page not verified'
                        evidence.append(item)
                    if public_url(doc.url) and doc.url not in seen and len(seen) < 10:
                        seen.add(doc.url)
                        fetched = self._fetch(doc.url)
                        if fetched:
                            facts.extend(published_emails(fetched, record))
                            evidence.extend(pattern_evidence(fetched, record, self.company_match))
                            evidence.extend(inferred_patterns(fetched, record))
            if not self._expired():
                self.pattern_cache[key] = select_pattern(evidence)
        # A format search may uncover a real staff address. Resolve every such
        # fact, including conflicts, before falling back to a generated address.
        if facts:
            rank = max(fact['_rank'] for fact in facts)
            best = [fact for fact in facts if fact['_rank'] == rank]
            if len({fact['email'] for fact in best}) != 1:
                return None, 'Ambiguous (unverified)'
            return best[0], FOUND
        if self._expired():
            return None, 'Time limit reached (unverified)'
        selected = self.pattern_cache.get(key)
        if selected:
            candidate = email_candidates(record['First Name'], record['Last Name'], selected['domain']).get(selected['pattern'])
            if candidate:
                return {**selected, 'email':candidate}, GENERATED
        return None, 'Not found'

    def enrich(self, record):
        output = copy.deepcopy(record)
        if not self.needs_enrichment(record):
            return output
        required = ('First Name', 'Last Name', 'Company Name', 'Company Website')
        website = clean(record.get('Company Website'))
        if '://' not in website:
            website = 'https://' + website
        if any(unknown(record.get(field)) for field in required) or not public_url(website):
            if not valid_email(record.get('Business Email ID')):
                output['Email Status'] = 'Not found'
            return output
        key = tuple(clean(record.get(field)).casefold() for field in required) + (record.get('Prospect LinkedIn Link',''),)
        # A generated candidate must still be checked for new published evidence
        # when the same record is revisited during this run.
        if key not in self.cache or clean(record.get('Email Status')).lower().startswith('generated'):
            self.deadline = self.clock() + self.time_budget
            budget = getattr(type(self.web), 'time_budget', None)
            try:
                with budget(self.web, self.time_budget) if callable(budget) else nullcontext():
                    discovered = self._discover({**record, 'Company Website':website})
                if not self._expired():
                    self.cache[key] = discovered
            finally:
                self.deadline = None
        else:
            discovered = self.cache[key]
        item, status = copy.deepcopy(discovered)
        if not item:
            if not valid_email(record.get('Business Email ID')):
                output['Business Email ID'] = '-'
                output['Email Status'] = status
            self.log(f'[BUSINESS EMAIL] {record["First Name"]} {record["Last Name"]} | {status}')
            return output
        output['Business Email ID'] = item['email']
        output['Email Status'] = status
        if clean(record.get('Business Email ID')).casefold() != item['email'].casefold():
            output['Bounce Email Status'] = 'Not evaluated'
        note = ('Business email: ' + status + (f'; pattern {item["pattern"]}' if status == GENERATED else ''))
        old = '' if unknown(output.get('Notes')) else output['Notes']
        output['Notes'] = old + (' | ' if old and note not in old else '') + (note if note not in old else '')
        output['_email_evidence'] = [{
            'person':clean(record['First Name']) + ' ' + clean(record['Last Name']),
            'profile':record.get('Prospect LinkedIn Link','-'), 'company':record['Company Name'],
            'email':item['email'], 'method':status + '; ' + item['method'],
            'pattern':item.get('pattern','-'), 'url':item['url'], 'excerpt':item['excerpt'],
            'checked_at':datetime.now(timezone.utc).isoformat(),
        }]
        self.log(f'[BUSINESS EMAIL] {record["First Name"]} {record["Last Name"]} | {status}')
        return output
