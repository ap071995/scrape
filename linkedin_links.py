"""Strict mixed-link parsing and durable, explicit Sales Navigator URL mappings."""
from contextlib import closing
from dataclasses import dataclass
from datetime import datetime, timezone
import re
import sqlite3
from pathlib import Path
from urllib.parse import unquote, urlsplit, urlunsplit


@dataclass(frozen=True)
class LinkInput:
    original: str
    canonical: str = ''
    kind: str = 'Invalid'
    reason: str = ''


def classify_link(value):
    original = str(value or '').strip().rstrip('.,;)]}\'"')
    url = original
    if re.match(r'^(?:[a-z0-9-]+\.)*linkedin\.com/', url, re.I):
        url = 'https://' + url
    try:
        parts = urlsplit(url)
        host = (parts.hostname or '').lower()
        if (parts.scheme.lower() not in ('http', 'https') or
                not (host == 'linkedin.com' or host.endswith('.linkedin.com')) or
                parts.username or parts.password or parts.port not in (None, 80, 443)):
            raise ValueError('Expected a LinkedIn profile or Sales Navigator lead URL.')
        path = parts.path.rstrip('/')
        if re.fullmatch(r'/in/[^/\s,]+', path):
            slug = unquote(path[4:])
            if re.search(r'[/\\\s?#]', slug) or slug in ('.', '..'):
                raise ValueError('Invalid public profile identifier.')
            kind = 'LinkedIn profile'
        elif re.fullmatch(r'/sales/lead/[A-Za-z0-9_-]+(?:,[^/\s]*)?', path):
            path = path.split(',', 1)[0]
            kind = 'Sales Navigator lead'
        else:
            raise ValueError('Use /in/ profile links or /sales/lead/ links; search/company URLs are not profiles.')
        return LinkInput(original, urlunsplit(('https', 'www.linkedin.com', path, '', '')), kind)
    except (ValueError, TypeError) as exc:
        return LinkInput(original, reason=str(exc))


def extract_link_inputs(value):
    """Read URLs inside plain cells, Markdown and HYPERLINK formula strings."""
    raw = str(value or '').strip()
    candidates = re.findall(r'https?://[^\s<>"\[\]]+', raw, re.I)
    if not candidates:
        candidates = re.findall(r'(?<![\w@.])(?:[a-z0-9-]+\.)*linkedin\.com/[^\s<>"\[\]]+', raw, re.I)
    # Markdown wrappers and formulas can contain the same address twice.
    return list(dict.fromkeys(classify_link(item) for item in candidates))


class LinkResolutionStore:
    """Only verified UI-derived mappings persist. Failures are always retryable.

    Entries have no expiry: Refresh saved LinkedIn profiles explicitly bypasses
    and invalidates the old mapping before another attempt.
    """
    def __init__(self, path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with closing(sqlite3.connect(self.path, timeout=15)) as db, db:
            db.execute('CREATE TABLE IF NOT EXISTS profile_links '
                       '(lead TEXT PRIMARY KEY, profile TEXT NOT NULL, checked_at TEXT NOT NULL)')

    def get(self, lead):
        key = classify_link(lead)
        if key.kind != 'Sales Navigator lead':
            return None
        with closing(sqlite3.connect(self.path, timeout=15)) as db:
            row = db.execute('SELECT profile FROM profile_links WHERE lead=?', (key.canonical,)).fetchone()
        if row and classify_link(row[0]).kind == 'LinkedIn profile':
            return classify_link(row[0]).canonical
        return None

    def put(self, lead, profile):
        source, target = classify_link(lead), classify_link(profile)
        if source.kind != 'Sales Navigator lead' or target.kind != 'LinkedIn profile':
            raise ValueError('A mapping must connect a Sales Navigator lead to a public LinkedIn profile.')
        with closing(sqlite3.connect(self.path, timeout=15)) as db, db:
            db.execute('INSERT OR REPLACE INTO profile_links VALUES (?,?,?)',
                       (source.canonical, target.canonical, datetime.now(timezone.utc).isoformat()))

    def invalidate(self, lead):
        with closing(sqlite3.connect(self.path, timeout=15)) as db, db:
            db.execute('DELETE FROM profile_links WHERE lead=?', (classify_link(lead).canonical,))


RESOLUTION_DOM = r"""() => {
    const visible = el => el && el.getBoundingClientRect().width > 0 &&
        el.getBoundingClientRect().height > 0 && getComputedStyle(el).visibility !== 'hidden';
    const text = (document.querySelector('main') || document.body)?.innerText || '';
    const auth = /\/(login|checkpoint|authwall|uas\/login)(\/|\?|$)/i.test(location.href) ||
        [...document.querySelectorAll('input[type=password]')].some(visible) ||
        /sign in to (view|see)|verify your identity|security verification|unusual activity|temporarily restricted|you.ve reached (your|the) .*limit|subscribe to (access|use) sales navigator/i.test(text);
    const unavailable = /profile not found|page not found|this (page|profile) (doesn.t exist|is(n.t| not) available)/i.test(text);
    const links = [...document.querySelectorAll('a[href]')].filter(a => visible(a) &&
        !a.closest('aside, footer') && (!a.closest('nav') || a.closest('[role="menu"]')) &&
        /^(view|open|copy)(?:\s+[^\n]*?)?\s+linkedin\s+(profile|profile url)(?:\s.*)?$/i.test(
            (a.getAttribute('aria-label') || a.innerText || '').trim()))
        .map(a => a.href);
    return {auth, unavailable, links};
}"""
