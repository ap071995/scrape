"""Durable caches, evidence snapshots, stage timing and isolated enrichment sessions."""
from __future__ import annotations

import copy
import json
import os
import sqlite3
import threading
import time
from functools import wraps
from collections.abc import MutableMapping
from concurrent.futures import ThreadPoolExecutor
from contextlib import contextmanager
from pathlib import Path


def timed(stage):
    def decorate(method):
        @wraps(method)
        def wrapper(self, *args, **kwargs):
            with self.metrics.measure(stage):
                return method(self, *args, **kwargs)
        return wrapper
    return decorate


class RunMetrics:
    def __init__(self, on_stage=None):
        self.started = time.monotonic()
        self.lock = threading.Lock()
        self.counts = {}
        self.seconds = {}
        self.on_stage = on_stage

    def notify(self, name, state, detail=''):
        if self.on_stage:
            self.on_stage(name, state, detail)

    def count(self, name, amount=1):
        with self.lock:
            self.counts[name] = self.counts.get(name, 0) + amount

    @contextmanager
    def measure(self, name):
        start = time.monotonic()
        self.notify(name, 'running')
        state = 'done'
        try:
            yield
        except BaseException:
            state = 'error'
            raise
        finally:
            with self.lock:
                self.seconds[name] = self.seconds.get(name, 0) + time.monotonic() - start
                self.counts[name + '_calls'] = self.counts.get(name + '_calls', 0) + 1
            self.notify(name, state, f'{time.monotonic() - start:.1f}s')

    def snapshot(self):
        with self.lock:
            return {'elapsed_seconds': round(time.monotonic() - self.started, 3),
                    'counts': dict(self.counts),
                    'stage_seconds': {k: round(v, 3) for k, v in self.seconds.items()}}


class EvidenceCache(MutableMapping):
    """SQLite entries expire; values are copied so one prospect cannot alter another.

    Connections are short lived and owned by the calling thread. Only named
    namespaces are cleared; an enrichment worker never clears another's cache.
    """
    def __init__(self, path, namespace, ttl=7 * 86400, metrics=None, clock=time.time):
        self.path = Path(path)
        self.namespace = namespace
        self.ttl = ttl
        self.metrics = metrics
        self.clock = clock
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with self._connect() as db:
            db.execute('PRAGMA journal_mode=WAL')
            db.execute('CREATE TABLE IF NOT EXISTS evidence '
                       '(namespace TEXT, key TEXT, value TEXT, expires REAL, '
                       'PRIMARY KEY(namespace, key))')
            db.execute('DELETE FROM evidence WHERE expires <= ?', (clock(),))

    @contextmanager
    def _connect(self):
        connection = sqlite3.connect(self.path, timeout=15)
        try:
            with connection:
                yield connection
        finally:
            connection.close()

    @staticmethod
    def _key(key):
        return json.dumps(key, ensure_ascii=False, sort_keys=True)

    def __getitem__(self, key):
        with self._connect() as db:
            row = db.execute('SELECT value FROM evidence WHERE namespace=? AND key=? AND expires>?',
                             (self.namespace, self._key(key), self.clock())).fetchone()
        if row is None:
            raise KeyError(key)
        if self.metrics:
            self.metrics.count(self.namespace + '_cache_hits')
        return json.loads(row[0])

    def __contains__(self, key):
        with self._connect() as db:
            return db.execute('SELECT 1 FROM evidence WHERE namespace=? AND key=? AND expires>?',
                              (self.namespace, self._key(key), self.clock())).fetchone() is not None

    def __setitem__(self, key, value):
        # Unresolved lookups expire quickly instead of suppressing retries for days.
        ttl = self.ttl
        if not value or isinstance(value, dict) and value.get('company_matched') is False:
            ttl = min(ttl, 900)
        if isinstance(value, dict):
            required = (('address', 'phone') if 'company_matched' in value else
                        ('Company Website', 'Industry', 'Company Size') if 'Company Name' in value else ())
            if any(str(value.get(key) or '').casefold() in {'', '-', 'unknown', 'none'} for key in required):
                ttl = min(ttl, 900)
        if self.namespace == 'public_company' and isinstance(value, list) and len(value) < 3:
            ttl = min(ttl, 900)
        with self._connect() as db:
            db.execute('INSERT OR REPLACE INTO evidence VALUES (?,?,?,?)',
                       (self.namespace, self._key(key), json.dumps(value, ensure_ascii=False),
                        self.clock() + ttl))

    def __delitem__(self, key):
        with self._connect() as db:
            if not db.execute('DELETE FROM evidence WHERE namespace=? AND key=?',
                              (self.namespace, self._key(key))).rowcount:
                raise KeyError(key)

    def __iter__(self):
        with self._connect() as db:
            rows = db.execute('SELECT key FROM evidence WHERE namespace=? AND expires>?',
                              (self.namespace, self.clock())).fetchall()
        for (key,) in rows:
            value = json.loads(key)
            yield tuple(value) if isinstance(value, list) else value

    def __len__(self):
        with self._connect() as db:
            return db.execute('SELECT COUNT(*) FROM evidence WHERE namespace=? AND expires>?',
                              (self.namespace, self.clock())).fetchone()[0]

    def clear(self):
        with self._connect() as db:
            db.execute('DELETE FROM evidence WHERE namespace=?', (self.namespace,))


class RecoveryJournal:
    """Durable row log: Excel can be batched without losing accepted rows on a crash."""
    def __init__(self, workbook_path):
        self.path = Path(str(workbook_path) + '.pending.jsonl')

    def append(self, row, record):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        payload = json.dumps({'row': row, 'record': record}, ensure_ascii=False, default=str)
        with self.path.open('a', encoding='utf-8') as stream:
            stream.write(payload + '\n')
            stream.flush()
            os.fsync(stream.fileno())

    def read(self):
        if not self.path.exists():
            return []
        lines = self.path.read_bytes().splitlines(keepends=True)
        records = []
        for index, line in enumerate(lines):
            try:
                records.append(json.loads(line))
            except (ValueError, UnicodeDecodeError):
                if index != len(lines) - 1 or line.endswith(b'\n'):
                    raise ValueError('Damaged recovery journal; preserve it for review: ' + str(self.path))
                # Only a torn final write is ignored. Earlier durable rows replay.
        return records

    def clear(self):
        self.path.unlink(missing_ok=True)


class EnrichmentSessions:
    """Two long-lived workers, each with its own Playwright runtime and browser.

    The caller waits for a record's enrichment before the next LinkedIn profile.
    Only dictionaries cross thread boundaries. No LinkedIn cookies are copied.
    """
    def __init__(self, factory, concurrent=True, on_error=None):
        self.factory = factory
        self.concurrent = concurrent
        self.on_error = on_error
        self.executors = {name: ThreadPoolExecutor(max_workers=1, thread_name_prefix=name)
                          for name in ('maps', 'company_email')}
        self.workers = {}

    def _run(self, stage, record):
        if stage not in self.workers:
            self.workers[stage] = self.factory(stage)
        return self.workers[stage].process(copy.deepcopy(record))

    def _result(self, future, stage, original):
        try:
            return future.result()
        except Exception as exc:
            if self.on_error is None:
                raise
            self.on_error(f'{stage} session failed: {type(exc).__name__}: {exc}')
            return {**original, 'Notes': (str(original.get('Notes', '')).strip(' -') +
                    f' | {stage} enrichment pending: {type(exc).__name__}').strip(' |')}

    def enrich(self, record):
        maps = self.executors['maps'].submit(self._run, 'maps', record)
        if not self.concurrent:
            result = self._result(maps, 'maps', record)
            future = self.executors['company_email'].submit(self._run, 'company_email', result)
            return self._result(future, 'company_email', result)
        company = self.executors['company_email'].submit(self._run, 'company_email', record)
        maps_result = self._result(maps, 'maps', record)
        company_result = self._result(company, 'company_email', record)
        # Company/email processing cannot overwrite address/phone selected by Maps.
        output = dict(maps_result)
        contact_fields = {'Address', 'Address Link', 'Zip Code', 'Phone Number', 'Phone Number Link'}
        for key, value in company_result.items():
            if key not in contact_fields and key != 'Notes' and value != record.get(key):
                output[key] = value
        notes = []
        for result in (maps_result, company_result):
            for note in str(result.get('Notes', '')).split(' | '):
                if note not in ('', '-') and note not in notes:
                    notes.append(note)
        if notes:
            output['Notes'] = ' | '.join(notes)
        return output

    def close(self):
        try:
            futures = [executor.submit(self.workers[name].close)
                       for name, executor in self.executors.items() if name in self.workers]
            for future in futures:
                future.result()
        finally:
            for executor in self.executors.values():
                executor.shutdown(wait=True, cancel_futures=True)
