"""Run a bounded Maps-only check on a disposable copy of the saved workbook.

Run with ``python run_maps_validation.py``. The original workbook is never
edited. Existing phone/source values are restored when a fresh lookup cannot
verify them; the JSON report distinguishes retained values from fresh evidence.
"""
import argparse
from collections import Counter
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import re
import shutil
import sys

import Scraper as app


MAPS_FIELDS = ('Address', 'Zip Code', 'Phone Number', 'Address Link', 'Phone Number Link')
FORCED_FIELDS = ('Phone Number', 'Phone Number Link', 'Address Link')
SOURCE_KEYS = {'Address': 'address', 'Zip Code': 'zip_code', 'Phone Number': 'phone',
               'Address Link': 'address_link', 'Phone Number Link': 'phone_link'}


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def snapshot(workbook):
    """Cell values and hyperlinks, including sheets unrelated to Maps."""
    return {sheet.title: [[(cell.value, cell.hyperlink.target if cell.hyperlink else None)
                           for cell in row] for row in sheet]
            for sheet in workbook}


def normalized(value):
    return app.clean_cell_value(value)


def supported(field, value, details, worker):
    """Check agreement with a validated place, without inventing missing facts."""
    expected = normalized(details.get(SOURCE_KEYS[field], '-'))
    if worker._is_unknown_value(value) or worker._is_unknown_value(expected):
        return False
    if field == 'Address':
        return worker._maps_addresses_equivalent(value, expected)
    if field == 'Phone Number':
        return worker._normalize_maps_phone(value) == worker._normalize_maps_phone(expected)
    if field == 'Zip Code':
        left, right = re.sub(r'\s+', '', value.upper()), re.sub(r'\s+', '', expected.upper())
        if left == right:
            return True
        return bool(re.fullmatch(r'\d{5}(?:-\d{4})?', left)
                    and re.fullmatch(r'\d{5}(?:-\d{4})?', right)
                    and left[:5] == right[:5])
    return value == expected


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--limit', type=int, choices=range(1, 6), default=5,
                        help='Number of distinct companies to check (1 to 5; default 5).')
    args = parser.parse_args()
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    base = Path(__file__).resolve().parent
    original = base / 'Output/Link_Scraper_Results.xlsx'
    output = base / 'Output/Link_Scraper_Results_maps_verified.xlsx'
    folder = base / 'Output/maps_verification'
    folder.mkdir(parents=True, exist_ok=True)
    if original.resolve() == output.resolve():
        raise ValueError('Validation output must differ from the original workbook.')
    original_digest = digest(original)
    shutil.copy2(original, output)
    workbook = app.WorkbookManager(str(output))
    before = snapshot(workbook.wb)
    before_rows = [dict(zip(app.HEADERS, [normalized(cell.value) for cell in row]))
                   for row in workbook.ws.iter_rows(min_row=2)]
    sample = []
    seen = set()
    for row_number, record in enumerate(before_rows, 2):
        company = record['Company Name']
        key = re.sub(r'\s+', ' ', company.casefold()).strip()
        if app.SearchWorker._is_unknown_value(company) or key in seen:
            continue
        sample.append((row_number, record))
        seen.add(key)
        if len(sample) >= args.limit:
            break

    log_path = folder / 'live_run.log'
    log_path.write_text('', encoding='utf-8')
    events, captures, records, errors = [], [], [], []
    active = {'sample_number': 0, 'excel_row': None, 'company': '', 'capture_number': 0}

    def log(message):
        message = str(message)
        print(message, flush=True)
        with log_path.open('a', encoding='utf-8') as stream:
            stream.write(message + '\n')

    def trace(category, payload):
        entry = {'timestamp': datetime.now(timezone.utc).isoformat(), 'category': category,
                 'sample_number': active['sample_number'], 'excel_row': active['excel_row'],
                 **payload}
        events.append(entry)
        log(json.dumps(entry, ensure_ascii=True, default=str))

    worker = app.SearchWorker({'log': log, 'trace': trace})
    original_extract = worker._extract_google_maps_place

    def capture_page(page, label, place=None):
        if page is None or page.is_closed():
            return
        active['capture_number'] += 1
        filename = (f"company_{active['sample_number']:02d}_"
                    f"place_{active['capture_number']:02d}_{label}.html")
        entry = {'excel_row': active['excel_row'], 'company': active['company'],
                 'url': page.url, 'file': filename, 'place': place}
        try:
            (folder / filename).write_text(page.content(), encoding='utf-8')
            captures.append(entry)
        except Exception as exc:
            entry['error'] = f'{type(exc).__name__}: {exc}'
            errors.append(entry)
            log(f'HTML CAPTURE ERROR | {entry["error"]}')

    def extract_with_capture(*positional, **keyword):
        place = original_extract(*positional, **keyword)
        page = positional[0] if positional else keyword.get('page', worker.google_maps_page)
        capture_page(page, 'extracted', place)
        return place

    worker._extract_google_maps_place = extract_with_capture
    worker._public_company_enrich_record = lambda *_args, **_kwargs: (_ for _ in ()).throw(
        AssertionError('Public-company enrichment must not run during Maps validation'))
    try:
        with app.sync_playwright() as playwright:
            browser = playwright.chromium.launch(channel='chrome', headless=True)
            context = browser.new_context(locale='en-US', viewport={'width': 1365, 'height': 900})
            context.set_default_timeout(10000)
            context.set_default_navigation_timeout(35000)
            worker.context = context
            worker.page = context.new_page()
            try:
                for number, (row_number, original_record) in enumerate(sample, 1):
                    active.update(sample_number=number, excel_row=row_number,
                                  company=original_record['Company Name'], capture_number=0)
                    event_start = len(events)
                    # Force a fresh lookup without discarding the existing address
                    # or ZIP; those exercise the real same-office merge checks.
                    candidate = dict(original_record)
                    for field in FORCED_FIELDS:
                        candidate[field] = '-'
                    worker.google_maps_cache.clear()
                    lookup_error = None
                    try:
                        candidate = worker._google_maps_enrich_record(
                            candidate, source_label='Maps validation existing row')
                    except Exception as exc:
                        lookup_error = f'{type(exc).__name__}: {exc}'
                        errors.append({'excel_row': row_number, 'company': active['company'],
                                       'error': lookup_error})
                        log(f'MAPS VALIDATION ERROR | {lookup_error}')
                    capture_page(worker.google_maps_page, 'final')
                    row_events = events[event_start:]
                    final_event = next((event for event in reversed(row_events)
                                        if event.get('event') == 'enrichment_result'), {})
                    details = final_event.get('details', {})
                    field_status = {}
                    for field in MAPS_FIELDS:
                        candidate_value = normalized(candidate.get(field, '-'))
                        fresh = bool(details.get('company_matched')
                                     and supported(field, candidate_value, details, worker))
                        # Failed fresh probes must not erase an existing contact.
                        if worker._is_unknown_value(candidate_value):
                            candidate[field] = original_record[field]
                        final_value = normalized(candidate.get(field, '-'))
                        field_status[field] = {
                            'original': original_record[field], 'value': final_value,
                            'freshly_verified': fresh,
                            'status': ('verified' if fresh else 'missing' if worker._is_unknown_value(final_value)
                                       else 'retained_not_freshly_verified'),
                            'observed_value': normalized(details.get(SOURCE_KEYS[field], '-')),
                            'source_url': normalized(details.get('address_link', '-')),
                        }
                    # This runner updates only Maps columns. Preserve notes and
                    # every profile, campaign, and public-company field verbatim.
                    deliverable = dict(original_record)
                    deliverable.update({field: candidate[field] for field in MAPS_FIELDS})
                    workbook.update_record(row_number, deliverable)
                    workbook.save(force=True)
                    records.append({'excel_row': row_number, 'company': active['company'],
                                    'lookup_error': lookup_error, 'fields': field_status,
                                    'validated_details': details,
                                    'events': row_events,
                                    'captures': [item for item in captures if item['excel_row'] == row_number]})
            finally:
                context.close()
                browser.close()
    except Exception as exc:
        errors.append({'error': f'{type(exc).__name__}: {exc}', 'stage': 'browser_or_workbook'})
        log(f'MAPS VALIDATION STOPPED | {type(exc).__name__}: {exc}')
    finally:
        workbook.save(force=True)
        workbook.wb.close()

    reopened = app.WorkbookManager(str(output))
    try:
        after = snapshot(reopened.wb)
    finally:
        reopened.wb.close()
    failures = []
    sample_rows = {row_number for row_number, _ in sample}
    if set(before) != set(after):
        failures.append('Workbook sheets changed')
    for sheet_name, old_rows in before.items():
        new_rows = after.get(sheet_name, [])
        if len(old_rows) != len(new_rows):
            failures.append(f'{sheet_name}: row count changed')
            continue
        for row_number, (old_row, new_row) in enumerate(zip(old_rows, new_rows), 1):
            if len(old_row) != len(new_row):
                failures.append(f'{sheet_name}!{row_number}: column count changed')
                continue
            for index, (old_cell, new_cell) in enumerate(zip(old_row, new_row)):
                maps_cell = (sheet_name == 'Leads' and row_number in sample_rows
                             and app.HEADERS[index] in MAPS_FIELDS)
                if not maps_cell and old_cell != new_cell:
                    failures.append(f'{sheet_name}!{row_number}:{index + 1}: unrelated cell changed')
    for row in records:
        for field, status in row['fields'].items():
            saved = normalized(after['Leads'][row['excel_row'] - 1][app.HEADERS.index(field)][0])
            if saved != status['value']:
                failures.append(f"Row {row['excel_row']}: {field} did not survive saving")
            if status['value'] != status['original'] and not status['freshly_verified']:
                failures.append(f"Row {row['excel_row']}: {field} changed without fresh source agreement")
    original_after_digest = digest(original)
    if original_after_digest != original_digest:
        failures.append('Original workbook changed')
    totals = Counter(status['status'] for row in records for status in row['fields'].values())
    coverage = {field: dict(Counter(row['fields'][field]['status'] for row in records))
                for field in MAPS_FIELDS}
    report = {
        'version': app.APP_VERSION, 'checked_at': datetime.now(timezone.utc).isoformat(),
        'original_workbook': str(original), 'output_workbook': str(output),
        'original_sha256_before': original_digest, 'original_sha256_after': original_after_digest,
        'original_workbook_preserved': original_digest == original_after_digest,
        'total_rows': len(before_rows), 'sample_requested': len(sample), 'sample_completed': len(records),
        'untouched_cells_and_rows_preserved': not any('changed' in failure for failure in failures),
        'checks_passed': not failures, 'validation_failures': failures,
        'fresh_field_status_counts': dict(totals), 'field_coverage': coverage,
        'source_errors': errors, 'records': records,
        'coverage_note': 'Retained values are preserved from the input and are not claimed as freshly verified.',
    }
    (folder / 'live_evidence.json').write_text(json.dumps({'events': events, 'captures': captures},
                                                       indent=2, ensure_ascii=False, default=str), encoding='utf-8')
    (folder / 'validation_report.json').write_text(json.dumps(report, indent=2, ensure_ascii=False,
                                                           default=str), encoding='utf-8')
    log(json.dumps({key: value for key, value in report.items() if key != 'records'},
                   indent=2, ensure_ascii=True, default=str))
    if failures:
        raise AssertionError('; '.join(failures))
    if len(records) != len(sample):
        raise RuntimeError('Live validation did not complete the bounded sample; see validation_report.json')


if __name__ == '__main__':
    main()
