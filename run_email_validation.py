"""Test public email discovery on up to five saved prospects in a separate XLSX.

Usage: python run_email_validation.py --limit 5
Original workbook and all unrelated cells/sheets must remain unchanged.
"""
import argparse
from collections import Counter
import dataclasses
import hashlib
import json
from pathlib import Path
import shutil
import sys

import Scraper as app
from business_email import FOUND, GENERATED, valid_email
from run_maps_validation import snapshot


EMAIL_FIELDS = {'Business Email ID', 'Email Status', 'Bounce Email Status', 'Notes'}


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def assert_unselected_email_details_preserved(old_rows, new_rows, selected_records):
    """Email detail rows may change only for profiles selected by this run."""
    assert old_rows and new_rows and old_rows[0] == new_rows[0], 'Email Details headers changed'

    def key(row):
        return (app.normalize_url(str(row[0][0] or '')),
                str(row[1][0] or '').strip().casefold())

    selected = {(app.normalize_url(row['Prospect LinkedIn Link']),
                 row['Company Name'].strip().casefold()) for row in selected_records}
    old = {key(row): row for row in old_rows[1:]}
    new = {key(row): row for row in new_rows[1:]}
    assert len(old) == len(old_rows) - 1 and len(new) == len(new_rows) - 1, 'Duplicate email detail keys'
    for identity, row in old.items():
        if identity not in selected:
            assert new.get(identity) == row, 'Unselected Email Details row changed'
    assert not (new.keys() - old.keys() - selected), 'Unexpected Email Details row added'


def main():
    base = Path(__file__).resolve().parent
    preferred = base / 'Output/Link_Scraper_Results_maps_verified.xlsx'
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--limit', type=int, choices=range(1, 6), default=5)
    parser.add_argument('--input', type=Path, default=preferred if preferred.exists()
                        else base / 'Output/Link_Scraper_Results.xlsx')
    parser.add_argument('--output', type=Path, default=base / 'Output/Link_Scraper_Results_email_verified.xlsx')
    args = parser.parse_args()
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    original, output = args.input.resolve(), args.output.resolve()
    if original == output:
        raise ValueError('Validation output must differ from input.')
    folder = output.parent / 'email_verification'
    folder.mkdir(parents=True, exist_ok=True)
    original_digest = digest(original)
    shutil.copy2(original, output)
    workbook = app.WorkbookManager(str(output))
    before = snapshot(workbook.wb)
    before_rows = [dict(zip(app.HEADERS, [app.clean_cell_value(cell.value) for cell in row]))
                   for row in workbook.ws.iter_rows(min_row=2)]
    sample, seen = [], set()
    for row_number, record in enumerate(before_rows, 2):
        key = tuple(record[field].casefold() for field in
                    ('First Name', 'Last Name', 'Company Name', 'Prospect LinkedIn Link'))
        if key in seen or any(app.SearchWorker._is_unknown_value(record[field])
                             for field in ('First Name', 'Last Name', 'Company Name')):
            continue
        seen.add(key)
        sample.append((row_number, record))
        if len(sample) >= args.limit:
            break
    log_path = folder / 'live_run.log'
    log_path.write_text('', encoding='utf-8')
    events, results = [], []
    active = {'excel_row': None}

    def log(message):
        print(message, flush=True)
        with log_path.open('a', encoding='utf-8') as stream:
            stream.write(str(message) + '\n')

    def trace(category, payload):
        entry = {'category': category, **active, **payload}
        events.append(entry)
        log(json.dumps(entry, ensure_ascii=True, default=str))

    worker = app.SearchWorker({'log': log, 'trace': trace})
    with app.sync_playwright() as pw:
        browser = pw.chromium.launch(channel='chrome', headless=True)
        worker.context = browser.new_context()
        try:
            for row_number, record in sample:
                active['excel_row'] = row_number
                log(f'EMAIL SAMPLE | row={row_number} | {record["First Name"]} {record["Last Name"]} | {record["Company Name"]}')
                enriched = worker._business_email_enrich_record(record, 'Email validation')
                workbook.update_record(row_number, enriched)
                workbook.save(force=True)
                results.append({'excel_row': row_number,
                                'record': {field: enriched.get(field, '-') for field in
                                           ('First Name', 'Last Name', 'Company Name', 'Company Website',
                                            'Business Email ID', 'Email Status')},
                                'evidence': enriched.get('_email_evidence', [])})
        finally:
            service = worker.public_company_enricher
            documents = {url: dataclasses.asdict(doc) if doc else None
                         for url, doc in service.documents.items()} if service else {}
            (folder / 'source_documents.json').write_text(json.dumps(documents, ensure_ascii=False, indent=2), encoding='utf-8')
            (folder / 'events.json').write_text(json.dumps(events, ensure_ascii=False, indent=2), encoding='utf-8')
            worker._reset_public_company_enricher()
            browser.close()
            workbook.wb.close()

    reopened = app.load_workbook(output)
    try:
        after = snapshot(reopened)
        after_rows = [dict(zip(app.HEADERS, [app.clean_cell_value(cell.value) for cell in row]))
                      for row in reopened['Leads'].iter_rows(min_row=2)]
        evidence = [dict(zip(app.EMAIL_EVIDENCE_HEADERS, row)) for row in
                    reopened['Email Evidence'].iter_rows(min_row=2, values_only=True)] if 'Email Evidence' in reopened.sheetnames else []
    finally:
        reopened.close()
    assert digest(original) == original_digest, 'Original workbook was modified'
    assert len(after_rows) == len(before_rows), 'Lead rows added or lost'
    sample_indices = {row_number for row_number, _ in sample}
    for sheet_name, rows in before.items():
        if sheet_name == 'Email Evidence':
            continue
        if sheet_name == 'Email Details':
            assert sheet_name in after, 'Email Details sheet removed'
            assert_unselected_email_details_preserved(
                rows, after[sheet_name], [record for _, record in sample])
            continue
        assert sheet_name in after and len(rows) == len(after[sheet_name]), sheet_name
        for row_index, (old, new) in enumerate(zip(rows, after[sheet_name]), 1):
            assert len(old) == len(new), (sheet_name, row_index)
            for column, (old_cell, new_cell) in enumerate(zip(old, new)):
                if sheet_name == 'Leads' and row_index in sample_indices and app.HEADERS[column] in EMAIL_FIELDS:
                    continue
                assert old_cell == new_cell, (sheet_name, row_index, column + 1, old_cell, new_cell)
    for result in results:
        row = after_rows[result['excel_row'] - 2]
        assert all(row[field] == result['record'][field] for field in result['record'])
        for entry in result['evidence']:
            assert valid_email(entry['email']) and entry['url'].startswith(('https://', 'http://'))
            assert entry['excerpt'] and any(saved['Business Email ID'] == entry['email']
                                           and saved['Source URL'] == entry['url'] for saved in evidence)
        old = before_rows[result['excel_row'] - 2]
        if row['Email Status'] in (FOUND, GENERATED) and old['Business Email ID'] != row['Business Email ID']:
            assert result['evidence'], 'New email without supporting evidence'
        if row['Email Status'] == GENERATED:
            assert 'unverified' in row['Email Status']
    report = {'version': app.APP_VERSION, 'input': str(original), 'output': str(output),
              'sample_rows': len(sample), 'total_rows': len(after_rows),
              'original_workbook_preserved': True, 'unrelated_cells_preserved': True,
              'duplicate_rows_added': 0, 'evidence_rows': len(evidence),
              'statuses': dict(Counter(result['record']['Email Status'] for result in results)),
              'source_errors': sum(event.get('event') in ('source_error', 'search_error', 'source_blocked', 'source_unavailable')
                                   for event in events), 'results': results}
    (folder / 'validation_report.json').write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding='utf-8')
    print(json.dumps(report, indent=2, ensure_ascii=True))


if __name__ == '__main__':
    main()
