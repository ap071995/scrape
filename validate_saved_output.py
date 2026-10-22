"""Audit a live run's reopened workbook against its retained source captures."""
import json
from pathlib import Path
from urllib.parse import urlsplit
import Scraper as app


def main():
    folder = Path('Output/verification')
    evidence = json.loads((folder / 'live_evidence.json').read_text(encoding='utf-8'))
    sources = {app.normalize_url(item['profile_url']): item for item in evidence}
    workbook = app.load_workbook('Output/Link_Scraper_Results_verified.xlsx', read_only=False)
    sheet = workbook.active
    values = list(sheet.values)
    assert list(values[0]) == app.HEADERS
    records = [dict(zip(values[0], row)) for row in values[1:]]
    required = ['First Name', 'Last Name', 'Job Title', 'Tenure', 'Company Name',
                'Company Website', 'Industry', 'Company Size', 'City', 'State', 'Country']
    failures = []
    seen = set()
    for row_number, record in enumerate(records, 2):
        url = app.normalize_url(record['Prospect LinkedIn Link'])
        if url in seen:
            failures.append(f'Row {row_number}: duplicate URL')
        seen.add(url)
        source = sources.get(url)
        if not source or not source.get('company'):
            failures.append(f'Row {row_number}: no successful live profile/About evidence')
            continue
        for key in required:
            if record[key] in (None, '', '-', 'Dead'):
                failures.append(f'Row {row_number}: missing {key}')
        for key in ['First Name','Last Name','Job Title','Tenure']:
            if record[key] != source['parsed'][key]:
                failures.append(f'Row {row_number}: {key} differs from profile evidence')
        for key in ['Company Name','Company Website','Industry','Company Size','Address']:
            if record[key] != source['company'][key]:
                failures.append(f'Row {row_number}: {key} differs from About evidence')
        for key in ['Zip Code', 'Phone Number', 'Phone Number Link', 'Address Link']:
            expected = source['company'].get(key, '-')
            if expected not in (None, '', '-') and record[key] != expected:
                failures.append(f'Row {row_number}: {key} differs from authoritative About evidence')
        expected_location = app.SearchWorker._location_to_fields(
            source['parsed']['Location'], source['parsed'].get('Profile Location', '-'))
        for key, expected in zip(['City','State','Country'], expected_location):
            if expected != '-' and record[key] != expected:
                failures.append(f'Row {row_number}: {key} differs from profile location')
        phone = record['Phone Number']
        phone_link = record['Phone Number Link']
        if phone not in (None, '', '-') and phone_link not in (None, '', '-'):
            if app.SearchWorker._normalize_maps_phone(phone) != app.SearchWorker._normalize_maps_phone(phone_link.removeprefix('tel:')):
                failures.append(f'Row {row_number}: phone/link mismatch')
        if not urlsplit(record['Company Website']).hostname:
            failures.append(f'Row {row_number}: invalid website')
    workbook.close()
    original = app.load_workbook('Output/Link_Scraper_Results.xlsx', read_only=True)
    original_values = list(original.active.values)
    original.close()
    original_rows = {app.normalize_url(row[original_values[0].index('Prospect LinkedIn Link')]):
                     dict(zip(original_values[0], row)) for row in original_values[1:]}
    if seen != set(original_rows):
        failures.append('Profile set changed during repair')
    for record in records:
        previous = original_rows.get(app.normalize_url(record['Prospect LinkedIn Link']), {})
        for key in ['Campaign Name', 'Agent Name', 'Calling Agent Name', 'Source File Name']:
            if previous.get(key) not in (None, '', '-') and record[key] != previous[key]:
                failures.append(f'{record["First Name"]}: existing {key} changed')
    report = {'app_version': app.APP_VERSION, 'rows': len(records),
              'live_evidence_rows': len(sources), 'required_fields': required,
              'failures': failures, 'passed': not failures}
    (folder / 'validation_report.json').write_text(json.dumps(report, indent=2), encoding='utf-8')
    print(json.dumps(report, indent=2))
    if failures:
        raise SystemExit(1)


if __name__ == '__main__':
    main()
