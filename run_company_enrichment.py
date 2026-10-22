"""Bounded live enrichment of the five saved sample rows, with source evidence."""
import dataclasses
import hashlib
import json
from pathlib import Path
import shutil
import sys

import Scraper as app


def main():
    sys.stdout.reconfigure(encoding='utf-8')
    folder = Path('Output/company_verification')
    folder.mkdir(exist_ok=True, parents=True)
    original = Path('Output/Link_Scraper_Results.xlsx')
    output = Path('Output/Link_Scraper_Results_revenue_codes.xlsx')
    original_digest = hashlib.sha256(original.read_bytes()).hexdigest()
    shutil.copy2(original, output)
    # Force fresh public-source extraction in this disposable copy, even when
    # the user has already enriched the main workbook with an earlier build.
    preparation = app.load_workbook(output)
    sheet = preparation['Leads']
    headers = [cell.value for cell in sheet[1]]
    sample_companies = set()
    for row_number in range(2, min(sheet.max_row + 1, 7)):
        sample_companies.add(sheet.cell(row_number, headers.index('Company Name') + 1).value)
        for field in app.PUBLIC_COMPANY_FIELDS:
            sheet.cell(row_number, headers.index(field) + 1).value = '-'
        note_cell = sheet.cell(row_number, headers.index('Notes') + 1)
        note_cell.value = ' | '.join(part for part in str(note_cell.value or '').split(' | ')
                                     if not part.startswith('Revenue period:')) or '-'
    if 'Company Evidence' in preparation.sheetnames:
        evidence_sheet = preparation['Company Evidence']
        for row_number in range(evidence_sheet.max_row, 1, -1):
            if evidence_sheet.cell(row_number, 1).value in sample_companies:
                evidence_sheet.delete_rows(row_number)
        # openpyxl moves cells when rows are deleted but does not relocate the
        # hyperlinks' own references. Old references recreate empty ghost rows.
        for row in evidence_sheet:
            for cell in row:
                if cell.hyperlink:
                    cell.hyperlink.ref = cell.coordinate
    preparation.save(output)
    preparation.close()
    workbook = app.WorkbookManager(str(output))
    before = [dict(zip(app.HEADERS, row)) for row in list(workbook.ws.values)[1:]]
    sample = before[:5]
    source = folder / 'profiles.csv'
    source.write_text('URL\n' + '\n'.join(row['Prospect LinkedIn Link'] for row in sample), encoding='utf-8')
    log_file = folder / 'live_run.log'
    log_file.write_text('', encoding='utf-8')
    def log(message):
        print(message, flush=True)
        with log_file.open('a', encoding='utf-8') as stream:
            stream.write(message + '\n')
    callbacks = {key: (lambda *args: None) for key in ['progress','record','harvest','done']}
    callbacks['log'] = log
    callbacks['trace'] = lambda *args, **kwargs: log(json.dumps([args,kwargs], ensure_ascii=True))
    worker = app.SearchWorker(callbacks)
    # Existing complete rows must use the new enrichment-only path.
    def no_profile(*args):
        raise AssertionError('Unexpected LinkedIn re-scrape during company-only repair')
    worker._capture_standard_profile = no_profile
    # This runner audits public company facts only. Maps has its own live runner.
    worker._google_maps_enrich_record = lambda record, **kwargs: record
    worker._business_email_needs_enrichment = lambda record: False
    worker._business_email_enrich_record = lambda record, **kwargs: record
    worker._launch_browser = lambda: None
    with app.sync_playwright() as pw:
        browser = pw.chromium.launch(channel='chrome', headless=True)
        worker.context = browser.new_context()
        worker.page = worker.context.new_page()
        try:
            worker._process_link_file({'input_path':str(source), 'max_records':len(sample),
                                      'workbook':workbook, 'extraction_mode':'FAST'})
            service = worker.public_company_enricher
            documents = {url:dataclasses.asdict(doc) if doc else None for url,doc in service.documents.items()} if service else {}
            (folder / 'source_documents.json').write_text(json.dumps(documents, ensure_ascii=False, indent=2), encoding='utf-8')
        finally:
            worker._reset_public_company_enricher()
            browser.close()
            workbook.wb.close()
    reopened = app.load_workbook(output, read_only=True)
    after = [dict(zip(app.HEADERS, row)) for row in list(reopened['Leads'].values)[1:]]
    evidence = [dict(zip(app.COMPANY_EVIDENCE_HEADERS,row)) for row in list(reopened['Company Evidence'].values)[1:]] if 'Company Evidence' in reopened.sheetnames else []
    reopened.close()
    assert all(entry['Company Name'] and entry['Field'] and entry['Value'] and entry['Source URL']
               for entry in evidence), 'Incomplete or ghost evidence row'
    assert len(after) == len(before)
    for old,new in zip(before[:5], after[:5]):
        for key in app.HEADERS:
            if key not in app.PUBLIC_COMPANY_FIELDS and key != 'Notes':
                assert old[key] == new[key], (key,old[key],new[key])
        for key in ['Revenue','SIC Code','NAICS Code']:
            if not app.SearchWorker._is_unknown_value(new[key]):
                assert any(e['Company Name']==new['Company Name'] and e['Field']==key and e['Value']==new[key]
                           for e in evidence), (new['Company Name'],key,'missing evidence')
    assert before[5:] == after[5:], 'Rows outside the validation sample changed'
    assert hashlib.sha256(original.read_bytes()).hexdigest() == original_digest, 'Original workbook changed during validation'
    result = {'version':app.APP_VERSION, 'rows':len(after), 'sample_rows':len(sample), 'core_fields_preserved':True,
              'original_workbook_preserved': True,
              'sample_sourced_fields': sum(not app.SearchWorker._is_unknown_value(row.get(field))
                                          for row in after[:5] for field in ['Revenue', 'SIC Code', 'NAICS Code']),
              'evidence_rows':len(evidence), 'records':[
                  {key:row[key] for key in ['Company Name','Revenue','Revenue Link','SIC Code','NAICS Code','SIC/NAICS Code Link','Notes']}
                  for row in after[:5]]}
    (folder / 'validation_report.json').write_text(json.dumps(result,indent=2,ensure_ascii=False),encoding='utf-8')
    print(json.dumps(result,indent=2,ensure_ascii=True))


if __name__ == '__main__':
    main()
