"""Run the existing workbook's profiles through the app and retain source evidence."""
import json
import shutil
import sys
from pathlib import Path
import Scraper as app


def main():
    sys.stdout.reconfigure(encoding='utf-8')
    folder = Path('Output/verification')
    folder.mkdir(parents=True, exist_ok=True)
    original = Path('Output/Link_Scraper_Results.xlsx')
    output = Path('Output/Link_Scraper_Results_verified.xlsx')
    if not output.exists():
        shutil.copy2(original, output)
    read = app.load_workbook(original, read_only=True)
    rows = list(read.active.values)
    read.close()
    index = rows[0].index('Prospect LinkedIn Link')
    urls = [row[index] for row in rows[1:] if row[index]]
    source = folder / 'validation_profiles.csv'
    source.write_text('URL\n' + '\n'.join(urls) + '\n', encoding='utf-8')
    evidence = []
    log_path = folder / 'live_run.log'
    def log(message):
        print(message, flush=True)
        with log_path.open('a', encoding='utf-8') as stream:
            stream.write(str(message) + '\n')
    callbacks = {key: (lambda *args: None) for key in ['progress','record','harvest','done']}
    callbacks['log'] = log
    callbacks['trace'] = lambda *args, **kwargs: log(json.dumps([args,kwargs], ensure_ascii=True))
    worker = app.SearchWorker(callbacks)
    capture = worker._capture_standard_profile
    company = worker._read_standard_company_about
    def capture_evidence(url):
        data = capture(url)
        item = {'profile_url': url, 'capture': data, 'parsed': worker._parse_standard_current_experience(data)}
        evidence.append(item)
        (folder / f'profile_{len(evidence)}.html').write_text(worker.page.locator('main').inner_html(), encoding='utf-8')
        return data
    def company_evidence(*args, **kwargs):
        fields = company(*args, **kwargs)
        evidence[-1]['company'] = fields
        (folder / f'company_{len(evidence)}.html').write_text(worker.page.locator('main').inner_html(), encoding='utf-8')
        return fields
    worker._capture_standard_profile = capture_evidence
    worker._read_standard_company_about = company_evidence
    manager = app.WorkbookManager(str(output))
    with app.sync_playwright() as pw:
        worker.pw = pw
        worker.context = pw.chromium.launch_persistent_context(
            str(app.CHROME_PROFILE_DIR), channel='chrome', headless=True,
            viewport={'width': 1366, 'height': 900},
        )
        worker.page = worker.context.new_page()
        try:
            worker._process_link_file({'input_path': str(source), 'max_records': len(urls),
                                       'workbook': manager, 'extraction_mode': 'FAST'})
        finally:
            (folder / 'live_evidence.json').write_text(json.dumps(evidence, indent=2, ensure_ascii=False), encoding='utf-8')
            worker.context.close()
            manager.wb.close()
    check = app.load_workbook(output, read_only=True)
    values = list(check.active.values)
    check.close()
    records = [dict(zip(values[0], row)) for row in values[1:]]
    (folder / 'output_records.json').write_text(json.dumps(records, indent=2, ensure_ascii=False), encoding='utf-8')
    for record in records:
        print(json.dumps({key:record[key] for key in ['First Name','Last Name','Job Title','Tenure','Company Name',
              'Company Website','Industry','Company Size','City','State','Country']}, ensure_ascii=True))


if __name__ == '__main__':
    main()
