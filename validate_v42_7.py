"""Small, explicit live Maps check. Uses no LinkedIn session or credentials."""
import argparse
import json
from pathlib import Path
import time

from playwright.sync_api import sync_playwright

from Scraper import SearchWorker, WorkbookManager, HEADERS


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--live-maps', action='store_true', required=True)
    parser.add_argument('--output-dir', default='Output/v42.7_validation/live_maps')
    args = parser.parse_args()
    directory = Path(args.output_dir)
    directory.mkdir(parents=True, exist_ok=True)
    events = []
    worker = SearchWorker({'log': lambda message: print(message, flush=True),
                           'trace': lambda category, data: events.append({'category': category, **data})})
    worker.prefer_local_office = True
    record = {**{key: '-' for key in HEADERS},
              'First Name': 'Sejal', 'Last Name': 'Pattni', 'Company Name': 'UBS',
              'Company Website': 'https://www.ubs.com/',
              'Prospect LinkedIn Link': 'https://www.linkedin.com/in/sejal-pattni-1234124',
              'City': 'London', 'State': 'England', 'Country': 'United Kingdom',
              'Address': 'Postbox, Zurich, 8001, CH', 'Zip Code': '8001',
              'Address Link': 'https://www.linkedin.com/company/1214/about/',
              'Notes': 'Validation input: supplied screenshot and saved profile context; LinkedIn not revisited'}
    start = time.monotonic()
    with sync_playwright() as pw:
        browser = pw.chromium.launch(channel='chrome', headless=True, chromium_sandbox=True)
        context = browser.new_context(viewport={'width': 1365, 'height': 900})
        worker.context = context
        try:
            result = worker._google_maps_enrich_record(record, 'UBS screenshot validation')
            if worker.google_maps_page:
                worker.google_maps_page.screenshot(path=str(directory / 'ubs_maps.png'))
        finally:
            browser.close()
    manager = WorkbookManager(str(directory / 'UBS_Maps_validation.xlsx'), checkpoint_every=10)
    if manager.find_link_row(record['Prospect LinkedIn Link']):
        manager.update_record(2, result)
    else:
        manager.add_record(result)
    manager.save(force=True)
    manager.wb.close()
    reopened = WorkbookManager(str(manager.output_path))
    saved = reopened.read_record(2)
    reopened.wb.close()
    report = {'seconds': round(time.monotonic() - start, 2),
              'address': saved['Address'], 'phone': saved['Phone Number'],
              'zip': saved['Zip Code'], 'source': saved['Address Link'],
              'matches_supplied_screenshot': ('5 Broadgate' in saved['Address']
                  and saved['Phone Number'] == '+442075678000' and saved['Zip Code'] == 'EC2M 2QS'),
              'previous_company_address': saved.get('_previous_company_location', {}).get('address'),
              'linkedin_visits': 0}
    (directory / 'result.json').write_text(json.dumps(report, indent=2), encoding='utf-8')
    (directory / 'trace.json').write_text(json.dumps(events, indent=2), encoding='utf-8')
    print(json.dumps(report, indent=2), flush=True)


if __name__ == '__main__':
    main()
