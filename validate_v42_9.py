"""Reproducible offline capacity/deadline validation; never visits LinkedIn."""
import csv
import json
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from unittest.mock import patch

import Scraper as app
from business_email import BusinessEmailEnricher
from company_public_data import PublicCompanyEnricher
from link_import import load_link_inputs


def main():
    folder = Path(__file__).resolve().parent / 'Output' / 'v42.9.0_validation'
    folder.mkdir(parents=True, exist_ok=True)
    result = {'external_requests': 0, 'fixture_data_only': True, 'imports': {}}
    count = 98084
    for kind in ('csv', 'xlsx'):
        path = folder / ('capacity_fixture.' + kind)
        links = ([f'https://www.linkedin.com/in/capacity-fixture-{i}'] for i in range(count))
        if kind == 'csv':
            with path.open('w', encoding='utf-8', newline='') as stream:
                csv.writer(stream).writerows(links)
        else:
            wb = app.Workbook(write_only=True)
            sheet = wb.create_sheet('Profiles')
            for row in links:
                sheet.append(row)
            wb.save(path)
            wb.close()
        started = time.monotonic()
        found = load_link_inputs(path)
        elapsed = time.monotonic() - started
        assert len(found) == count
        result['imports'][kind] = {'links': len(found), 'seconds': round(elapsed, 3)}
    path = folder / 'worksheet_fixture.xlsx'
    if path.exists():
        raise FileExistsError('Preserve previous validation output; choose a fresh validation folder.')
    manager = app.WorkbookManager(str(path), checkpoint_every=100)
    manager.add_columns(['Territory', 'User Notes'])
    for i in range(75):
        manager.add_record({'First Name': 'Fixture', 'Last Name': str(i),
            'Prospect LinkedIn Link': f'https://www.linkedin.com/in/capacity-fixture-{i}',
            'Territory': 'Manual territory', 'User Notes': 'Offline test data'})
    manager.add_manual_rows(2)
    manager.update_record(77, {'First Name': 'User entered', 'Territory': 'West'})
    manager.save(force=True)
    manager.wb.close()
    reopened = app.WorkbookManager(str(path))
    assert reopened.ws.max_row == 78
    assert reopened.read_record(77)['Territory'] == 'West'
    result['worksheet'] = {'file': str(path), 'data_rows': 77, 'columns': len(reopened.headers)}
    reopened.wb.close()

    class Handler(BaseHTTPRequestHandler):
        def do_GET(self):
            time.sleep(25)
            try:
                self.send_response(200)
                self.end_headers()
                self.wfile.write(b'<main>Slow local fixture</main>')
            except OSError:
                pass
        def log_message(self, *args):
            pass
    server = ThreadingHTTPServer(('127.0.0.1', 0), Handler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    try:
        with app.sync_playwright() as pw:
            browser = pw.chromium.launch(channel='chrome', headless=True, chromium_sandbox=True)
            context = browser.new_context()
            web = PublicCompanyEnricher(context, app.SearchWorker._link_company_names_match)
            class LocalPage:
                def __init__(self):
                    self.actual = context.new_page()
                def goto(self, url, **options):
                    return self.actual.goto(f'http://127.0.0.1:{server.server_port}/', **options)
                def __getattr__(self, name):
                    return getattr(self.actual, name)
            def page():
                if web.page is None or web.page.is_closed():
                    web.page = LocalPage()
                return web.page
            web._page = page
            enricher = BusinessEmailEnricher(web, app.SearchWorker._link_company_names_match)
            started = time.monotonic()
            with patch.dict('os.environ', {'BRAVE_SEARCH_API_KEY': ''}):
                enriched = enricher.enrich({'First Name': 'Fixture', 'Last Name': 'Person',
                    'Company Name': 'Example Industries', 'Company Website': 'https://example.com'})
            elapsed = time.monotonic() - started
            assert elapsed < 17, elapsed
            assert enriched['Email Status'] == 'Time limit reached (unverified)', enriched
            assert not enricher.cache
            result['email_deadline'] = {'budget_seconds': 15, 'elapsed_seconds': round(elapsed, 3),
                                        'status': enriched['Email Status']}
            browser.close()
    finally:
        server.shutdown()
        server.server_close()
    (folder / 'validation_results.json').write_text(json.dumps(result, indent=2), encoding='utf-8')
    print(json.dumps(result, indent=2))


if __name__ == '__main__':
    main()
