"""Bounded live diagnostics using the scraper's existing Chrome session."""
import json
from pathlib import Path
import Scraper as app


def main():
    folder = Path('Output/verification')
    folder.mkdir(parents=True, exist_ok=True)
    worker = app.SearchWorker({'log': print})
    with app.sync_playwright() as pw:
        context = pw.chromium.launch_persistent_context(
            str(app.CHROME_PROFILE_DIR), channel='chrome', headless=True,
            viewport={'width': 1366, 'height': 900},
        )
        try:
            page = context.new_page()
            worker.page = page
            worker.context = context
            data = worker._capture_standard_profile('https://www.linkedin.com/in/zach-loch-ba07b015/')
            result = {'url': page.url, 'data': data,
                      'parsed': worker._parse_standard_current_experience(data)}
            (folder / 'zach_capture.json').write_text(json.dumps(result, indent=2, ensure_ascii=False), encoding='utf-8')
            (folder / 'zach_main.html').write_text(page.locator('main').inner_html(), encoding='utf-8')
            print(json.dumps(result, ensure_ascii=True, indent=2))
            fields = worker._read_standard_company_about(
                result['parsed']['Company Profile Link'], result['parsed']['Company Name'])
            (folder / 'zach_company.json').write_text(json.dumps(fields, indent=2, ensure_ascii=False), encoding='utf-8')
            (folder / 'zach_company_main.html').write_text(page.locator('main').inner_html(), encoding='utf-8')
            print('COMPANY', json.dumps(fields, ensure_ascii=True))
        finally:
            context.close()


if __name__ == '__main__':
    main()
