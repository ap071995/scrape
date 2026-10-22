"""Optional single-lead live resolution check. Never scrapes a bulk list."""
import argparse
import json
from pathlib import Path
import time
import Scraper as app
from linkedin_links import classify_link, LinkResolutionStore


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--live-sales-link', required=True)
    parser.add_argument('--output', default='Output/v42.8_validation/live_resolution')
    args = parser.parse_args()
    item = classify_link(args.live_sales_link)
    if item.kind != 'Sales Navigator lead':
        parser.error('Supply one valid /sales/lead/ URL.')
    folder = Path(args.output)
    folder.mkdir(parents=True, exist_ok=True)
    worker = app.SearchWorker({'log': print, 'status': print})
    worker.headless = True
    worker.link_resolutions = LinkResolutionStore(folder / 'mappings.sqlite3')
    result = {'input': item.original, 'kind': 'Single live Sales Navigator resolution'}
    started = time.monotonic()
    with app.sync_playwright() as runtime:
        worker.pw = runtime
        try:
            profile, status = worker._resolve_profile_link(item)
            result.update(resolved_url=profile, status=status)
        except Exception as exc:
            result.update(status='Access required' if isinstance(exc, app.ProfileAccessError) else 'Unresolved',
                          error=f'{type(exc).__name__}: {exc}')
        finally:
            result['seconds'] = round(time.monotonic() - started, 2)
            if worker.page:
                # Diagnostics contain rendered content, never cookies or credentials.
                worker.page.screenshot(path=str(folder / 'page.png'))
                (folder / 'page.html').write_text(worker.page.content(), encoding='utf-8')
            worker._close_browser()
    (folder / 'result.json').write_text(json.dumps(result, indent=2), encoding='utf-8')
    print(json.dumps(result, indent=2))
    return 0 if result['status'] in ('Resolved', 'Cached') else 2


if __name__ == '__main__':
    raise SystemExit(main())
