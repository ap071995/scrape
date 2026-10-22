"""Offline timings for batched Excel persistence and concurrent enrichment scheduling."""
import json
from pathlib import Path
import tempfile
import time

from Scraper import WorkbookManager
from run_support import EnrichmentSessions


def main():
    results = {'rows': 100, 'website_requests': 0, 'kind': 'Synthetic offline workload'}
    with tempfile.TemporaryDirectory(prefix='runtime-benchmark-', dir='.') as directory:
        for batch in (1, 10):
            manager = WorkbookManager(str(Path(directory) / f'batch_{batch}.xlsx'), checkpoint_every=batch)
            initial_saves = manager.metrics.counts.get('excel_save_calls', 0)
            start = time.monotonic()
            for index in range(results['rows']):
                manager.add_record({'First Name': 'Fixture', 'Last Name': str(index),
                    'Company Name': 'Fixture Company',
                    'Prospect LinkedIn Link': f'https://www.linkedin.com/in/fixture-{index}',
                    'Phone Number': '+442075678000', 'Zip Code': 'EC2M 2QS'})
            manager.save()
            results[f'batch_{batch}'] = {'seconds': round(time.monotonic() - start, 3),
                'xlsx_writes': manager.metrics.counts['excel_save_calls'] - initial_saves}
            manager.wb.close()
            reopened = WorkbookManager(str(manager.output_path))
            assert reopened.ws.max_row == results['rows'] + 1
            reopened.wb.close()
    class Session:
        def __init__(self, stage):
            self.stage = stage
        def process(self, record):
            time.sleep(0.15)
            return {**record, 'Address' if self.stage == 'maps' else 'Industry': 'Fixture'}
        def close(self):
            pass
    for concurrent in (False, True):
        sessions = EnrichmentSessions(Session, concurrent=concurrent)
        start = time.monotonic()
        try:
            for _ in range(5):
                sessions.enrich({})
        finally:
            sessions.close()
        results['parallel' if concurrent else 'sequential'] = round(time.monotonic() - start, 3)
    output = Path('Output/v42.7_validation/runtime_benchmark.json')
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(results, indent=2), encoding='utf-8')
    print(json.dumps(results, indent=2))


if __name__ == '__main__':
    main()
