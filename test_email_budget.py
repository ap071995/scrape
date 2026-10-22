"""Deadline tests with a fake clock and a real sandboxed local browser fixture."""
from contextlib import contextmanager
import time
import unittest
from unittest.mock import Mock, patch

from business_email import BusinessEmailEnricher
from company_public_data import PublicCompanyEnricher
from test_business_email import record, company_match, document, search_row


class EmailBudgetTests(unittest.TestCase):
    def test_timeout_is_retryable_and_does_not_generate(self):
        now = [0]
        web = Mock()
        web.documents = {}
        web.source_index = {}
        def search(query):
            now[0] += 5
            return []
        web.search.side_effect = search
        enricher = BusinessEmailEnricher(web, company_match, clock=lambda: now[0])
        result = enricher.enrich(record())
        self.assertEqual(result['Email Status'], 'Time limit reached (unverified)')
        self.assertEqual(web.search.call_count, 3)
        self.assertFalse(enricher.cache)
        self.assertFalse(enricher.pattern_cache)
        enricher.enrich(record())
        self.assertEqual(web.search.call_count, 6)

    def test_evidence_found_before_deadline_is_preserved(self):
        now = [0]
        doc = document('Jane Doe\nEmail: jane.doe@example.com')
        web = Mock()
        web.documents, web.source_index = {}, {}
        def search(query):
            now[0] = 15
            return [search_row(doc)]
        web.search.side_effect = search
        enricher = BusinessEmailEnricher(web, company_match, clock=lambda: now[0])
        result = enricher.enrich(record())
        self.assertEqual(result['Business Email ID'], 'jane.doe@example.com')
        self.assertEqual(len(result['_email_evidence']), 1)
        web.fetch.assert_not_called()

    def test_budget_restores_web_state_on_exception(self):
        web = PublicCompanyEnricher(Mock(), company_match)
        with self.assertRaises(RuntimeError):
            with web.time_budget(15):
                self.assertLessEqual(web._timeout(18000), 15000)
                raise RuntimeError()
        self.assertIsNone(web.deadline)
        self.assertEqual(web._timeout(18000), 18000)

    def test_configured_search_api_uses_remaining_budget(self):
        context = Mock()
        response = context.request.get.return_value
        response.ok = True
        response.json.return_value = {'web': {'results': [{
            'url': 'https://example.com/staff', 'title': 'Jane Doe',
            'description': 'Jane Doe email: jane.doe@example.com'}]}}
        web = PublicCompanyEnricher(context, company_match)
        with patch.dict('os.environ', {'BRAVE_SEARCH_API_KEY': 'fixture-key'}):
            with web.time_budget(1):
                rows = web.search('Jane Doe email')
        self.assertEqual(len(rows), 1)
        self.assertLessEqual(context.request.get.call_args.kwargs['timeout'], 1000)
        response.dispose.assert_called_once()

    def test_real_browser_delayed_navigation_respects_total_budget(self):
        from playwright.sync_api import sync_playwright
        # A local server accepts connections but intentionally delays the page.
        from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
        import threading
        class Handler(BaseHTTPRequestHandler):
            def do_GET(self):
                time.sleep(3)
                try:
                    self.send_response(200)
                    self.end_headers()
                    self.wfile.write(b'<main>Delayed fixture</main>')
                except (BrokenPipeError, ConnectionResetError, ConnectionAbortedError):
                    pass
            def log_message(self, *args):
                pass
        server = ThreadingHTTPServer(('127.0.0.1', 0), Handler)
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        self.addCleanup(server.server_close)
        self.addCleanup(server.shutdown)
        with sync_playwright() as pw:
            browser = pw.chromium.launch(channel='chrome', headless=True, chromium_sandbox=True)
            context = browser.new_context()
            web = PublicCompanyEnricher(context, company_match)
            page = context.new_page()
            # Exercise the actual goto timeout against a local server; no external request.
            class SlowPage:
                def is_closed(self):
                    return page.is_closed()
                def goto(self, url, **kwargs):
                    return page.goto(f'http://127.0.0.1:{server.server_port}/', **kwargs)
                def close(self):
                    page.close()
            web.page = SlowPage()
            started = time.monotonic()
            with web.time_budget(0.8):
                web.fetch('https://example.com/staff')
            self.assertLess(time.monotonic() - started, 2.5)
            self.assertNotIn('https://example.com/staff', web.documents)
            browser.close()


if __name__ == '__main__':
    unittest.main(verbosity=2)
