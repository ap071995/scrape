"""Browser security regression checks; no external website requests."""
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from test_scraper import app
from enrichment_session import IsolatedEnrichmentSession


class BrowserSandboxTests(unittest.TestCase):
    def test_profile_fallback_keeps_sandbox_in_both_display_modes(self):
        for headless in (False, True):
            with self.subTest(headless=headless), tempfile.TemporaryDirectory() as folder:
                worker = app.SearchWorker({})
                worker.headless = headless
                worker.pw = Mock()
                context = Mock(pages=[Mock()])
                launch = worker.pw.chromium.launch_persistent_context
                launch.side_effect = [RuntimeError('Chrome unavailable'), context]
                with patch.object(app, 'CHROME_PROFILE_DIR', Path(folder)):
                    worker._launch_browser()
                self.assertEqual(launch.call_count, 2)
                for call in launch.call_args_list:
                    self.assertIs(call.kwargs['chromium_sandbox'], True)
                    self.assertIs(call.kwargs['headless'], headless)
                    self.assertNotIn('--no-sandbox', call.kwargs.get('args', []))

    def test_profile_launch_failure_does_not_retry_without_sandbox(self):
        worker = app.SearchWorker({})
        worker.pw = Mock()
        launch = worker.pw.chromium.launch_persistent_context
        launch.side_effect = RuntimeError('Sandbox startup denied')
        with tempfile.TemporaryDirectory() as folder, patch.object(app, 'CHROME_PROFILE_DIR', Path(folder)):
            with self.assertRaisesRegex(RuntimeError, 'Sandbox startup denied'):
                worker._launch_browser()
        self.assertEqual(launch.call_count, 2)
        self.assertTrue(all(call.kwargs['chromium_sandbox'] for call in launch.call_args_list))

    def test_enrichment_fallback_keeps_sandbox_in_both_display_modes(self):
        for headless in (False, True):
            with self.subTest(headless=headless), patch('enrichment_session.sync_playwright') as runtime:
                owner = app.SearchWorker({})
                owner.headless = headless
                pw = runtime.return_value.start.return_value
                pw.chromium.launch.side_effect = [RuntimeError('Chrome unavailable'), Mock()]
                session = IsolatedEnrichmentSession(owner, 'maps')
                try:
                    self.assertEqual(pw.chromium.launch.call_count, 2)
                    for call in pw.chromium.launch.call_args_list:
                        self.assertIs(call.kwargs['chromium_sandbox'], True)
                        self.assertIs(call.kwargs['headless'], headless)
                finally:
                    session.close()

    def test_enrichment_failure_cleans_up_without_disabling_sandbox(self):
        with patch('enrichment_session.sync_playwright') as runtime:
            pw = runtime.return_value.start.return_value
            pw.chromium.launch.side_effect = RuntimeError('Sandbox startup denied')
            with self.assertRaisesRegex(RuntimeError, 'Sandbox startup denied'):
                IsolatedEnrichmentSession(app.SearchWorker({}), 'company_email')
            self.assertEqual(pw.chromium.launch.call_count, 2)
            self.assertTrue(all(call.kwargs['chromium_sandbox'] for call in pw.chromium.launch.call_args_list))
            pw.stop.assert_called_once()

    def assert_sandbox_launch(self, context):
        page = context.new_page()
        try:
            # Chrome exposes its own launch arguments here even when the CDP
            # getBrowserCommandLine endpoint is disabled for this installation.
            page.goto('chrome://version/')
            arguments = page.locator('#command_line').inner_text()
            self.assertNotIn('--no-sandbox', arguments)
            self.assertNotIn('--disable-gpu-sandbox', arguments)
            page.goto('about:blank')
            page.set_content('<main><h1>Sandboxed browser ready</h1></main>')
            self.assertEqual(page.locator('h1').inner_text(), 'Sandboxed browser ready')
        finally:
            page.close()

    def test_actual_profile_browser_launch_has_no_disable_sandbox_flag(self):
        with tempfile.TemporaryDirectory() as folder, app.sync_playwright() as pw:
            worker = app.SearchWorker({})
            worker.pw = pw
            worker.headless = True
            try:
                with patch.object(app, 'CHROME_PROFILE_DIR', Path(folder)):
                    worker._launch_browser()
                self.assert_sandbox_launch(worker.context)
            finally:
                worker._close_browser()

    def test_actual_enrichment_browsers_have_no_disable_sandbox_flag(self):
        owner = app.SearchWorker({})
        owner.headless = True
        for stage in ('maps', 'company_email'):
            with self.subTest(stage=stage):
                session = IsolatedEnrichmentSession(owner, stage)
                try:
                    self.assert_sandbox_launch(session.worker.context)
                finally:
                    session.close()


if __name__ == '__main__':
    unittest.main(verbosity=2)
