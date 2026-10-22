"""Start the real desktop UI and worker without opening external websites."""
import time
import tempfile
from pathlib import Path
import unittest
from unittest.mock import patch

from test_scraper import app


class DesktopSmokeTests(unittest.TestCase):
    def test_progress_upload_and_worksheet_controls(self):
        with tempfile.TemporaryDirectory() as folder, \
             patch.object(app.SearchWorker, 'start'), \
             patch.object(app.LeadExtractorApp, '_check_llm_async'):
            window = app.LeadExtractorApp()
            window.withdraw()
            try:
                window._update_stage('email', 'running')
                self.assertTrue(window.stage_widgets['email']['running'])
                window._update_stage('email', 'done', '0.5s')
                self.assertFalse(window.stage_widgets['email']['running'])
                self.assertEqual(window.stage_widgets['email']['bar'].get(), 1)
                window._reset_stages()
                self.assertEqual(window.stage_widgets['email']['bar'].get(), 0)
                source = Path(folder) / 'input.csv'
                source.write_text('\n'.join(f'https://www.linkedin.com/in/person-{i}' for i in range(70)))
                window._preview_link_input(str(source))
                deadline = time.monotonic() + 5
                while window.importing and time.monotonic() < deadline:
                    window.update()
                    time.sleep(0.01)
                self.assertFalse(window.importing)
                self.assertIn('70 unique', window.link_count_label.cget('text'))
                window.output_path = str(Path(folder) / 'output.xlsx')
                window._open_worksheet()
                editor = window.worksheet_window
                editor.withdraw()
                with patch('worksheet_editor.simpledialog.askstring', return_value='My column'):
                    editor.add_column()
                with patch('worksheet_editor.simpledialog.askinteger', return_value=2):
                    editor.add_rows()
                self.assertEqual(editor.book.headers[-1], 'My column')
                self.assertEqual(len(editor.table.get_children()), 2)
                editor.close()
                window.link_input_path = str(source)
                window.all_links_option.select()
                with patch.object(window.worker, 'send') as send:
                    window._start_processing()
                    self.assertIsNone(send.call_args.args[1]['max_records'])
                window._processing_done()
                self.assertEqual(window.workbook.headers[-1], 'My column')
                window.workbook.wb.close()
            finally:
                window._on_close()

    def test_window_modes_and_worker_shutdown(self):
        # The optional local LLM service is independent of window construction.
        with patch.object(app.LeadExtractorApp, '_check_llm_async'):
            window = app.LeadExtractorApp()
        callback_errors = []
        window.report_callback_exception = lambda *error: callback_errors.append(error)
        window.withdraw()
        worker = window.worker
        try:
            window.update()
            self.assertEqual(window.title(), f'{app.APP_NAME} v{app.APP_VERSION}')
            self.assertEqual(window.email_domain_checks.get(), 0)
            self.assertEqual(window.headless_option.get(), 0)
            self.assertEqual(window.parallel_enrichment.get(), 1)
            self.assertEqual(window.local_office_option.get(), 1)
            self.assertEqual(window.refresh_profiles_option.get(), 0)
            window.headless_option.select()
            self.assertEqual(window.headless_option.get(), 1)
            window.email_domain_checks.select()
            self.assertEqual(window.email_domain_checks.get(), 1)
            self.assertEqual(window.benchmark_sample_entry.get(), '150')
            self.assertEqual(window.benchmark_button.cget('text'), 'Build Offline Benchmark')
            for mode, active, hidden, button in (
                ('Link Scraper', window.link_input_frame, window.sales_input_frame, 'START LINK SCRAPER'),
                ('Sales Navigator', window.sales_input_frame, window.link_input_frame, 'START SALES NAVIGATOR'),
            ):
                window.process_mode.set(mode)
                window._on_process_mode_change(mode)
                window.update()
                self.assertEqual(active.winfo_manager(), 'pack')
                self.assertEqual(hidden.winfo_manager(), '')
                self.assertEqual(window.start_btn.cget('text'), button)
            deadline = time.monotonic() + 10
            while worker.pw is None and worker.is_alive() and time.monotonic() < deadline:
                window.update()
                time.sleep(0.05)
            self.assertIsNotNone(worker.pw, 'Playwright worker failed to initialize')
            self.assertTrue(worker.is_alive())
            self.assertEqual(callback_errors, [])
        finally:
            window._on_close()
            worker.join(timeout=10)
        self.assertFalse(worker.is_alive(), 'Worker did not shut down with the application')


if __name__ == '__main__':
    unittest.main(verbosity=2)
