"""One non-LinkedIn browser per enrichment thread; no signed-in profile sharing."""
from __future__ import annotations

from urllib.parse import urlsplit
from playwright.sync_api import sync_playwright


class IsolatedEnrichmentSession:
    def __init__(self, owner, stage):
        self.stage = stage
        self.pw = sync_playwright().start()
        self.browser = None
        self.worker = type(owner)(owner.callbacks)
        self.worker.stop_event = owner.stop_event
        self.worker.pause_event = owner.pause_event
        self.worker.metrics = owner.metrics
        self.worker.prefer_local_office = owner.prefer_local_office
        self.worker.email_domain_checks_enabled = owner.email_domain_checks_enabled
        self.worker.cache_path = owner.cache_path
        try:
            try:
                self.browser = self.pw.chromium.launch(
                    channel='chrome', headless=owner.headless, chromium_sandbox=True)
            except Exception:
                self.browser = self.pw.chromium.launch(
                    headless=owner.headless, chromium_sandbox=True)
            context = self.browser.new_context(viewport={'width': 1365, 'height': 900},
                                               service_workers='block')
            # Enrichment workers never visit LinkedIn, including redirects or embeds.
            def route_request(route):
                hostname = (urlsplit(route.request.url).hostname or '').lower()
                if hostname == 'linkedin.com' or hostname.endswith('.linkedin.com'):
                    route.abort()
                else:
                    route.continue_()
            context.route('**/*', route_request)
            context.on('request', lambda req: owner.metrics.count(stage + '_navigations')
                       if req.is_navigation_request() else None)
            self.worker.context = context
            self.worker._install_caches()
        except BaseException:
            self.close()
            raise

    def process(self, record):
        worker = self.worker
        if worker.stop_event.is_set():
            return record
        with worker.metrics.measure(self.stage):
            if self.stage == 'maps':
                return worker._google_maps_enrich_record(record, 'Maps session')
            result = worker._public_company_enrich_record(record, 'Company session')
            return worker._business_email_enrich_record(result, 'Company session')

    def close(self):
        try:
            self.worker._reset_public_company_enricher()
            if self.browser:
                self.browser.close()
        finally:
            self.pw.stop()
