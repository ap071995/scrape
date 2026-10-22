# Scraper v42.8.2 — browser sandboxing

Chromium sandboxing is explicitly enabled for all four application launch paths:
the persistent LinkedIn Chrome session, its Chromium fallback, enrichment Chrome
sessions, and their Chromium fallback. The setting applies in both visible and
headless modes. If sandboxed startup fails, the application does not retry with
sandboxing disabled. The windowless launcher remains available.

## Verification

**81 checks passed**, with no failures or skips:

- 48 sandbox, run-support and mixed-link tests in 26.927 seconds.
- 33 browser extraction regression tests in 24.768 seconds.

Real headless Chrome instances were launched through the application's LinkedIn
session code and both enrichment stages. Their own `chrome://version` diagnostic
pages confirmed that `--no-sandbox` and `--disable-gpu-sandbox` were absent. Each
browser successfully rendered and read a local test page. Failure/fallback tests
confirmed that sandboxing stays enabled in both visible and headless settings.

The checks also cover isolated concurrent enrichment sessions, local Maps and
email fixtures, company extraction, mixed-link resolution and Excel output.
Browser fixtures and the shipped Maps validator now use the same sandbox option.
Python compilation and the actual windowless launcher preflight passed.

Logs are in `Output/v42.8.2_validation/sandbox_tests.log` and
`Output/v42.8.2_validation/browser_regressions.log`.

These checks used local browser fixtures and did not scrape live LinkedIn
profiles or alter user output workbooks. The entire historical test suite was
not rerun for this browser-launch setting change. This is a configuration and
startup validation, not a comprehensive operating-system security audit.

## Apply the update

Close the old application and its Chrome windows, then start this release with
Run_Scraper.vbs. Existing browser processes retain their old flags until closed.
The application's saved sign-in profile is not deleted or reset.
