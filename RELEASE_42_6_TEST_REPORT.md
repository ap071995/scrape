# Scraper v42.6.0 — validation and usage

Validated on **14 October 2026** on Windows with Python 3.14, Chrome and Playwright.

## Result

**474 tests passed, zero failures and zero skips** in the final combined run
(57.310 seconds). This includes the previous 354 tests and **120 new feature tests**.
Runtime compilation and launcher preflight passed. The real desktop smoke test
checked startup, both scraping modes, the new default-off DNS option, benchmark
controls and clean worker shutdown.

The new test files are `test_email_domain_check.py`, `test_maps_office_scope.py`,
`test_benchmark.py`, and `test_feature_integration.py`. Their coverage includes:

- Published, generated and unknown-origin emails remain separate after saving,
  reopening, retries and replacement by a published address.
- Disabled DNS checks make no requests. Optional checks cover MX, null MX,
  NXDOMAIN, A/AAAA fallback, malformed responses, timeouts, caching and expiry.
  DNS failures preserve contact data; no result claims mailbox verification.
- Fresh and cached Maps candidates retain company/prospect geography separately.
  Explicit office descriptors classify headquarters or local offices; uncertain,
  conflicting and phone-only candidates are held. Exact names cannot override
  a conflicting website in the actual collector.
- New metadata sheets reject incorrect headers before changing lead cells.
  Repeated details upsert, preserve unchanged historical DNS checks and reject
  supplied details for a different email.
- Benchmark sampling, duplicate handling, human-review denominators, explicit
  absence, false matches, baseline comparisons, formula-safe exports and input
  preservation. Negative email-status wording and stale companion metadata cannot
  inflate published/generated coverage.

Final test log: `Output/v42.6_validation/full_tests.log`.

## What changed in the application

### Email

**Email Details** keeps Published Email ID, Generated Email Candidate and Existing
Email (origin unknown) in separate columns. The original 42-column Leads format
is retained. **Check email domains (DNS/MX)** is optional and off by default.
It sends only domains to Google Public DNS over HTTPS; it does not visit LinkedIn,
send messages or probe SMTP. **Mailbox Verification remains Not verified**.

A fresh two-domain check on 14 October returned:

| Domain | DNS result |
| --- | --- |
| example.com | Does not accept mail (null MX) |
| ecu.edu | MX records found |

These results validate the resolver integration, not individual email addresses.
The raw checks are in `Output/v42.6_validation/dns_live_check.json`.

### Maps

**Maps Evidence** records location type, company address/phone, prospect location,
classification/source evidence, merge eligibility, outcome and applied fields.
Prospect locality can guide a search but never proves employment at that office.
Headquarters may differ from the prospect's location when company-address evidence
agrees. Registered-office-only evidence, missing address/source or contradictory
identity/geography holds the contact for review. Existing reliable fields remain
preserved; cached candidates are revalidated before use.

### Benchmark

Click **Build Offline Benchmark** under Output Workbook and select saved results.
The default target is 150 profiles; valid targets are 1–200. A run creates a new
folder containing a review workbook, Markdown report, JSON/CSV field results and
a CSV review template. It makes **zero website requests**.

Coverage is the fraction filled. Accuracy is computed only for explicitly reviewed
known expectations. Yellow columns in the Review sheet accept checked values,
confirmed absence and source notes. Save and rerun with `--reviews` to recalculate.
The sample favors variety; it is not a weighted population accuracy estimate.

## Saved-data validation

The available saved enriched sample contained **five unique profiles**. It was
copied from `Output/strict_validation/Link_Scraper_Results_checked.xlsx` to
`Output/v42.6_validation/Saved_results_v42.6.xlsx`. Application persistence added
Email Details with DNS disabled. All original cells, hyperlinks and rows were
checked for preservation, and the source hash remained unchanged.

The offline benchmark requested 150 and correctly returned **5**, reporting a
**145-profile shortfall**. It did not invent or fetch additional profiles.
There were **zero manually reviewed known cells**, so accuracy is **Not evaluated**.
The saved sample has one published email and two generated candidates; their
coverage is shown separately. Missing financial/classification fields and missing
office-scope evidence remain visible in the report.

This historical sample was not freshly scraped from LinkedIn or Maps for this
release. It tests data handling and reporting. The office-classification changes
were exercised with controlled browser/pipeline fixtures. A full 100–200-profile
real-data accuracy benchmark requires more saved profiles and completed reviews.

Artifacts and source hashes are recorded in:

- `Output/v42.6_validation/offline_validation.json`
- `Output/v42.6_validation/benchmark_console.json` (paths to the final report)
- `Output/v42.6_validation/benchmarks/` (each run gets a separate folder)

## Efficiency and account traffic

Live request pacing, scraping limits and challenge handling were not accelerated.
The benchmark works from disk. DNS caches repeated domains with bounded lifetime;
metadata row indexes avoid scanning a growing sheet for every update. Existing
company and office caches continue to reduce repeated lookups, and optional DNS
checks do not use the logged-in browser. These choices reduce unnecessary work;
they do not promise that live scraping cannot trigger a website restriction.

## Reproduce

From the extracted release folder:

```powershell
.\Run_Scraper.bat --check
py -3 -m unittest test_scraper test_browser test_company_match test_link_scraper test_live_regressions test_maps_validation test_maps_matching test_maps_browser test_public_company_integration test_company_public_data test_company_search_results test_company_search_dom test_business_email test_email_browser test_email_integration test_email_patterns test_strict_email test_strict_company_maps test_strict_workbook test_launcher test_app_smoke test_email_domain_check test_maps_office_scope test_benchmark test_feature_integration -q
py -3 benchmark.py --input "path/to/saved_results.xlsx" --output-dir "Output/Benchmarks" --sample-size 150
```

The test suite requires Chrome and briefly constructs the desktop window. Tests
and the offline benchmark do not log in or collect additional real profiles.
See README_EMAIL.md, README_MAPS.md and README_BENCHMARK.md for feature details.
