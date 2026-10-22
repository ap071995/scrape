# Scraper v42.7.0 — 15 October 2026

## Verified result

The combined regression suite passed **498 tests in 61.872 seconds**, with no
failures or skips. This includes the previous 474 tests and 24 new tests covering
durable caching, journal replay, indexed resume, company identity, office scope,
concurrent browser ownership, partial failures, and headless controls. The real
desktop smoke test also checks the new UI defaults. Three targeted checks passed
after the final change to stop immediately on headless sign-in/verification.

The application and support modules compile. The Windows launcher preflight
reports **Universal Lead Extractor v42.7.0 is ready**.

Logs:

- `Output/v42.7_validation/final_tests.log`
- `Output/v42.7_validation/final_auth_check.log`
- `Output/v42.7_validation/session_ui_tests.log`

## UBS defect and live output

The saved profile uses London, England, United Kingdom. Earlier logs show Maps
queries targeting Zurich from the saved company address, then rejecting results
against the prospect's UK geography. Other historical logs found London but
held its phone because the existing address belonged to Zurich.

The new local-office option searches with the prospect's available geography,
validates the opened company and location, and replaces the complete contact
group only after validation. The previous company contact group is retained in
Previous Company Location. Maps Evidence still reports the office type and does
not claim the prospect works there.

A fresh headless Maps run recovered and saved:

| Field | Saved value |
| --- | --- |
| Company listing | UBS AG |
| Address | 5 Broadgate, London EC2M 2QS, United Kingdom |
| Postal code | EC2M 2QS |
| Phone | +442075678000 |
| Previous address preserved | Postbox, Zurich, 8001, CH |

The workbook was reopened and its values checked against the supplied screenshot.
The run took **6.62 seconds** and made **zero LinkedIn visits**. Evidence and the
captured Maps screenshot are in `Output/v42.7_validation/live_maps/`; the separate
workbook is `UBS_Maps_validation.xlsx`. Existing user result files were not edited.

## Performance validation

An offline workload appended 100 rows, flushed the final workbook, reopened it,
and confirmed all 100 rows remained. Both modes journal each row before writing.

| Persistence mode | XLSX rewrites | Measured time |
| --- | ---: | ---: |
| Save each row | 100 | 9.308 s |
| Save every ten rows | 10 | 1.111 s |

A synthetic five-record workload with two independent 150 ms stages took
**1.508 s sequentially** and **0.763 s concurrently**. This verifies scheduling;
it is not a forecast for live websites. Results are in
`Output/v42.7_validation/runtime_benchmark.json`.

Recovery tests cover a crash before XLSX checkpoint, a crash after atomic XLSX
replacement but before journal deletion, a torn final journal line, locked Excel
files, repeated replay, and evidence-sheet preservation. Journal data is retained
when a save fails. Company and office caches are tested for expiry and identity
isolation. Two real headless browser sessions were also started in separate
threads; their LinkedIn requests were blocked and they shared no cookies.

## Delivered changes

- Indexed URL lookup and resume distinguish missing profile data from missing
  company/contact fields. Recent unchanged enrichment attempts do not use up the
  run's saved/updated-row target.
- Captured profile evidence and company URLs are retained for reuse and repairs.
- SQLite caches persist company About details, public facts, Maps results and
  supported email patterns, with expiry and shorter incomplete-result lifetime.
- Maps and company/email work can run in two independent sessions. Workbook
  writes remain in the main worker; no Playwright objects cross threads.
- Performance counters and stage timings are written beside each output workbook.
- The UI adds headless, concurrent enrichment, local office preference and
  explicit saved-profile refresh options.

The LinkedIn collector remains sequential and its existing page wait/scroll
constants were not reduced. Faster enrichment can shorten elapsed time between
profile visits. Headless and concurrency do not guarantee freedom from account
restrictions. The validation did not run a bulk LinkedIn collection or establish
accuracy on 100–200 live profiles. The live verification covers the supplied UBS
Maps example; the other changes were validated with offline and browser fixtures.

Usage and reproducible commands are in `README_PERFORMANCE.md`.
