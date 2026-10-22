# Lead Extractor 42.9.0 — validation report

Validated locally on Windows with Python 3.14 and sandboxed Chrome, 19 October 2026.

## Result

**542 automated tests passed in 86.834 seconds.** Launcher preflight and Python
compilation also passed. The separate runtime validation completed successfully.
No live LinkedIn profiles were visited for these checks. User result workbooks
were not used as test destinations.

## Findings and changes

- No hard-coded 60-entry import cap was present. The previous Excel reader
  traversed the entire declared worksheet dimensions, including formatted empty
  cells. The replacement reads actual stored cells across sheets and columns,
  preserving plain links, Excel hyperlinks and HYPERLINK formulas.
- File selection now counts links in a background thread. The UI distinguishes
  the imported count from the saved-lead target and offers **All imported links**.
  The precise cause of the reported stop at 60 is not established without its
  input file/activity log; access stops and requested targets still apply.
- Email enrichment previously chained multiple queries/page fetches without a
  total budget. It now shares a 15-second budget across browser/API operations
  and uses remaining time for optional DNS/MX checks. Timed-out lookups remain
  unverified and retryable. Published evidence and generated candidates retain
  separate status/evidence handling.
- Added stage progress for import, browser, resolution, profile, company, Maps,
  public company facts, email, Excel, Sales search, SMART validation and benchmark.
  Request stages use moving indicators because their percentage is unknown.
- Added a paged worksheet editor, custom columns, manual rows and row editing.
  Custom data survives later scraper updates, workbook reopening and journal
  recovery. Recovery now handles manual rows whose profile URL is later edited.
- Retained the existing LinkedIn pacing, authentication/error stops, sandboxed
  browser launch settings, isolated enrichment workers and windowless launcher.

## Evidence

| Check | Result |
|---|---|
| CSV import, one column | 98,084 links in 2.103 seconds |
| XLSX import, one column | 98,084 links in 4.439 seconds |
| Sparse Excel dimensions | 1,004 links across columns/sheets; formatting at XFD1048576 did not cause empty-cell traversal |
| Offline Link Scraper loop | 75 of 75 fixture profiles written and reopened successfully |
| Extended output workbook | 77 data rows, 44 columns; two manual rows and two custom columns retained |
| Slow local browser server | Default 15-second email budget returned in 15.017 seconds with Time limit reached (unverified) |
| Recovery | Columns, manual rows and changed profile URLs replay without duplicate rows |
| Desktop UI | Real widgets, import preview, progress lifecycle, worksheet controls, All option and worker shutdown passed |
| Regression coverage | Both scraping modes, mixed links, email evidence/patterns, Maps, company matching, cache, benchmark, sandbox and launcher |

The runtime measurements exclude browser startup from email lookup time. They
use synthetic data and a deliberately delayed local server, so they establish
capacity/deadline behavior rather than live source accuracy or email discovery
rates. Remote page behavior may affect yield; browser cleanup/local parsing may
add a small amount beyond the search budget. Fifteen seconds is not the total
time for all enrichment of a lead.

## Reproduce

```text
py -3 -m unittest discover -p test_*.py -q
Run_Scraper.bat --check
py -3 validate_v42_9.py
```

The runtime validator writes fixtures under `Output/v42.9.0_validation` and
refuses to overwrite its existing output workbook. Use a fresh extracted release
folder when rerunning it.

Local evidence: `Output/v42.9.0_validation/full_tests.log`, `ui_tests.log`,
`validation_results.json`, `runtime_validation.log`, and `worksheet_fixture.xlsx`.
