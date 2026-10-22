# Scraper v42.8.0 validation

## Result

Mixed normal-profile and Sales Navigator lead URLs now enter the same file-based
workflow. Normal links use the existing scraper; Sales Navigator links are
resolved from an explicit public-profile action, cached, then processed by that
scraper. No public identifier is guessed from the Sales Navigator ID.

The Leads schema remains 42 columns. A new Link Resolution companion sheet
records original and resolved URLs, resolution and processing statuses, errors
and timestamps. Its changes use the existing recovery journal.

## Live check of the supplied link

The supplied Sales Navigator lead ID
`ACwAADv0aHgBhjpTrvfVOtoOH4d-brIoiRfKxcM` opened a lead page for **KM Nirupama**.
The profile action resolved to:

`https://www.linkedin.com/in/km-nirupama-708a46240`

This is a different person from the separate Greg Mirigian normal-profile URL.
The successful live check took 6.77 seconds. A second check returned the cached
mapping in 0.44 seconds without launching a browser. The live check covered URL
resolution only, not a new bulk enrichment or contact-data accuracy benchmark.

The first live attempt exposed a missing action-menu selector. It was corrected
to support the current **Open actions overflow menu**, then verified on the same
lead. No public profile page had to be opened to read its explicit URL.

Local evidence (excluded from the ZIP):

- `Output/v42.8_validation/live_check.log`
- `Output/v42.8_validation/cached_check.log`
- `Output/v42.8_validation/live_resolution/page.png`
- `Output/v42.8_validation/live_resolution/page.html`

## Automated coverage

The initial complete regression suite passed 514 tests. The final focused
mixed-link suite passed 18 tests, including two added after that full run.
The final regression pass passed **515 non-desktop tests in 73.180 seconds**,
with no failures, errors or skips. Its log is
`Output/v42.8_validation/final_regressions.log`. The real desktop smoke test
passed earlier, including window creation, mode changes, option defaults and
worker shutdown: **516 distinct tests passed across these runs**.

The new coverage includes:

- strict host/path validation and regional/tracking URL normalization;
- Excel hyperlinks, Markdown links and spreadsheet hyperlink formula text;
- explicit links, current/legacy menus, delayed menu links and new-tab actions;
- rejection of recommendations, ambiguous links, invalid hosts and access screens;
- persistent mapping reuse, explicit refresh and retryable failed resolution;
- crash recovery for link audit entries and protection of an existing custom sheet;
- mixed input through the real browser parser and Excel writer: one profile row
  for a normal link plus two Sales Navigator variants resolving to that person;
- reopened workbook checks for name/company fields, duplicate statuses,
  unresolved leads and invalid inputs.

Browser fixtures use intercepted local HTML and do not access LinkedIn. The
Windows launcher preflight and Python compilation checks also passed.

## Interface

The new dark navy/teal interface groups settings into Setup, Options and Tools,
makes mixed-link file input the default, keeps run controls outside the scroll
area, and improves progress/preview/log contrast. Run controls are disabled when
idle; the source mode cannot be changed during processing.

Manual desktop visual inspection was stopped with Escape. No further desktop
interaction was performed. Automated UI checks passed, but complete visual
verification of every tab and screen size was not completed.

## Scope and limitations

LinkedIn is still sequential and uses its existing waits. Browser concurrency
remains limited to the existing offsite enrichment workers. The first resolution
adds a Sales Navigator visit, and cached mappings avoid it on later runs.
Account-restriction risk cannot be eliminated.

Unresolved links are recorded rather than converted by guessing. The resolver
currently recognizes English action labels. Use a normal profile link if the
account cannot access the Sales Navigator lead or its profile action is absent.
Use Refresh saved LinkedIn profiles to recheck a stale mapping. Links beyond the
saved-lead target or after Stop are not attempted. Existing result workbooks were
not modified during validation; release packages exclude browser sessions,
mapping databases, live HTML, screenshots and user prospect data.
