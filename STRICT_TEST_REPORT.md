# Scraper v42.5.1 — strict test report

Validated on **8 October 2026**, Windows, Python 3.14, installed Chrome and Playwright.

## Result

**354 tests passed, zero failures, zero skips** in the final combined run (57.222 seconds).
This includes **86 new tests** added for this maintenance release. Runtime and
validation-runner compilation passed. `Run_Scraper.bat --check` reported v42.5.1 ready.

| New tests | Count | Main checks |
| --- | ---: | --- |
| `test_strict_email.py` | 34 | Contact ownership, uppercase/inline staff names, masked fragments, assistant contacts, malformed inputs, official-source priority, conflicts, retries, published addresses found during pattern searches |
| `test_strict_company_maps.py` | 20 | Parent/subsidiary identity, annual versus forecast/partial-year revenue, classification ranges, invalid URLs, conflicting Maps websites, correct office and phone-source association |
| `test_strict_workbook.py` | 17 | Partial updates, duplicate URLs, invalid row numbers, evidence schemas, evidence ownership, missing profile markers, save failure/retry, phone text, user sheets |
| `test_launcher.py` | 14 | Windows special-character paths, working directory, interpreter selection, missing dependencies/files, broken environments, exit codes, noninteractive requirements check |
| `test_app_smoke.py` | 1 | Real Tk window, mode controls, Playwright worker initialization, clean shutdown |
| **Total new** | **86** | Several tests include multiple adversarial subcases |

The existing 268 tests were retained. Browser tests execute extraction JavaScript
in real Chrome against controlled HTML fixtures; they do not require LinkedIn login.
The desktop smoke test constructs the real application and starts its worker.
Only its optional local-LLM availability check is disabled for that test.

## Reproduced defects and fixes

1. **Excel updates could clear fields omitted by a caller.** Updates now merge
   supplied fields into the existing row. Explicit empty values still clear a
   field. Invalid row numbers and duplicate profile corrections are rejected.
   Correcting a profile URL also updates the duplicate index across restart.
2. **Bad optional evidence could interrupt or partially apply a row update.**
   Evidence is filtered and sheet headers are checked before the lead changes.
   Email evidence must match the saved person, company, profile and address.
   Unknown profile markers do not become duplicate keys. Phone control characters
   are removed while preserving leading `+` and Excel text type.
3. **Email association accepted some neighbouring or delegated contacts.**
   Ownership checks now handle uppercase and inline staff headings, masked
   addresses and assistant/delegate wording. A search title cannot override
   rejected ownership in its result body.
4. **A weak email snippet could win before official evidence was checked.**
   Configured/cached official sources are considered first, later official
   searches retain a fetch allowance, and generated candidates are retried.
   Pages and snippets fetched during pattern searches are also checked for
   published person-specific addresses. Equal-ranked conflicts remain ambiguous.
5. **Company facts could be borrowed from a subsidiary or unrelated named entity.**
   Explicit company identity now overrides a matching parent-page title. Revenue
   parsing rejects budgets, guidance, forecasts and partial-year figures while
   allowing a later actual annual section. Code ranges and alternative codes
   are not truncated into a single SIC/NAICS value.
6. **Maps matching could accept a conflicting website or named subsidiary.**
   A conflicting website now rejects the result; a parent domain alone cannot
   validate a named subsidiary. Generic office suffixes and regional subdomains
   remain accepted. Invalid address links and phone links for another number
   cannot enter the saved fields.
7. **Launcher failures lost useful diagnostics.** The batch launcher now preserves
   the application's nonzero exit code, distinguishes a broken Python interpreter
   from missing requirements, and rejects nonzero preflight failures.

Failing cases were reproduced before repairs. Relevant retained logs are under
`Output/strict_validation`; the earlier company/Maps failures are recorded in
`Output/strict_company_maps_before.log`. The final combined result is in
`Output/strict_validation/full_tests.log`.

## Fresh live email and Excel check

The earlier `Output/Link_Scraper_Results.xlsx` and Maps-verified workbook were
no longer present. The live check therefore used the existing historical
five-prospect workbook `Output/verification/first_pass.xlsx`, supplied explicitly
to the runner. It wrote a separate output and did not change that source.

| Check | Result |
| --- | --- |
| Prospects attempted | 5 |
| Published addresses found | 1 |
| Supported pattern candidates | 2, marked `Generated (unverified)` |
| Unresolved | 2, marked `Not found` |
| Saved lead rows | 5; zero added or lost |
| Saved email evidence rows | 3 |
| Unrelated cells and hyperlinks | Preserved |
| Source workbook | Hash unchanged |

Zach Loch's published address was read from ECU's official IT leadership page.
The 3M and UBS candidates use published company-pattern evidence; this does not
verify those individuals' mailboxes. Sheila Reddic and Robert Parish remained
unresolved. Google required verification and was stopped for that run. Accessible
sources and Bing fallback continued; four source errors/unavailable results were
recorded. No verification challenge was bypassed.

Local artifacts (validation data is not included in the distributable ZIP):

- `Output/strict_validation/Link_Scraper_Results_checked.xlsx`
- `Output/strict_validation/email_verification/validation_report.json`
- `Output/strict_validation/email_verification/source_documents.json`
- `Output/strict_validation/email_verification/events.json`

**Scope:** This run checked email enrichment against live sources and verified
Excel preservation. It did not freshly verify the historical workbook's other
fields against LinkedIn, Google Maps or financial sources. Those paths were
covered by regression and browser-fixture tests. Passing tests does not guarantee
that every website layout or every one of the 98,084 input links will yield all
fields. Missing/ambiguous facts remain missing rather than being invented.

## Run the application

Use `Run_Scraper.bat` beside the v42.5.1 runtime files, or run `python Scraper.py`.
When using the ZIP, extract all its files into one folder before launching.
The old v42.5.0 ZIP and extracted folder are retained as the previous release.

## Repeat validation

Run from the release folder. Chrome, Tkinter and the application's existing Python
dependencies must be installed. The full suite includes a brief desktop window.

```powershell
python -m unittest test_scraper test_browser test_company_match test_link_scraper test_live_regressions test_maps_validation test_maps_matching test_maps_browser test_public_company_integration test_company_public_data test_company_search_results test_company_search_dom test_business_email test_email_browser test_email_integration test_email_patterns test_strict_email test_strict_company_maps test_strict_workbook test_launcher test_app_smoke -q
python -m py_compile Scraper.py business_email.py company_public_data.py company_search_dom.py google_maps_dom.py run_email_validation.py run_maps_validation.py run_company_enrichment.py
.\Run_Scraper.bat --check
```

For a separate live sample, replace the input path with an existing Leads workbook:

```powershell
python run_email_validation.py --limit 5 --input "path/to/existing_leads.xlsx" --output "Output/strict_validation/Link_Scraper_Results_checked.xlsx"
```

Run live validation after browser tests finish. It preserves the input, updates
email fields on up to five distinct prospects, and checks the reopened output.
