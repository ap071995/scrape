# Scraper 42.6.0

This release adds three quality controls without increasing live scraping speed:

- **Email Details** separates published addresses, generated candidates and
  existing addresses of unknown origin. Optional **Check email domains (DNS/MX)**
  is off by default and never verifies an individual mailbox.
- **Maps Evidence** distinguishes headquarters, local offices and unclassified
  company locations from the prospect's location. Contacts without sufficient
  office evidence are held for review.
- **Build Offline Benchmark** samples saved results (150 by default), creates a
  review workbook and reports field coverage separately from reviewed accuracy.
  It makes no website requests and reports a shortfall if fewer records exist.

Start with **Run_Scraper.bat**. Keep all release modules together. See
[README_EMAIL.md](README_EMAIL.md), [README_MAPS.md](README_MAPS.md), and
[README_BENCHMARK.md](README_BENCHMARK.md) for usage. Validation and limitations
are in [RELEASE_42_6_TEST_REPORT.md](RELEASE_42_6_TEST_REPORT.md).

## Previous v42.5.1 maintenance fixes

This maintenance release fixes issues reproduced by 86 new strict tests.
All **354 tests passed**, including browser extraction, Excel persistence,
Windows launcher checks, and real desktop startup/mode switching/shutdown.
See [STRICT_TEST_REPORT.md](STRICT_TEST_REPORT.md) for coverage and live results.

- Partial Excel updates preserve omitted fields; duplicate profile corrections
  and invalid evidence-sheet layouts are rejected before changing lead rows.
- Email ownership checks reject neighbouring staff contacts, delegated contacts
  and masked fragments. Official evidence takes priority; published addresses
  found during pattern searches take priority over generated candidates.
- Revenue parsing rejects forecasts, budgets, partial-year amounts and facts
  belonging to another company. Ambiguous classification-code ranges are rejected.
- Maps rejects conflicting company websites, named subsidiaries sharing a parent
  domain, invalid address sources, and phone links for a different number.
- The launcher preserves application exit codes and explains broken Python setups.

Business Email ID enrichment searches for a published person-specific work
address first, then uses a supported company email pattern as a fallback.
Published and generated results have distinct unverified statuses and a separate
Email Evidence sheet. Existing rows can be enriched in place; manual contacts
are preserved when the repaired person, employer and domain still agree.
See [README_EMAIL.md](README_EMAIL.md) for setup, retry behavior, tests and limits.

Extract all release files together and restart the application from that folder:

Double-click **Run_Scraper.bat**, or use the terminal command below.

```powershell
python Scraper.py
```

No new Python dependencies are required. Keep `business_email.py` alongside
`Scraper.py` and the other runtime modules. Select the same input and output
workbook to retry missing fields; close the workbook in Excel while scraping.

Current release checks and their limitations are in STRICT_TEST_REPORT.md.
See README_EMAIL.md for the source limitations and validation runner.

## Earlier features retained

Google Maps extraction, company/location matching and existing-row retries are
updated. See [README_MAPS.md](README_MAPS.md) for the fixes and validation runner.

Revenue, SIC and NAICS enrichment is now integrated in both modes. See
[README_PUBLIC_COMPANY.md](README_PUBLIC_COMPANY.md) for setup, source evidence,
validation and coverage limits. Version 42.4.1 also reads visible Google search
snippets and company fact cards when linked source pages are unavailable.
The sections below are a historical record of the **42.3.5** LinkedIn extraction
fixes and five-profile validation. Their counts and missing-field observations
describe that release, not the current release's business-email coverage.

## Historical v42.3.5 result

The five profiles in the supplied workbook were processed through the actual
Link Scraper with the existing authenticated Chrome session. The final run saved
5 of 5 rows with zero extraction failures. The XLSX was reopened and its requested
role, company and location fields were compared with retained profile/About data.
All 11 requested fields are populated in all five rows. No profile rows were added
or lost during repair.

Zach Loch now has Chief Information Officer, East Carolina University,
https://ecu.edu/, Higher Education, and 5,001-10,000 employees.

### Defects corrected in v42.3.5

- Current-role links sometimes contain the whole job card. The parser now finds
  the enclosing employer name/logo instead of treating Full-time or job text as
  the company. Nested titles and the first current role retain their DOM order.
- University employers use /school/ links; both school and company About pages
  are supported. Missing H1 headings and regional LinkedIn URLs are handled.
- About fields are read by label, and company-size ranges remain intact.
  Company names are verified with conservative fuzzy matching (88% threshold),
  legal suffix normalization, and explicit rebrand aliases. Similar subsidiaries
  are rejected. Unscoped education/recommendation links cannot invent an employer.
- Profile-header location supplies missing Experience geography. Company address
  extraction uses the Primary office only, excludes descriptions/adjacent offices,
  and separates supported US/Swiss postal codes. Telephone links supply clean digits.
- A corporate website match no longer overrides a wrong Maps city. Existing
  addresses/phones cannot be combined with another office's postcode/source links.
  Explicit foreign dialing codes on US/Canadian places are rejected and logged.
- Retrying incomplete rows replaces stale company/contact enrichment even when
  the employer name stayed the same. Campaign/admin data and original source-file
  provenance remain. Successful rows are immediately checkpointed with atomic saves.

### Historical v42.3.5 validation

All 57 regression tests passed; Python compilation and the saved-workbook audit
also passed. The original workbook was backed up as Output/Link_Scraper_Results.before_fix_*.xlsx
before installing the verified result at Output/Link_Scraper_Results.xlsx.

The regression suite includes browser fixtures reproducing the supplied screenshots
and the actual nested Zach page layout, plus Excel append/repair, duplicate handling,
locked-file recovery, fuzzy matching, location and Maps conflict tests.

Run from this project directory:

```powershell
python -m unittest test_scraper test_company_match test_link_scraper test_live_regressions test_maps_validation -v
python -m py_compile Scraper.py
python validate_saved_output.py
```

The following archive paths and historical helpers are not included in the
current release ZIP. Use the checks in README_EMAIL.md and README_MAPS.md for
the current release.

Evidence and reports are in Output/verification: live_evidence.json, rendered
profile/company HTML captures, live_run.log, regression_tests.log,
output_records.json and validation_report.json. run_live_validation.py is a bounded
live verification helper for the profiles in the supplied output workbook. It
operates on Link_Scraper_Results_verified.xlsx; complete rows are skipped as in the
app, so use a copy of the original incomplete workbook for a fresh repair run.

### Historical v42.3.5 usage notes

Restart the application using this project's Scraper.py. Keep the output workbook
closed in Excel while scraping. In Link Scraper, select the input file, output file,
and desired record limit. Reusing an output file repairs incomplete rows and skips
complete duplicates. Login/verification barriers stop the batch for manual action.

### Historical v42.3.5 data boundaries

This verification covers the five supplied output rows and FAST mode. It does not
claim all 98,084 input links or SMART/LLM enrichment were live-verified. The app uses
what pages display, including tenure strings, rather than recalculating them.
Display names have no guaranteed structured name components: the existing rule
uses the first token as First Name and the remainder as Last Name; multipart names
may need manual review. City/State/Country describe the prospect, while Address
can be the company's primary headquarters in a different location.

Unavailable or conflicting optional data remains '-' (including some phones,
3M's postal code, email, revenue and classification codes). These fields are not
invented to make the workbook appear complete. A future website layout change or
an unavailable/login-restricted page can still require another selector update.
