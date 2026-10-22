# Revenue, SIC and NAICS enrichment — v42.6.0

This release rejects budget/forecast and partial-year revenue, facts about a
different named company, ambiguous classification-code ranges, and malformed
source URLs. Strict regression cases are in `test_strict_company_maps.py`.
It also includes the [Google Maps fixes](README_MAPS.md) and
[business email enrichment](README_EMAIL.md).

## Run

Keep `Scraper.py`, `business_email.py`, `google_maps_dom.py`, `company_public_data.py`, `company_search_dom.py`, and
`company_sources.json`, `email_domain_check.py`, `maps_office_scope.py`, and
`benchmark.py` in the same folder. Run `python Scraper.py` from that folder
with the existing Python environment. No new Python
dependency is required. Enrichment runs automatically in FAST and SMART for Link
Scraper and Sales Navigator.

Select the existing output workbook and the same profile input to enrich its
missing company fields. Complete profile rows are updated in place without
visiting LinkedIn again. An attempted company enrichment counts toward the run's
record limit, even if a source is unavailable; check the unresolved-field log.

## Where the data comes from

- The worker searches company name + annual revenue, SIC code and NAICS code.
  Browser search uses Google, then Bing. A verification challenge disables that
  provider for the remainder of the run; it is not bypassed.
- An optional `BRAVE_SEARCH_API_KEY` environment variable enables the documented
  Brave Search API before browser search. No key is bundled. The API path is
  tested with mocked responses; live validation uses browser search.
- `company_sources.json` contains only public source URLs, keyed by official
  website and company name. It seeds the five validation companies and can be
  extended. Values are fetched and parsed at runtime, never hardcoded in this file.
- Source pages must match the company name or its verified official domain.
  The scraper follows a bounded number of relevant annual-report HTML links.
- Visible search results can now fill Revenue, SIC and NAICS even when the
  linked page cannot be read. This fixes v42.4.0, which used snippets only to find
  links and discarded the numbers visible on Google.
- Google organic snippets, featured snippets and labelled company fact cards
  are supported. Bing and optional Brave API snippets use the same parser.
  Cards stay separate and must identify the company through their visible title;
  the search query alone is never company evidence. Company names use fuzzy
  matching with legal-suffix normalization.
- Source-page facts take priority over search-only facts. Search-only values
  link to the search page and are labelled `search result; source page not
  verified` in Company Evidence. The excerpt also retains the result's original
  destination URL. Unsupported or ambiguous cards remain unresolved.
- Each company is cached per run, including unsuccessful lookups. A bounded
  document cache prevents large batches from retaining every downloaded page.

## Excel output

The existing 42-column Leads sheet is retained:

| Field | Meaning |
| --- | --- |
| Revenue | Source amount and currency symbol/unit; no currency conversion |
| Revenue Link | Source page, or the search page for a search-only value |
| Notes | Revenue period and reported/estimated status |
| SIC Code | Four-digit SIC; a labelled eight-digit commercial extension can supply its four-digit SIC parent |
| NAICS Code | Published six-digit industry code; a broad sector such as 531 is not padded or guessed |
| SIC/NAICS Code Link | An accepted classification source; use Company Evidence for each code's own source |

The **Company Evidence** sheet stores company, website, field, value, URL, period,
basis, a short supporting excerpt, and retrieval time. Dates refer to the source's
stated reporting period, not automatically the current year. Third-party amounts
are labelled estimates/listings; they are not audited financial statements.

The parser rejects budgets, forecasts, quarterly figures, revenue per employee,
unrelated tax amounts, subsidiary/competitor snippets and conflicting equal-ranked
facts. SMART mode cannot invent these financial/classification fields or replace
their source links with generic classification-definition pages.
Ads, unanswered related questions and AI Overview blocks are excluded. Currency
must be present in the result; both `$12 million` and `48.61 billion USD (2024)`
are supported. Four-digit SIC codes retain leading zeros.

Existing missing-value markers such as `NA` and `Not found`, and malformed
revenue/code values, can be repaired from a matched source. Valid conflicting
existing values are preserved. Browser failure logs include the underlying error.

## Validation and limits

Version 42.4.1 passed **138 regression tests** and Python compilation checks.
New regression cases exercise search-result parsing, fuzzy identity checks,
scoped browser DOM extraction, result/source priority and browser-to-Excel saving.
The end-to-end fixture confirms all three fields reach the saved workbook when
destination pages are unavailable. Google returned a traffic-verification page
during the live probe on 22 September 2026, so successful extraction from an
unblocked live Google response could not be verified in this environment.

### Previous v42.4.0 baseline

The final regression run passed **101 tests**, including browser DOM fixtures,
Excel append/update behavior, fuzzy company matching, source-value pairing,
annual/quarterly distinctions, and rejection of unrelated asset/subsidiary figures.
Python compilation also passed.

The live run on **21 September 2026** processed five companies and saved six of
the fifteen requested Revenue/SIC/NAICS values with supporting evidence:

| Company | Sourced fields | Remaining gaps / qualification |
| --- | --- | --- |
| East Carolina University | Revenue, NAICS | SIC unavailable; revenue period 2024 |
| 3M | Revenue, SIC, NAICS | Revenue period 2025; classification source report is dated 2022 |
| City of Glendale AZ | Revenue | Directory estimate with no stated year; SIC/NAICS unavailable |
| UBS | None | No supported accessible source supplied an accepted value |
| Village Green | None | No supported accessible source supplied an accepted value |

The saved test workbook retained all ten leads. The remaining five leads and all
profile fields were unchanged. Eight evidence rows comprise six new sample facts
and two existing facts for other leads; no empty evidence rows remained. The
original workbook's SHA-256 was unchanged. This is partial source coverage, not
a claim that every requested company field was filled.

### Running checks

Run the browser regression suite and the live runner sequentially. The test
fixtures clean up Playwright even if browser startup fails. Public enrichment
closes its failed tab after a navigation error so the next source starts fresh.

Run regression tests from the project directory:

```powershell
python -m unittest test_scraper test_company_match test_link_scraper test_live_regressions test_maps_validation test_public_company_integration test_company_public_data test_company_search_results test_company_search_dom -v
```

`run_company_enrichment.py` makes a separate workbook, enriches the first five
existing profiles, and checks that all other fields and rows are preserved. It
clears only the sample's public company facts in that disposable copy first, so
the test fetches fresh evidence even when the original already contains values.
The original workbook is checked by SHA-256 and is never written by this runner. It
stores logs, source captures and a report in `Output/company_verification`.

Browser search availability varies. During live testing Google intermittently
required verification, Bing returned irrelevant results, and some directories
returned errors. Sources that require login, payment or verification are skipped.
This version parses supported English HTML layouts, including annual tables with
explicit year columns and currency amounts; it does not parse PDF-only accounts or
every financial-report layout. Unresolved fields remain `-`. Private companies and
public bodies may have no usable published six-digit NAICS/SIC or annual figure.
Published classifications can also differ by establishment or source date.

The output is source-backed enrichment, not a guarantee of complete or current
coverage for all 98,084 input profiles. The evidence sheet makes estimates,
older reporting periods, and remaining gaps reviewable.
