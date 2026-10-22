# Business email enrichment — v42.6.0

The scraper searches for a published work email for the named prospect first.
It generates a candidate only when a company email pattern has supporting public
evidence. Both FAST and SMART modes use this workflow in Link Scraper and
Sales Navigator.

## Run

Extract all release files into the same folder. Keep these runtime files together:

- `Scraper.py`
- `business_email.py`
- `company_public_data.py`
- `company_search_dom.py`
- `google_maps_dom.py`
- `company_sources.json`
- `email_domain_check.py`
- `maps_office_scope.py`
- `benchmark.py`

Use the existing Python environment; this feature adds no Python dependencies.
Double-click **Run_Scraper.bat** in the extracted folder to open the application.
Keep its console open while the application is running; startup errors appear
there. The launcher uses a local `.venv`, `env`, or `venv` when present, otherwise
the installed Python 3. To check requirements without opening the app, run
`Run_Scraper.bat --check`.

You can also start the application from a terminal:

```powershell
python Scraper.py
```

Close the output workbook in Excel before running. In Link Scraper, select the
same input and existing output workbook to retry missing business emails.
Already saved profiles are updated in place without adding duplicate rows.
Complete LinkedIn profiles do not need another profile visit for email lookup.
An attempted enrichment counts toward the selected record limit even if no
acceptable email is found.

## Search and fallback order

1. Search the prospect's full name, company name and email, then search within
   the company's website domain. A first name, last name, company name and
   company website are required.
2. Read relevant search snippets and source pages, including `mailto:` links.
   Match the address to the named person's contact block. Accept supported
   obfuscation such as `jane.doe [at] example [dot] com`.
3. If no published address is found, search the company name/domain with
   `email format` and `email pattern`. Company-name checks use the existing
   conservative fuzzy matcher. A pattern can also come from multiple named
   employees on an official company page when they support one common format.
4. Generate one candidate from the supported format, such as
   `{first}.{last}` or `{last}_{f}`, where `{f}` is the first-name initial.
   Without sufficient pattern evidence, leave the email blank (`-`).

Cached company pages and configured official staff directories are also checked
for the named person's published address before generation. `company_sources.json`
supports an optional `email_urls` list under a company's existing domain entry;
these URLs must belong to that official domain. The ECU entry includes its public
IT leadership directory. Existing `urls` entries can supply company-pattern
evidence when search misses them. All values are fetched from pages at runtime;
the configuration contains no email addresses or hardcoded patterns.

Known official staff directories are checked before person searches. Official
source pages take precedence over third-party pages and search-only evidence.
A weaker snippet does not stop later searches for an official source. A visible
snippet can supply an address when its destination cannot
be fetched; its evidence identifies it as a search result and retains the
original destination URL. Google/Bing searches share the existing company web
service and its optional Brave API configuration.

The parser excludes known personal email providers, generic inboxes, masked
addresses, examples explicitly described as generated/guessed, and addresses
belonging to a neighbouring staff entry. Company-domain checks prevent unrelated
third-party addresses from becoming a person's work email. A named contact
published on the official site can establish a different corporate email domain;
the scraper does not invent a replacement domain from the website's spelling.

Conflicting equally ranked published addresses remain unresolved. Conflicting
patterns require a supported preferred format or remain unresolved. The code
does not try every possible address until one appears to work.

## Excel fields and evidence

The existing 42-column **Leads** sheet is retained.

### Separate addresses and optional domain checks

The **Email Details** sheet now keeps **Published Email ID**, **Generated Email
Candidate**, and **Existing Email (origin unknown)** in separate columns. A
generated candidate does not become published because its domain has MX records.
When a published address replaces a generated candidate, the current detail row
updates; older source evidence remains in Email Evidence. The Leads sheet keeps
the selected address and its existing status for compatibility.

Enable **Check email domains (DNS/MX)** under Run Settings to check the selected
address's domain. It is **off by default**. Checks use Google Public DNS over
HTTPS and send only the domain, never the address's local part or the prospect's
name. They require no browser/account access and make no SMTP connections.

Results distinguish MX records, an explicit null MX, a nonexistent domain,
implicit A/AAAA fallback, no route, timeout and inconclusive failure. A sole null
MX means the domain declares that it does not accept mail; when MX is absent,
SMTP defines an address-record fallback. Neither result verifies a mailbox.
See [RFC 7505](https://www.rfc-editor.org/rfc/rfc7505) and
[RFC 5321 section 5](https://www.rfc-editor.org/rfc/rfc5321#section-5).
The resolver format follows [Google's DNS JSON API](https://developers.google.com/speed/public-dns/docs/doh/json).

The checker reuses domain results with bounded TTL caching (up to 1,024 domains)
and limits each uncached domain to MX plus, if needed, A/AAAA queries. Lookup
failures preserve the email. **Mailbox Verification always stays Not verified**;
Bounce Email Status is unaffected. A previous saved DNS check can be retained
when checks are disabled and the selected address is unchanged; its **Checked
At** timestamp records when it actually ran.

Published/generated separation is also available to the
[offline benchmark](README_BENCHMARK.md), independently of the DNS option.

| Field/status | Meaning |
| --- | --- |
| Business Email ID | Accepted published address or evidence-based generated candidate |
| `Found online (unverified)` | Address appears in an accepted public source for the named person |
| `Generated (unverified)` | Candidate created from a supported company format; not found as that person's published address |
| `Not found` | No acceptable published address or usable company pattern was available |
| `Ambiguous (unverified)` | Conflicting published addresses prevent selecting one |
| Notes | Lookup method/status; generated results also record the chosen pattern |

**Neither a published address nor a generated candidate is proof of mailbox
ownership, current employment or deliverability.** The scraper does not send
messages, probe SMTP servers, or mark an address as verified. It preserves the
existing **Bounce Email Status** when the address is unchanged; a changed address
resets that field to **Not evaluated**.

The **Email Evidence** sheet records:

- Person and Prospect LinkedIn Link
- Company Name and Business Email ID
- Method and Pattern
- Source URL and Supporting Text
- Retrieved At

Evidence remains associated with the person, company and address. Repeated
identical evidence updates its row; evidence for a changed address remains
available for review. Search-only evidence links to the search page, while a
fetched source links to the source itself.

## Existing rows and later runs

- Valid existing email addresses and their manual statuses are retained.
- Generated addresses are searched again on a later run and can be upgraded to
  a published address. An unsuccessful retry retains the generated candidate
  and its unverified status.
- Repairing other LinkedIn fields preserves the existing email/status when the
  normalized person name, company identity and website domain still agree.
- A changed employer or corrected person identity does not inherit the previous
  person's/employer's email. Administrative fields and original source-file
  provenance remain preserved.
- Lookup failures preserve existing row data. Missing fields can be retried by
  selecting the same input/output again; search and pattern caches reset per run.

## Checks and limits

See [RELEASE_42_6_TEST_REPORT.md](RELEASE_42_6_TEST_REPORT.md) for current tests.
The earlier [STRICT_TEST_REPORT.md](STRICT_TEST_REPORT.md) records the v42.5.1 live sample,
saved-file checks, and exact limits. These results do not establish mailbox
deliverability or coverage for the full input list.

The new regressions cover person ownership, obfuscated and masked addresses,
company patterns, conflicting evidence, saved-row retries, manual-contact
preservation, both scraping modes, and extraction through saving/reopening Excel.
Run from the project folder:

```powershell
python -m unittest test_business_email test_email_browser test_email_integration test_email_patterns test_strict_email -v
python -m py_compile Scraper.py business_email.py company_public_data.py company_search_dom.py google_maps_dom.py
```

The browser tests use local fixtures and Chrome through the existing Playwright
setup. Run browser tests and live checks sequentially. Fixture success does not
establish live coverage for every prospect. Source availability, supported page
layouts, ambiguous names and company email-format exceptions can leave rows
unresolved. Verification challenges stop the affected search provider for the
run; they are not bypassed.

For a bounded live check, run `python run_email_validation.py --limit 5`.
It copies the saved Maps-verified workbook (or the original if unavailable) to
`Output/Link_Scraper_Results_email_verified.xlsx`, updates up to five distinct
prospects, and reopens the result. It checks row counts, email evidence, unrelated
cells/sheets, and the original file hash. Use `--input` and `--output` to select
different paths. Captured sources and the report are in `Output/email_verification`.
When selecting a different output folder, evidence and reports are written into
an `email_verification` subfolder beside that output workbook. If the default
input workbook has been moved, supply its current path with `--input`.
