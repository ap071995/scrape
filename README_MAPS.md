# Google Maps fixes — v42.6.0

## Company office versus prospect location

Each fresh or cached Maps candidate is classified before its address or phone
can be merged. **Maps Evidence** stores the company location separately from
the prospect's City/State/Country, along with its source, classification basis,
merge decision, reason and fields actually applied.

- Explicit headquarters/head-office labels identify **Headquarters**.
- Explicit branch/local/regional-office labels identify **Local office**.
- Generic corporate-office listings stay **Company location (unclassified)**.
- A registered address alone, conflicting labels, missing address/source or
  rejected company/location identity hold the contact for review.

A matching prospect city is search guidance, not evidence of employment at that
office. **Prospect Works Here Verified** remains **No**. Company headquarters can
be in a different country from the prospect when the saved company address and
candidate agree. Phone-only listings are held because they cannot be attached
to a particular office. Cached candidates pass the same validation again.

The actual collector now checks website conflicts even when the display name
matches exactly. Search pacing and live request limits remain unchanged.

## Earlier matching fixes retained

The strict validation pass adds checks for conflicting company websites, named
subsidiaries on shared parent domains, malformed source links, and `tel:` links
for a different phone number. Regional subdomains and generic corporate-office
names remain supported. See [STRICT_TEST_REPORT.md](STRICT_TEST_REPORT.md).

## Run the updated scraper

Extract the release files into one folder and restart `python Scraper.py`.
Keep `google_maps_dom.py` alongside the other Python modules. Select the same
input and existing output workbook to retry missing Maps fields. Complete
LinkedIn profiles are updated in place; they do not require another profile visit.

## Defects corrected

- Existing workbook rows previously retried only Revenue/SIC/NAICS. They now
  also retry missing Maps address, postcode, phone and source links.
- The merge rejected equivalent addresses such as `E 5th St` and
  `East 5th Street`, including ZIP versus ZIP+4. It now recognizes those formats.
- Known company office addresses guide searches even when the prospect lives in
  another city. City-only placeholders can receive a complete sourced address.
- Queries include state/country when no company address is available. Domestic
  addresses can omit the country; common city/state spellings and abbreviations
  are normalized. Explicit conflicting locations are still rejected.
- Office labels such as `Corporate Headquarters` and fuzzy company-name matches
  no longer fail the result selector's exact-token requirement. A website query
  can open a trading name for verification against its actual website.
- A shared city/university website cannot turn a department, bookstore or other
  marked subunit into the organization's general contact. The live Glendale
  water-department result exposed this distinction and is now rejected.
- Rejected or incomplete listings no longer end the lookup immediately. The
  scraper checks up to three candidates per query and retains valid partial data.
- Readiness checks and extraction use the same visible place panel. Hidden
  panels and neighbouring results cannot supply its address/phone. Phone data
  comes from `tel:` attributes even when the displayed label is only `Call`.
- Scrolling targets the place details instead of the longer results feed.
- Generic address headings and omitted suite numbers no longer block a matching
  business at the same street address. Different explicit units remain conflicts.
  An omitted suite is not claimed as independently verified.
- National and international formats of the same existing phone can receive a
  source link. A conflicting international number for a US/Canadian place is
  excluded, including when the country is omitted from its address.
- Cache keys distinguish company offices. Location failures include the observed
  place and expected location in the activity log.

## Validation

The full regression suite passed **194 tests**. Python compilation also passed.

The live run on **23 September 2026** completed five companies, saved and reopened
the output workbook, and passed its preservation checks. All ten original rows,
unrelated fields and the original workbook's SHA-256 were preserved.

| Company | Live result |
| --- | --- |
| East Carolina University | Phone `+12523286131` saved; existing address/ZIP matched despite formatting differences |
| 3M | City-only address upgraded to `2510 Conway Ave, St Paul, MN 55144, United States`; phone unresolved |
| City of Glendale AZ | Existing address/ZIP retained; water-department phone rejected as a different contact |
| UBS | Zurich phone `+41442341111` saved using the company office location; existing postal address retained |
| Village Green | Phone `+12488519600` verified at the matching Southfield street address; existing suite retained |

Three of five company phone fields were freshly verified. The St Paul 3M listing
displayed a `+44` number; another listing was in Maplewood. These were not combined
with the St Paul office. The report marks the existing UBS postal address and
Village Green suite as retained rather than independently verified by Maps.

`test_maps_matching.py`, `test_maps_validation.py`, `test_maps_browser.py`, and
`test_public_company_integration.py` exercise these cases, wrong-company and
wrong-office rejection, and browser extraction through saving/reopening Excel.

Run the complete regression suite:

```powershell
python -m unittest test_scraper test_browser test_company_match test_link_scraper test_live_regressions test_maps_validation test_maps_matching test_maps_browser test_public_company_integration test_company_public_data test_company_search_results test_company_search_dom -q
```

Run a bounded live Maps check separately:

```powershell
python run_maps_validation.py
```

The runner reads `Output/Link_Scraper_Results.xlsx`, checks five distinct
companies, and saves `Output/Link_Scraper_Results_maps_verified.xlsx`. It preserves
unresolved existing values and checks that unrelated cells and rows remain
unchanged. `Output/maps_verification/validation_report.json` distinguishes fresh
source agreement from retained values; source captures and logs are kept nearby.
The original workbook's hash is checked before and after. The runner overwrites
its disposable validation output on each run. Run browser tests and live checks
sequentially.

Maps listings can themselves contain missing or conflicting details. The code
keeps unresolved fields rather than attaching another office's contact details.
The live sample is a bounded check, not proof of coverage for every input company.
