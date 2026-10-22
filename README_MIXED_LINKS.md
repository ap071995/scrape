# Mixed LinkedIn links — v42.8.2

Double-click `Run_Scraper.vbs` for the application without a black console window. Keep all release
files together; `Scraper.py` imports the other Python modules.

## Start a run

1. In **Setup**, use **Open Chrome / Sign In** if you need to sign in.
2. Select **Link Scraper** (now the default).
3. Choose Excel/CSV containing normal LinkedIn `/in/` URLs, Sales Navigator
   `/sales/lead/` URLs, or both. Both types can be mixed in one column.
4. Choose an output workbook and target saved lead count, then start.

**Sales Navigator** mode remains available for a Sales Navigator search URL.
Individual Sales Navigator lead links belong in the **Link Scraper** file.

## How mixed links work

- Normal profile links go directly to the existing profile scraper.
- Sales Navigator leads open sequentially in the existing LinkedIn session.
  The resolver reads the explicit **View LinkedIn profile** action, including
  the profile's **Open actions overflow menu**. It never manufactures a public
  URL from a lead ID or guesses from a person's name.
- Duplicate input links and different formats that resolve to the same public
  URL are handled without another profile scrape in the same run.
- Successful mappings persist beside the output workbook in
  `.linkedin_link_mappings.sqlite3`. Repeated runs reuse them without opening
  Sales Navigator. Mappings do not expire automatically.
- **Options → Refresh saved LinkedIn profiles** bypasses and invalidates the
  mapping for each encountered Sales Navigator link before resolving it again.
  Use this if a person's public URL changes. Failed lookups are not cached.
- A saved row under the Sales Navigator URL can be updated to its resolved URL
  and reused. Pre-existing duplicate rows are not deleted automatically.
- A login, verification or explicit account-limit screen stops the run.
  An unavailable lead or missing/ambiguous profile action is recorded as
  unresolved, and processing continues. Non-English action labels may require
  a normal profile URL supplied directly.

LinkedIn remains one session with the existing page waits. The first resolution
adds a Sales Navigator visit; no setting guarantees freedom from account limits.

## Find results and diagnose missing links

The **Leads** sheet retains its existing 42 columns. **Prospect LinkedIn Link**
contains the resolved public URL for newly processed mixed-link rows.

The **Link Resolution** companion sheet records each attempted original URL,
resolved URL, input type, resolution status, processing status, reason and time.
It includes invalid URLs, duplicates and unresolved links even if no lead row
could be saved. A row is updated on retry. Links after the saved target or after
Stop are not attempted. The recovery journal protects audit entries as well as
lead records if Excel cannot be saved or the process is interrupted.

## Refreshed interface

- **Setup:** browser connection, input source, output and saved-lead target.
- **Options:** FAST/SMART, headless, concurrent enrichment, local-office lookup,
  saved-profile refresh and optional domain checks.
- **Tools:** offline benchmark, optional local LLM status and session tools.
- Start/Pause/Resume/Stop stay below the scrollable settings area.
- High-contrast progress cards, latest-lead preview and an expanding activity log
  make progress easier to follow. Pause/Resume/Stop are disabled when idle.

Headless remains off by default. Sign in visibly first, then enable it in Options.

## Validation commands

```text
py -3 -m unittest discover -s . -p "test_*.py" -v
Run_Scraper.bat --check
```

`validate_v42_8.py --live-sales-link <one-sales-lead-url>` optionally checks one
live mapping using the application's saved session. It writes local diagnostics
and does not run the full enrichment pipeline. Browser fixtures in the tests
exercise the full mixed-input-to-Excel pipeline without accessing websites.
