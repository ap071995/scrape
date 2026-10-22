# Version 42.9.0 — imports, email budget and worksheet controls

## Importing more than 60 links

Choose an Excel (.xlsx/.xlsm) or CSV file under **Setup**. The app reads all sheets
and columns in the background, then shows the number of unique valid links.
Normal LinkedIn profiles and Sales Navigator lead links can be mixed.

**Target saved leads** is separate from file size; its default remains 25.
Enter a larger target, or select **All imported links (Link Scraper only)**.
LinkedIn continues at its existing pace, with the existing access/error stops.
Importing a large file does not promise that every profile is accessible or will
be saved. Duplicates, access restrictions and unresolved profiles still apply.

No 60-row cap was found in the previous code. Its full-sheet scan could become
very slow when formatting extended into empty rows/columns. This release reads
actual spreadsheet cells, including embedded hyperlinks and HYPERLINK formulas.
The original report of stopping at exactly 60 remains unconfirmed without the
affected file/run log. The count and explicit target now make that distinction visible.

## Email lookup

Each email lookup has a 15-second network/search budget, including any remaining
time used for optional domain checks. Slow pages cannot start a fresh full-length
timeout after this budget. Browser startup, company research, Maps, LinkedIn and
user pauses are separate operations; 15 seconds is not a whole-lead promise.
Browser cleanup and local evidence parsing may add a small amount of time.

Published evidence still comes first. Pattern-generated addresses remain labelled
unverified. If time runs out before a usable result is found, Email Status shows
**Time limit reached (unverified)**. It does not count as a negative mailbox result
and is not cached as a completed lookup. Retry later when needed.

## Progress

Open **Process progress** for each stage: import, browser, link resolution,
profile, company, Maps, company facts, email, saving, Sales search results,
SMART validation and benchmark. A moving bar means a request is in progress;
its duration cannot be predicted. Finished counts represent completed calls,
not guaranteed field matches. The overall bar tracks the saved-lead target.
**Latest lead** includes the email and its status. **Activity** has detailed logs.

## Custom columns and manual rows

1. Select the output workbook under **Setup**.
2. Click **Worksheet · Add rows / columns**.
3. Use **Add column** to append a custom field after the standard 42 columns.
4. Use **Add rows** to create 1–1,000 manual rows, or **Edit selected row** to enter values.
5. Changes are saved through the same recovery journal as scraped rows.

The editor displays 100 rows per page and has horizontal/vertical scrolling.
Custom columns are manual fields; adding a name does not create a new scraping
rule. Manual rows are not automatically scraped. To scrape their URLs, select
the saved workbook as a Link Scraper input file in a later run.

Stop an active run before editing its worksheet. Close Excel before saving.
Custom values survive later enrichment updates and reopening. The required
standard headers must remain intact. Keep using this release for extended
workbooks; older releases required exactly 42 columns.

Start the application with **Run_Scraper.vbs**. Sandbox settings and the
windowless launcher are preserved.
