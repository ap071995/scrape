# Scraper v42.7.0: Maps, resume and browser options

Start the application by double-clicking **Run_Scraper.bat**. Keep every release
file in the same folder. Existing Leads workbooks keep their 42-column layout.

## Maps location selection

**Find company office near prospect** is on by default. Available prospect city,
state and country guide the Maps search. The opened business must still pass the
company name/domain and geography checks. A city match never proves the prospect
works at that particular office.

If a saved company address is in another location, a validated local result
replaces the complete Address / Zip / Phone / source-link group. The previous
group is preserved in **Previous Company Location**. This prevents a London
phone from being attached to a Zurich address. If local validation fails, the
existing group stays in Leads and Maps Evidence explains the failure. Disable
the option to keep searches anchored to an existing company address.

## Browser options

- **Run browsers headless** hides the browser windows. It is off by default.
  Use **Open Chrome / Sign In** to authenticate first; that button always opens
  a visible window. Runs reuse the application's dedicated Chrome profile.
  If a headless run reaches sign-in or verification, it stops and requests a
  visible sign-in. Headless mode is not a method of avoiding detection.
- **Two concurrent enrichment sessions** is on by default. Maps runs in one
  browser and company facts plus business email run in another. Each thread
  owns its own Playwright runtime, browser and pages. These sessions have no
  LinkedIn cookies and block LinkedIn requests, including redirects.
  Uncheck this option to run the enrichment stages sequentially.
- LinkedIn retains a single browser session and its existing page wait/scroll
  settings. Each record's enrichment completes before the next profile starts.
  Faster enrichment may shorten total time between profiles; this setting is
  not an account restriction guarantee or a published safe request rate.

## Five performance priorities implemented

1. Normalize and deduplicate input links; index saved rows by URL. Complete saved
   profiles go directly to missing-field enrichment. A missing State or company
   website alone no longer forces another profile visit.
2. Save captured profile evidence for reuse and re-parsing. Company URLs and
   capture dates are stored in **Profile State**, so company-only repairs can
   use the company URL without reopening the profile. A normal Experience-page
   fallback is still allowed when the initial profile does not expose the role.
3. Persist company About results, public company facts, email pattern evidence
   and Maps lookups in `.scraper_cache_v42_7.sqlite3` beside the output. Company
   names are rechecked before reuse. Maps cache keys also contain locality,
   domain and existing office, preventing cross-office contact reuse.
4. Run Maps and company/email enrichment concurrently in isolated sessions.
   Only the main worker writes Excel. Cache hits, parsing and evidence checks
   work locally.
5. Write `<output>.performance.json` after each run and log the same counters:
   elapsed time, profile/company time, enrichment time, navigation counts, cache
   hits, avoided profile visits and XLSX writes. Parallel stage durations can
   overlap, so their sum is not total elapsed time.

Successful company, Maps and email-pattern cache entries last up to seven days;
incomplete/failed entries expire after fifteen minutes. Profile captures last
thirty days. Dated saved profiles older than thirty days are refreshed. Legacy
saved rows without capture dates are reused with unknown age. Select **Refresh
saved LinkedIn profiles** to force profile recapture.

Unchanged unresolved rows do not consume the saved/updated-row target. A recent
enrichment attempt with the same fields/options is skipped for fifteen minutes
so a new run can reach new links. Changing company, locality, contact fields or
the relevant options invalidates that skip.

## Excel recovery

Application runs checkpoint XLSX after ten changes or at the next checkpoint
opportunity after 45 seconds; final and recovery saves flush all changes. A row
is first written and flushed to `<output>.xlsx.pending.jsonl`. If the process
ends before the XLSX checkpoint, reopening that output replays the journal.
Replay is idempotent and retains the evidence sheets. Do not remove a pending
journal before reopening its workbook. Close Excel when the app writes results.

The Python `WorkbookManager` API keeps its previous immediate-save default for
existing scripts; pass `checkpoint_every=10` to enable batching there.

## Validation commands

```powershell
.\Run_Scraper.bat --check
py -3 -m unittest discover -s . -p "test_*.py" -v
py -3 benchmark_runtime.py
py -3 validate_v42_7.py --live-maps
```

The first three commands need no live LinkedIn data. The final command explicitly
visits Google Maps in a fresh headless browser and creates a separate UBS check
workbook under Output/v42.7_validation/live_maps. It never visits LinkedIn.
