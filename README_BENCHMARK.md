# Benchmark saved scraper results

The benchmark samples saved results and produces a field-level review report. It makes **no website requests** and does not change scraping delays or account activity. Run it after a normal collection session, or use an existing saved workbook.

## Run

Use **Build Offline Benchmark** under Output Workbook in the application, or run the command from the application folder:

```powershell
python benchmark.py --input "Output/Link_Scraper_Results.xlsx" --output-dir "Output/Benchmarks" --sample-size 150
```

The default target is 150 profiles. You may choose 1–200, which also supports small diagnostic runs. If fewer unique saved records are available, the report states the shortfall. It does not fetch or invent extra profiles.

Each run creates a new folder, preserving the input and previous review files. The application accepts `.xlsx`, `.csv`, and `.tsv`. The source must have a `Prospect LinkedIn Link` column. XLSX input uses `Leads` if present, otherwise the first worksheet. Review XLSX input must have a `Review` worksheet.

## Files produced

- **benchmark_report.md:** concise coverage and accuracy report, sample size, limits, and baseline changes.
- **benchmark_review.xlsx:** field results, editable review form, selected profiles, recorded field evidence, and instructions.
- **field_results.json / field_results.csv:** results for follow-up analysis.
- **review_template.csv:** alternative to editing the workbook.
- **baseline_changes.csv:** changes between matched current and baseline profiles, when a baseline is supplied.

The workbook is a report snapshot. Edit its yellow review columns, save it, and rerun the benchmark to recalculate results:

```powershell
python benchmark.py --input "Output/Link_Scraper_Results.xlsx" --output-dir "Output/Benchmarks" --sample-size 150 --reviews "Output/Benchmarks/benchmark_PREVIOUS/benchmark_review.xlsx"
```

## How to review

Each selected profile has one row per field in the `Review` worksheet. Review the fields for which you have independent evidence. Record the evidence URL and useful notes in `Review Source` and `Review Notes`.

| Situation | Reviewed | Expected State | Expected Value |
| --- | --- | --- | --- |
| Independently established value | Yes | Value | The checked value |
| Independently established absence | Yes | Confirmed absent | Leave empty |
| Cannot establish the answer | No, or Yes with Unknown | Unknown | Leave empty |

A failed search does not establish that a value is absent. Mark `False Match = Yes` when a reviewed known value belongs to the wrong person, company, or location. That result counts as incorrect even if the text happens to match.

Comparison ignores case and extra whitespace. It does not use fuzzy matching for correctness. For example, an alternative company spelling or website spelling should be reconciled by the reviewer before using the result as a verified expectation. Revenue comparisons should use the same reporting year and currency.

CSV exports prefix formula-like values with an apostrophe so spreadsheet applications cannot execute them. Keep the `CSV Escaped` column when importing the exported review CSV: the benchmark removes that protective prefix during comparison. XLSX values are stored as literal text. Neither input formulas nor external workbook links are executed.

## What the numbers mean

- **Coverage:** filled cells divided by saved or sampled records. A filled cell may still be wrong.
- **Reviewed Known:** cells explicitly reviewed with a known expected value or confirmed absence.
- **Accuracy:** correct cells divided by Reviewed Known. Blank means not evaluated, not zero accuracy.
- **Unknown:** unreviewed cells and explicitly unknown expectations. They are excluded from accuracy.
- **False Matches:** explicitly marked incorrect person/company/location associations among reviewed known cells.
- **Missing Reasons:** recorded causes only. If no cause is saved, the report says `No evidence recorded`; it does not invent a reason.

The report keeps all-saved coverage and selected-sample results separate. Unmatched review IDs are reported and excluded.

## Sampling and matching

The sampler groups saved records by industry, company size, country, completeness, and recorded experience layout. It visits the groups in a seeded order, favors different companies, and then relaxes the default three-record company cap only when needed to fill the remaining sample places. The same saved data and seed produce the same sample. Use `--seed` and `--company-cap` to reproduce a deliberate variation.

LinkedIn country subdomains, HTTP/HTTPS, query parameters, and trailing slashes are normalized for duplicate detection. When duplicate profile rows conflict, the most complete single row is kept; conflicting rows are not combined. Records without a valid profile key are retained separately and counted in the summary. They cannot be reliably matched to a baseline.

This is a **diversity sample of the available saved data**, not a statistically weighted estimate of all prospects. Unknown attributes remain unknown. In particular, grouped job histories cannot be inferred from a flat job-title field. To evaluate that case, record an `Experience Layout` or `Role Layout` value in the saved input.

## Email and Maps context

For workbooks from the updated application, the benchmark reads companion sheets:

- **Email Details:** published addresses and generated candidates remain separate. Domain/MX and mailbox status appear as context. An MX result does not establish that an individual mailbox exists.
- **Maps Evidence:** company location type/address/phone are scored only when the saved candidate is explicitly eligible for merge. Held or rejected candidates remain missing, with the recorded outcome as the reason. Prospect location and `Prospect Works Here Verified` remain separate context.
- **Email Evidence / Company Evidence:** matching recorded source URLs support review, but their presence does not establish correctness.
- **Field Status:** optional rows with profile link, field, and reason/status support missing-field diagnostics.

Companion records must match the profile and company. Company-level evidence additionally rejects conflicting recorded websites. Legacy email values are classified only when their saved status explicitly identifies a published or generated origin. Unclassified legacy addresses are not guessed into either category.

## Compare a previous result

```powershell
python benchmark.py --input "Output/current.xlsx" --baseline "Output/previous.xlsx" --reviews "Output/Benchmarks/benchmark_PREVIOUS/benchmark_review.xlsx" --output-dir "Output/Benchmarks"
```

The comparison matches normalized profile links and labels fields `Newly filled`, `Became missing`, or `Changed value`. A newly filled field is not automatically a successful repair. `Current Assessment` in the change file and `newly_filled_review_results` in JSON show which changes have actually been reviewed as correct, incorrect, or unknown. Baseline comparison uses the same saved ground truth supplied by the reviewer.

## Regression tests

```powershell
python -m unittest test_benchmark -v
```

Tests cover sampling, deduplication, coverage/accuracy denominators, explicit absence, false matches, rejected Maps candidates, email provenance, baseline changes, malformed review input, formula-safe output, and byte-for-byte preservation of input and prior review files. Network creation is blocked in the end-to-end no-network test.
