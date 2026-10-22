from pathlib import Path

source = Path(r'C:/Users/Aditya.paikine/Desktop/Scraper/Scraper.py')
s = source.read_text(encoding='utf-8-sig')
def replace(old, new, count=1):
    global s
    assert s.count(old) >= count, old[:100]
    s = s.replace(old, new, count)

replace('import threading\n', 'import threading\nimport tempfile\n')
replace('APP_VERSION = "42.0.0"', 'APP_VERSION = "42.1.0"')
replace('EXCEL_CHECKPOINT_EVERY = 25', 'EXCEL_CHECKPOINT_EVERY = 1')
replace('        path = parts.path.rstrip("/")', '''        path = parts.path.rstrip("/")
        if netloc in {"linkedin.com", "www.linkedin.com"}:
            netloc = "www.linkedin.com"
            scheme = "https"
            if path.startswith("/sales/lead/"):
                path = path.split(",", 1)[0]''')
replace('    text = str(value)\n    if text == "-":', '''    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", str(value))[:32767]
    if text == "-":'''.replace('\x00','\\x00').replace('\x08','\\x08').replace('\x0b','\\x0b').replace('\x0c','\\x0c').replace('\x0e','\\x0e').replace('\x1f','\\x1f'))
start = s.index('    def __init__(self, output_path: str):', s.index('class WorkbookManager:'))
end = s.index('    def _create_header', start)
s = s[:start] + '''    def __init__(self, output_path: str):
        self.output_path = Path(output_path)
        self.seen_links: set[str] = set()
        self._dirty_rows = 0
        self._last_save_ts = time.time()
        if self.output_path.exists():
            self.wb = load_workbook(self.output_path)
            if "Leads" not in self.wb.sheetnames:
                raise ValueError("Existing workbook has no Leads sheet; choose another output file.")
            self.ws = self.wb["Leads"]
            if [c.value for c in self.ws[1]] != HEADERS:
                raise ValueError("Existing Leads headers do not match; choose another output file.")
            column = HEADERS.index("Prospect LinkedIn Link") + 1
            for row in self.ws.iter_rows(min_row=2, min_col=column, max_col=column):
                value = row[0].value
                if value and value != "-":
                    self.seen_links.add(normalize_url(str(value)))
        else:
            self.wb = Workbook()
            self.ws = self.wb.active
            self.ws.title = "Leads"
            self._create_header()
            self.save(force=True)

''' + s[end:]
replace('''        if link and link != "-":
            self.seen_links.add(link)

        raw_values''', '''        if link and link in self.seen_links:
            raise ValueError("Duplicate prospect URL: " + link)

        raw_values''')
replace('''        self._dirty_rows += 1
        now = time.time()''','''        if link and link != "-":
            self.seen_links.add(link)
        self._dirty_rows += 1
        now = time.time()''')
replace('''        self.wb.save(self.output_path)
        self._dirty_rows = 0''','''        self.ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}{self.ws.max_row}"
        # Replace only after the complete ZIP has been written successfully.
        fd, temporary = tempfile.mkstemp(suffix=".xlsx", dir=self.output_path.parent)
        os.close(fd)
        try:
            self.wb.save(temporary)
            os.replace(temporary, self.output_path)
        except PermissionError as exc:
            raise PermissionError(
                f"Cannot save {self.output_path}. Close it in Excel and retry."
            ) from exc
        finally:
            Path(temporary).unlink(missing_ok=True)
        self._dirty_rows = 0''')
replace('''    def send(self, name: str, payload: Any = None):
        self.commands.put((name, payload))''','''    def send(self, name: str, payload: Any = None):
        # These events must reach the active loop, not queue behind it.
        if name == "pause":
            self.pause_event.set()
            return
        if name == "resume":
            self.pause_event.clear()
            return
        if name in {"stop", "shutdown", "close_browser"}:
            self.stop_event.set()
            self.pause_event.clear()
            if name == "stop":
                return
        self.commands.put((name, payload))''')
replace('''        if path.exists():
            overwrite = messagebox.askyesno(''', '''        if path.exists():
            self._log("Appending to existing workbook; existing prospects will be skipped.")

        if False:  # Legacy overwrite dialog disabled: existing workbooks are resumed.
            overwrite = messagebox.askyesno(''')
# Remove dead dialog entirely.
a = s.index('        if False:  # Legacy overwrite')
b = s.index('        try:\n            self.workbook = WorkbookManager', a)
s = s[:a] + s[b:]
replace('title="Save new Excel workbook"', 'title="Select workbook to append, or enter a new filename"')
replace('''                        "w",
                        encoding="utf-8",
                    ) as handle:''','''                        "a",
                        encoding="utf-8",
                    ) as handle:''')
# No generic h1: Maps uses one for the search heading before results load.
replace("document.querySelector('h1.DUwDvf')\n                            || document.querySelector('h1');", "document.querySelector('h1.DUwDvf');")
replace('''            direct_place = (
                "/maps/place/"
                in page.url
            )''','''            direct_place = (
                "/maps/place/" in page.url
                or page.locator("h1.DUwDvf").count() > 0
            )''')
replace('''            if not (movement or {}).get(
                "moved",
                False,
            ):
                break''','''            if not (movement or {}).get("moved", False) and best:
                break''')
# Do not claim an address source if no address was extracted.
replace('''            result = {
                "company_matched": True,''','''            if self._is_unknown_value(full_address) and self._is_unknown_value(normalized_phone):
                last_reason = "Matched Maps place contains no usable address or phone."
                continue

            result = {
                "company_matched": True,''')
replace('''                "address_link": clean_cell_value(
                    place.get(
                        "Address Link",
                        "-",
                    )
                ),''','''                "address_link": (
                    clean_cell_value(place.get("Address Link", "-"))
                    if not self._is_unknown_value(full_address) else "-"
                ),''')
# Avoid substring false positives such as tax in taxi or lead in misleading.
replace('if any(n in t for n in needles):', 'if any(re.search(r"(?<!\\w)" + re.escape(n.strip()) + r"(?!\\w)", t) for n in needles):')
replace('if any(n in padded for n in needles):', 'if any(re.search(r"(?<!\\w)" + re.escape(n.strip()) + r"(?!\\w)", padded) for n in needles):')
replace('(("vice president", "vp "), "VP")', '(("vice president", "vp"), "VP")')
replace('''                "operations",
                "operational",''','''                "operations",
                "manufacturing",
                "operational",''')
# Source attribution for company revenue when captured from its page.
replace('''                try:
                    record = self._google_maps_enrich_record(''','''                if record.get("Revenue", "-") != "-" and record.get("Revenue Link", "-") == "-":
                    record["Revenue Link"] = record.get("Industry Link", "-")
                if company_error:
                    self._append_enrichment_note(record, "Company extraction failed: " + company_error)

                try:
                    record = self._google_maps_enrich_record(''')
# Save checkpoints even when navigation/audit/enrichment raises unexpectedly.
replace('''                    except Exception as exc:
                        self.status("Error")''','''                    except Exception as exc:
                        try:
                            payload["workbook"].save(force=True)
                        except Exception as save_exc:
                            self.log(f"Recovery save failed: {save_exc}")
                        self.status("Error")''',2)
Path('Scraper.py').write_text(s, encoding='utf-8')
print('Updated Scraper.py')
