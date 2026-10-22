"""Read actual XLSX cells, ignoring formatted empty rows and inflated dimensions."""
import csv
import posixpath
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from linkedin_links import extract_link_inputs


class ImportCancelled(Exception):
    pass


def load_link_inputs(input_path, progress=None, cancelled=None):
    path = Path(input_path)
    if not path.is_file():
        raise FileNotFoundError(f'Input file not found: {path}')
    found, seen = [], set()
    cells = 0

    def active():
        if cancelled and cancelled():
            raise ImportCancelled('File import stopped.')

    def add(value):
        nonlocal cells
        active()
        cells += 1
        for item in extract_link_inputs(value):
            if item.original not in seen:
                seen.add(item.original)
                found.append(item)
        if progress and cells % 500 == 0:
            progress(f'{cells:,} cells read · {len(found):,} links found')

    if path.suffix.lower() == '.csv':
        for encoding in ('utf-8-sig', 'cp1252', 'latin-1'):
            try:
                found.clear()
                seen.clear()
                with path.open(encoding=encoding, newline='') as stream:
                    for row in csv.reader(stream):
                        for value in row:
                            add(value)
                break
            except UnicodeDecodeError:
                continue
    elif path.suffix.lower() in ('.xlsx', '.xlsm'):
        with ZipFile(path) as archive:
            ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            rid = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'

            def relationships(part):
                folder, filename = posixpath.split(part)
                relpath = posixpath.join(folder, '_rels', filename + '.rels')
                if relpath not in archive.namelist():
                    return {}
                return {e.get('Id'): (e.get('Target'), e.get('TargetMode'))
                        for e in ET.fromstring(archive.read(relpath))}

            def resolve(part, target):
                return posixpath.normpath(posixpath.join(posixpath.dirname(part), target)) if not target.startswith('/') else target.lstrip('/')

            shared = []
            bookrels = relationships('xl/workbook.xml')
            shared_path = 'xl/sharedStrings.xml'
            for target, mode in bookrels.values():
                if target and target.endswith('sharedStrings.xml'):
                    shared_path = resolve('xl/workbook.xml', target)
            if shared_path in archive.namelist():
                with archive.open(shared_path) as stream:
                    for _, element in ET.iterparse(stream, events=('end',)):
                        if element.tag.endswith('}si'):
                            active()
                            shared.append(''.join(e.text or '' for e in element.iterfind('.//s:t', ns)))
                            element.clear()
            sheets = ET.fromstring(archive.read('xl/workbook.xml')).findall('s:sheets/s:sheet', ns)
            for sheet in sheets:
                active()
                target, mode = bookrels[sheet.get(rid)]
                if mode == 'External':
                    continue
                part = resolve('xl/workbook.xml', target)
                rels = relationships(part)
                hyperlinks = {}
                # Hyperlinks follow sheetData, so collect their targets first.
                with archive.open(part) as stream:
                    for _, element in ET.iterparse(stream, events=('end',)):
                        if element.tag.endswith('}hyperlink'):
                            target, mode = rels.get(element.get(rid), ('', ''))
                            if target and mode == 'External':
                                hyperlinks[element.get('ref', '').split(':')[0]] = target
                        if element.tag.endswith(('}row', '}hyperlink')):
                            active()
                            element.clear()
                with archive.open(part) as stream:
                    for _, element in ET.iterparse(stream, events=('end',)):
                        if element.tag.endswith('}c'):
                            if element.get('r') in hyperlinks:
                                add(hyperlinks.pop(element.get('r')))
                            formula = element.find('s:f', ns)
                            value = element.find('s:v', ns)
                            if formula is not None:
                                add('=' + (formula.text or ''))
                            elif element.get('t') == 's' and value is not None:
                                add(shared[int(value.text)])
                            elif element.get('t') == 'inlineStr':
                                add(''.join(e.text or '' for e in element.iterfind('.//s:t', ns)))
                            elif value is not None:
                                add(value.text)
                            element.clear()
                        elif element.tag.endswith('}row'):
                            active()
                            element.clear()
                for target in hyperlinks.values():
                    add(target)
    else:
        raise ValueError('Link Scraper supports .csv, .xlsx and .xlsm input files.')
    if progress:
        progress(f'{len(found):,} links found across all sheets and columns')
    return found
