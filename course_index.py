#!/usr/bin/env python3
"""Offline course extraction. Import is inert: no Canvas, network, or sync calls."""
import argparse
from collections import Counter
import csv
from datetime import datetime, timezone
import hashlib
import io
import json
from pathlib import Path
import re
import sqlite3
from urllib.parse import quote
from bs4 import BeautifulSoup

BASE = Path(__file__).resolve().parent
ARCHIVE = BASE / 'course-archive'
VERSION = '1.0'
POLICY = re.compile(r'\b(?:generative|artificial intelligence|chatgpt|AI|academic integrity|honor code|plagiar\w*|unauthorized|collaborat\w*|individual work|closed.book|open.book|assessment|exam\w*|syllabus)\b', re.I)
RESTRICTED = re.compile(r'solutions?[ _-]*(?:manual|key)|solutions?manual|\bsolutions?\b|(?:^|_)Solutions', re.I)
STOP = set('the and for with that this from are will have your you not can all which their they been was were has but also our these those into more than course class students student mit sloan www http https com edu pdf page copyright reserved rights use one two three about should would could each such other some may must how what when where who why its it as is to of in on a an or be by at if we do no us so'.split())


def html_segments(html, pointer='/body'):
    soup = BeautifulSoup(html or '', 'html.parser')
    for node in soup(['script', 'style', 'link', 'noscript']):
        node.decompose()
    return [{'locator': f'{pointer} / text line {i}', 'text': text.strip()}
            for i, text in enumerate(soup.get_text('\n', strip=True).splitlines(), 1)
            if text.strip()]


def load(path, default=None):
    return json.loads(path.read_text(encoding='utf-8')) if path.exists() else default


def save(path, obj):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(obj, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')


def digest(path):
    h = hashlib.sha256()
    with path.open('rb') as f:
        for block in iter(lambda: f.read(1024 * 1024), b''):
            h.update(block)
    return h.hexdigest()


def object_hash(obj):
    return hashlib.sha256(json.dumps(obj, sort_keys=True, ensure_ascii=False).encode()).hexdigest()


def read_text(path, warnings):
    data = path.read_bytes()
    try:
        return data.decode('utf-8-sig')
    except UnicodeDecodeError:
        warnings.append('Non-UTF-8 text decoded as cp1252; verify encoding.')
        return data.decode('cp1252', errors='replace')


def extract_file(path):
    """Returns segments, status, warnings. Never executes code, macros or formulas."""
    segments, warnings = [], []
    def add(locator, text):
        text = str(text).replace('\x00', '').strip()
        if text:
            segments.append({'locator': locator, 'text': text})
    ext = path.suffix.lower()
    try:
        if ext == '.pdf':
            import pymupdf
            with pymupdf.open(path) as doc:
                if doc.needs_pass:
                    return [], 'failed', ['Encrypted PDF requires password; no bypass attempted.']
                for i, page in enumerate(doc, 1):
                    try:
                        text = page.get_text(sort=True)
                        add(f'PDF page {i}', text)
                        if len(re.sub(r'\W', '', text)) < 20:
                            reason = 'image present: possible scan' if page.get_images() else 'blank, vector-only, or low-text page'
                            warnings.append(f'PDF page {i}: {reason}; no OCR performed')
                    except Exception as e:
                        warnings.append(f'PDF page {i}: extraction failed: {type(e).__name__}: {e}')
        elif ext == '.docx':
            from docx import Document
            from docx.table import Table
            doc = Document(path)
            for i, block in enumerate(doc.iter_inner_content(), 1):
                if isinstance(block, Table):
                    for j, row in enumerate(block.rows, 1):
                        add(f'DOCX body block {i} / table row {j}', ' | '.join(c.text for c in row.cells))
                else:
                    add(f'DOCX body block {i} / paragraph ({block.style.name})', block.text)
            for i, section in enumerate(doc.sections, 1):
                for kind in ['header', 'footer', 'first_page_header', 'first_page_footer', 'even_page_header', 'even_page_footer']:
                    part = getattr(section, kind)
                    if not part.is_linked_to_previous:
                        for j, para in enumerate(part.paragraphs, 1):
                            add(f'DOCX section {i} / {kind} paragraph {j}', para.text)
            warnings.append('DOCX has logical block locators, not rendered page numbers; images, text boxes, revisions, footnotes not guaranteed.')
        elif ext in ('.xlsx', '.xlsm'):
            from openpyxl import load_workbook
            book = load_workbook(path, read_only=True, data_only=False, keep_links=False)
            try:
                for sheet in book:
                    for i, row in enumerate(sheet.iter_rows(), 1):
                        values = [f'{c.coordinate}={c.value}' for c in row if c.value is not None]
                        if values:
                            add(f'XLSX sheet {sheet.title!r} / row {i}', ' | '.join(values))
            finally:
                book.close()
            warnings.append('Formula strings only; formulas/macros not executed. Charts/images not extracted.')
        elif ext == '.xlsb':
            from pyxlsb import open_workbook
            with open_workbook(str(path)) as book:
                for name in book.sheets:
                    with book.get_sheet(name) as sheet:
                        for row in sheet.rows():
                            values = [f'column {c.c + 1}={c.v}' for c in row if c.v is not None]
                            if values:
                                add(f'XLSB sheet {name!r} / row {row[0].r + 1}', ' | '.join(values))
            warnings.append('XLSB stored cell values only; no formula evaluation, macros, chart or image extraction.')
        elif ext == '.pptx':
            from pptx import Presentation
            from pptx.enum.shapes import MSO_SHAPE_TYPE
            def shape_text(shapes):
                chunks = []
                for shape in shapes:
                    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                        chunks.extend(shape_text(shape.shapes))
                    if shape.has_text_frame:
                        chunks.append(shape.text)
                    if shape.has_table:
                        chunks.extend(' | '.join(c.text for c in row.cells) for row in shape.table.rows)
                    if shape.has_chart:
                        for series in shape.chart.series:
                            chunks.append(f'{series.name}: {list(series.values)}')
                return chunks
            for i, slide in enumerate(Presentation(path).slides, 1):
                add(f'PPTX slide {i}', '\n'.join(shape_text(slide.shapes)))
                if slide.has_notes_slide:
                    add(f'PPTX slide {i} / speaker notes', slide.notes_slide.notes_text_frame.text)
            warnings.append('Text/tables/chart values and notes extracted; diagrams/images not interpreted.')
        elif ext == '.ipynb':
            nb = load(path)
            for i, cell in enumerate(nb.get('cells', []), 1):
                source = cell.get('source', '')
                add(f'Notebook cell {i} ({cell.get("cell_type", "unknown")}) / /cells/{i-1}/source', ''.join(source) if isinstance(source, list) else source)
            warnings.append('Notebook source only; NOT executed. Outputs and execution history deliberately omitted.')
        elif ext in ('.csv', '.tsv'):
            text = read_text(path, warnings)
            try:
                dialect = csv.Sniffer().sniff(text[:16384], delimiters=',;\t|')
            except csv.Error:
                dialect = csv.excel_tab if ext == '.tsv' else csv.excel
            rows = csv.reader(io.StringIO(text), dialect)
            header = next(rows, [])
            add('CSV record 1 / schema (header assumption)', 'Columns: ' + ' | '.join(header))
            for i, row in enumerate(rows, 2):
                add(f'CSV record {i}', ' | '.join(f'{header[j] if j < len(header) else "column " + str(j+1)}={v}' for j, v in enumerate(row)))
            warnings.append('CSV locators use logical records (quoted multiline fields can span physical lines); first row assumed header.')
        elif ext in ('.html', '.htm'):
            segments = html_segments(read_text(path, warnings), 'HTML')
        elif ext in ('.txt', '.md', '.py', '.r', '.sql'):
            for i, line in enumerate(read_text(path, warnings).splitlines(), 1):
                add(f'Text line {i}', line)
        else:
            return [], 'unsupported', [f'{ext or "extensionless"}: metadata only; image OCR / legacy-format conversion not attempted.']
    except Exception as e:
        warnings.append(f'{type(e).__name__}: {e}')
        return segments, 'partial' if segments else 'failed', warnings
    status = 'extracted' if segments else ('scanned-or-no-text' if ext == '.pdf' else 'empty')
    if ext == '.pdf' and warnings and segments:
        status = 'partial'
    return segments, status, warnings


def terms(segments, count=8):
    frequencies = Counter(w.lower() for s in segments for w in re.findall(r'[A-Za-z][A-Za-z-]{3,}', s['text']) if w.lower() not in STOP)
    return [word for word, _ in frequencies.most_common(count)]


def policy_excerpt(segment):
    text = segment['text']
    match = POLICY.search(text)
    if not match:
        return None
    start, end = max(0, match.start() - 180), min(len(text), match.end() + 700)
    return {'locator': segment['locator'], 'excerpt': ('…' if start else '') + text[start:end] + ('…' if end < len(text) else '')}


def course_folder(archive, course):
    """Use a direct, non-symlinked archive child for each course."""
    name = str(course.get('folder', course['id']))
    if not re.fullmatch(r'[A-Za-z0-9_-]+', name):
        raise ValueError('Unsafe course folder')
    folder = archive / name
    if folder.resolve() != folder:
        raise ValueError('Course folder must stay inside the archive')
    return folder


def build(archive=ARCHIVE, output=BASE / '.local-index'):
    archive, output = Path(archive).resolve(), Path(output).resolve()
    if archive == output or archive in output.parents:
        raise ValueError('Output must be separate from archive')
    manifest = load(archive / 'manifest.json')
    previous = load(output / 'catalog.json', {'records': []})
    output.mkdir(parents=True, exist_ok=True)
    records, courses = [], []
    temp_db = output / 'search.building.sqlite3'
    if temp_db.exists():
        temp_db.unlink()
    conn = sqlite3.connect(temp_db)
    conn.execute('CREATE VIRTUAL TABLE segments USING fts5(record_key UNINDEXED, course_id UNINDEXED, kind UNINDEXED, title, locator UNINDEXED, source_path UNINDEXED, source_url UNINDEXED, text, tokenize="unicode61")')

    def record(course, kind, rid, title, path, url, updated=None, payload=None, segments=None, filemeta=None):
        cid = str(course['id'])
        if not re.fullmatch(r'[A-Za-z0-9_.-]+', str(rid)):
            raise ValueError(f'Unsafe record id: {rid}')
        key = f'{cid}:{kind}:{rid}'
        path = Path(path).resolve()
        course_root = course_folder(archive, course)
        if not path.is_relative_to(course_root):
            raise ValueError(f'Source escapes course namespace: {path}')
        exists = path.is_file()
        warnings = []
        restricted = bool(RESTRICTED.search(title))
        sha = digest(path) if kind == 'file' and exists else object_hash(payload) if payload is not None else None
        if restricted:
            segments, status = [], 'policy-check-before-use'
            warnings.append('Solution material: metadata only. Authorization not inferred from archive presence; check course policy before use.')
        elif kind == 'file':
            segments, status, warnings = extract_file(path) if exists else ([], 'missing', ['Archive file missing'])
        else:
            segments = segments or []
            status = 'extracted' if segments else 'empty'
        if not restricted:
            for supplement in load(output / 'supplements.json', []):
                if supplement['record_key'] == key:
                    if supplement['source_sha256'] == sha:
                        segments.extend(supplement['segments'])
                        status = 'partial'
                        warnings.append('Includes explicitly labelled, source-hash-bound visual transcription from supplements.json. Only the specified image/table was reviewed; other images are not extracted.')
                    else:
                        warnings.append('Stale visual supplement skipped: source hash changed; re-review required.')
        if filemeta and filemeta.get('sha256') and sha != filemeta['sha256']:
            warnings.append('SHA256 differs from download manifest; verify source before relying on content.')
        if filemeta and exists and filemeta.get('bytes') is not None and path.stat().st_size != filemeta['bytes']:
            warnings.append('Byte size differs from download manifest.')
        extract_path = output / cid / 'text' / f'{kind}-{rid}.json'
        excerpts = [e for s in segments if (e := policy_excerpt(s))]
        info = {
            'key': key, 'course_id': cid, 'course_name': course['name'], 'kind': kind,
            'source_id': str(rid), 'file_id': str(rid) if kind == 'file' else None,
            'title': title, 'source_path': str(path), 'source_url': url,
            'updated_at': updated, 'sha256': sha,
            'hash_scope': 'file_bytes' if kind == 'file' else 'whitelisted_source_fields',
            'bytes': path.stat().st_size if exists and kind == 'file' else None,
            'archive_status': 'present' if exists else 'missing',
            'download_status': (filemeta or {}).get('status', 'archived_json' if kind != 'file' else 'unlisted_local'),
            'extraction_status': status, 'indexed': bool(segments),
            'review_status': 'not_fully_reviewed', 'read_status': 'machine_extracted_only' if segments else 'metadata_only',
            'segment_count': len(segments), 'characters': sum(len(s['text']) for s in segments),
            'extraction_path': str(extract_path), 'topics': terms(segments),
            'locator_samples': [{'locator': s['locator'], 'excerpt': re.sub(r'\s+', ' ', s['text'])[:170]} for s in segments[:3]],
            'syllabus_or_policy_pointer': bool(re.search(r'syllab|polic|academic integrity|course information', title, re.I)),
            'policy_candidates': excerpts[:30], 'policy_candidate_total': len(excerpts),
            'warnings': warnings,
        }
        save(extract_path, {'record_key': key, 'review_status': 'not_fully_reviewed', 'segments': segments})
        conn.executemany('INSERT INTO segments VALUES (?,?,?,?,?,?,?,?)', [(key, cid, kind, title, s['locator'], str(path), url, s['text']) for s in segments])
        records.append(info)

    for course in manifest['courses']:
        cid = str(course['id'])
        if not cid.isdigit():
            raise ValueError('Course IDs must be numeric')
        folder = course_folder(archive, course)
        canonical = f'https://canvas.mit.edu/courses/{cid}'
        course_info = {'id': cid, 'name': course['name'], 'source_url': canonical,
                       'archive_errors': course.get('errors', []), 'external_links': course.get('external_links', []), 'missing_module_targets': []}
        courses.append(course_info)
        seen_paths = set()
        for meta in course.get('files', []):
            path = Path(meta['path']).resolve()
            seen_paths.add(path)
            record(course, 'file', meta['id'], meta['name'], path, f'{canonical}/files/{meta["id"]}', meta.get('updated_at'), filemeta=meta)
        # Discover new local files in numeric file-ID subdirectories without merging course namespaces.
        for filedir in sorted(folder.iterdir()):
            if filedir.is_dir() and filedir.name.isdigit():
                for path in sorted(filedir.iterdir()):
                    if path.is_file() and path.resolve() not in seen_paths:
                        record(course, 'file', filedir.name, path.name, path, f'{canonical}/files/{filedir.name}')
        for i, item in enumerate(load(folder / 'assignments.json', [])):
            rid = item['id']
            payload = {k: item.get(k) for k in ['id', 'name', 'description', 'updated_at']}
            record(course, 'assignment', rid, item.get('name', f'Assignment {rid}'), folder / 'assignments.json', f'{canonical}/assignments/{rid}', item.get('updated_at'), payload, html_segments(item.get('description'), f'/{i}/description (assignment {rid})'))
        for path in sorted((folder / 'page_bodies').glob('*.json')):
            item = load(path)
            rid = item.get('page_id', path.stem)
            payload = {k: item.get(k) for k in ['page_id', 'title', 'body', 'url', 'updated_at']}
            record(course, 'page', rid, item.get('title', path.stem), path, f'{canonical}/pages/{quote(item.get("url", path.stem))}', item.get('updated_at'), payload, html_segments(item.get('body'), '/body'))
        syllabus = load(folder / 'syllabus.json', {})
        record(course, 'syllabus', cid, 'Canvas syllabus body', folder / 'syllabus.json', f'{canonical}/assignments/syllabus', syllabus.get('updated_at'), {'syllabus_body': syllabus.get('syllabus_body')}, html_segments(syllabus.get('syllabus_body'), '/syllabus_body'))
        local_files = {r['file_id'] for r in records if r['course_id'] == cid and r['kind'] == 'file'}
        # The archive sanitizes filename characters (e.g. '+' -> '_'); Canvas slugs do not change.
        page_slugs = {load(p).get('url', p.stem) for p in (folder / 'page_bodies').glob('*.json')}
        for i, module in enumerate(load(folder / 'modules.json', [])):
            rid = module['id']
            items_path = folder / 'module_items' / f'{rid}.json'
            items = load(items_path, [])
            safe_items = [{k: item.get(k) for k in ['id', 'title', 'type', 'content_id', 'page_url']} for item in items]
            seg = [{'locator': f'/{i}/name (module {rid})', 'text': module['name']}]
            for j, item in enumerate(safe_items):
                seg.append({'locator': f'module_items/{rid}.json /{j} (item {item["id"]})', 'text': f'{item["type"]}: {item["title"]}'})
                if (item['type'] == 'File' and str(item['content_id']) not in local_files) or (item['type'] == 'Page' and item['page_url'] not in page_slugs):
                    course_info['missing_module_targets'].append(item)
            record(course, 'module', rid, module['name'], folder / 'modules.json', f'{canonical}/modules/{rid}', None, {'id': rid, 'name': module['name'], 'items': safe_items}, seg)
    keys = [r['key'] for r in records]
    if len(keys) != len(set(keys)):
        raise ValueError('Duplicate source identifiers within a course; refusing ambiguous catalog')
    old, new = ({r['key']: r for r in rows} for rows in [previous['records'], records])
    changes = {'added': sorted(new.keys() - old.keys()), 'removed': sorted(old.keys() - new.keys()),
               'changed': sorted(k for k in new.keys() & old.keys() if (new[k]['sha256'], new[k]['updated_at']) != (old[k]['sha256'], old[k]['updated_at']))}
    catalog = {'schema_version': 1, 'extractor_version': VERSION, 'generated_at': datetime.now(timezone.utc).isoformat(),
               'archive_root': str(archive), 'manifest_sha256': digest(archive / 'manifest.json'),
               'archive_refreshed_at': manifest.get('refreshed_at'), 'review_status': 'extracted/indexed NOT fully reviewed',
               'courses': courses, 'records': records, 'changes': changes}
    conn.commit()
    conn.close()
    temp_db.replace(output / 'search.sqlite3')
    save(output / 'catalog.json', catalog)
    save(output / 'changes.json', changes)
    write_markdown(catalog, output)
    # Publish baseline last and atomically: searches never write this file.
    baseline = {'courses': [{'id': int(c['id']), 'name': c['name'], 'files': [
        {'id': int(r['file_id']), 'updated_at': r['updated_at'], 'size': r['bytes'], 'name': r['title']}
        for r in records if r['course_id'] == c['id'] and r['kind'] == 'file'
    ]} for c in courses]}
    save(output / 'upload_baseline.building.json', baseline)
    (output / 'upload_baseline.building.json').replace(output / 'upload_baseline.json')
    return catalog


def md(text):
    return str(text).replace('|', '\\|').replace('\n', ' ').replace('[', '\\[').replace(']', '\\]')


def write_markdown(catalog, output):
    root = ['# Course source index', '', '**Extracted/indexed — NOT fully reviewed.** No learner history, grades, or completion tracking.',
            '', 'Offline archive only. Downloaded ≠ text-extracted ≠ indexed ≠ semantically read. Solution materials are metadata-only, policy-check-before-use.',
            '', f'Generated: {catalog["generated_at"]}; archive refreshed: {catalog["archive_refreshed_at"]}.',
            '', '| Course | Archived files | Indexed sources | Text segments |', '|---|---:|---:|---:|']
    for course in catalog['courses']:
        cid = course['id']
        records = [r for r in catalog['records'] if r['course_id'] == cid]
        files = [r for r in records if r['kind'] == 'file']
        root.append(f'| [{md(course["name"])}]({cid}/INDEX.md) | {len(files)} | {sum(r["indexed"] for r in records)} | {sum(r["segment_count"] for r in records)} |')
        lines = [f'# {course["name"]}', '', f'Course ID: `{cid}` · [Canvas]({course["source_url"]}) · [All courses](../INDEX.md)',
                 '', '**Extracted/indexed — NOT fully reviewed.** Read status is machine extraction only, not proof of understanding or authorization.',
                 '', '## Archive limitations', '']
        lines += [f'- {e["endpoint"]}: {e["error"]} (recorded by prior archive fetch; no new network request).' for e in course['archive_errors']]
        lines += [f'- Missing local module target: {md(i["title"])} ({i["type"]}, ID/slug {i.get("content_id") or i.get("page_url")}).' for i in course['missing_module_targets']]
        lines += [f'- External link, not downloaded: [{md(e["title"])}]({e["url"]})' for e in course['external_links']]
        lines += ['', '## Content-derived topic hints', '', 'Frequency-derived terms, not a semantic curriculum review; inspect source locators below.', '']
        aggregate = Counter(t for r in records for t in r['topics'])
        lines += [', '.join(t for t, _ in aggregate.most_common(20)) or 'No extractable text.', '', '## Syllabus / policy pointers', '']
        pointers = [r for r in records if r['syllabus_or_policy_pointer']]
        for r in pointers:
            lines.append(f'- [{md(r["title"])}]({Path(r["source_path"]).as_uri()}) · {r["key"]} · {r["extraction_status"]}; see source entry below for locators.')
        lines += ['', 'Policy mentions are automatically detected candidates, not adjudicated rules.', '', '## Sources', '']
        for r in records:
            lines += [f'### {md(r["title"])}', '', f'`{r["key"]}` · {r["kind"]} · [Canvas source]({r["source_url"]}) · [Local archive]({Path(r["source_path"]).as_uri()}) · [Extracted text](text/{Path(r["extraction_path"]).name})',
                      '', f'- Local path: `{r["source_path"]}`',
                      f'- Archive: {r["archive_status"]}; download: {r["download_status"]}; extraction: **{r["extraction_status"]}**; indexed: {r["indexed"]}; read: {r["read_status"]}; review: **not fully reviewed**.',
                      f'- Segments: {r["segment_count"]}; updated_at: `{r["updated_at"]}`; SHA256: `{r["sha256"]}`.',
                      '- Content-derived terms: ' + (', '.join(r['topics']) or 'None; metadata only or no text.')]
            lines += [f'- {md(s["locator"])}: “{md(s["excerpt"])}”' for s in r['locator_samples']]
            lines += [f'- Warning: {md(w)}' for w in r['warnings']]
            if r['syllabus_or_policy_pointer']:
                lines += [f'- Policy candidate at **{md(s["locator"])}**: {md(s["excerpt"])}' for s in r['policy_candidates']]
            lines.append('')
        (output / cid / 'INDEX.md').write_text('\n'.join(lines) + '\n', encoding='utf-8')
    root += ['', '## Scope gaps', '', '- Source listing and access gaps are recorded per course from the supplied archive manifest; indexing does not prove complete course coverage.',
             '- Scans, images, equations, layout, missing external documents and some Office features require further inspection. No OCR or notebook execution.',
             '- Searches exclude solution contents; metadata remains in the catalog and course indexes.',
             '- Keep generated indexes private. This tool installs no scheduler and never fetches Canvas.']
    (output / 'INDEX.md').write_text('\n'.join(root) + '\n', encoding='utf-8')


def search(output, query, course=None, limit=20):
    """All query tokens must occur within one segment; ranks via SQLite FTS5 BM25."""
    tokens = re.findall(r'\w+', query, re.UNICODE)
    if not tokens:
        return []
    match = ' AND '.join('"' + t.replace('"', '""') + '"' for t in tokens)
    path = (Path(output) / 'search.sqlite3').resolve()
    conn = sqlite3.connect(path.as_uri() + '?mode=ro', uri=True)
    conn.row_factory = sqlite3.Row
    sql = 'SELECT record_key, course_id, kind, title, locator, source_path, source_url, snippet(segments,7,"[", "]"," … ",40) AS snippet FROM segments WHERE segments MATCH ?'
    params = [match]
    if course:
        sql += ' AND course_id = ?'
        params.append(str(course))
    sql += ' ORDER BY bm25(segments), record_key, rowid LIMIT ?'
    params.append(max(1, min(limit, 1000)))
    try:
        return [dict(r) for r in conn.execute(sql, params)]
    finally:
        conn.close()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest='command', required=True)
    b = sub.add_parser('build', help='Rebuild offline; only after user consent to refresh')
    b.add_argument('--archive', type=Path, default=ARCHIVE)
    b.add_argument('--output', type=Path, default=BASE / '.local-index')
    s = sub.add_parser('search', help='Literal token AND search; optional exact course ID')
    s.add_argument('query')
    s.add_argument('--course')
    s.add_argument('--limit', type=int, default=20)
    s.add_argument('--output', type=Path, default=BASE / '.local-index')
    s.add_argument('--json', action='store_true')
    args = parser.parse_args()
    if args.command == 'build':
        catalog = build(args.archive, args.output)
        print(json.dumps({'courses': len(catalog['courses']), 'records': len(catalog['records']), 'files': sum(r['kind'] == 'file' for r in catalog['records']), 'indexed_sources': sum(r['indexed'] for r in catalog['records']), 'segments': sum(r['segment_count'] for r in catalog['records']), 'statuses': dict(Counter(r['extraction_status'] for r in catalog['records'])), 'changes': {k: len(v) for k,v in catalog['changes'].items()}}, indent=2))
    else:
        hits = search(args.output, args.query, args.course, args.limit)
        if args.json:
            print(json.dumps(hits, ensure_ascii=False, indent=2))
        else:
            for hit in hits:
                print(f'{hit["record_key"]} | {hit["title"]}\n  {hit["source_path"]}\n  {hit["locator"]} | {hit["source_url"]}\n  {hit["snippet"]}\n')
            print(f'{len(hits)} hits shown (limit {args.limit}); extracted/indexed NOT fully reviewed.')


if __name__ == '__main__':
    main()
