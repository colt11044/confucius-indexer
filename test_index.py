"""Offline tests; fixtures live only inside this project, never the archive."""
import importlib.util
from pathlib import Path

ROOT = Path(__file__).parent

def test_html_has_locators_and_no_active_code():
    spec = importlib.util.spec_from_file_location('course_index', ROOT / 'course_index.py')
    assert spec is not None and Path(spec.origin).exists(), 'index implementation missing'
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    segments = module.html_segments('<h2>Policy</h2><p>No AI on exams.</p><script>secret()</script>', '/body')
    assert any('No AI on exams.' in s['text'] and '/body' in s['locator'] for s in segments)
    assert not any('secret' in s['text'] for s in segments)


def test_offline_build_search_preserves_namespaces_and_restrictions():
    import json
    import tempfile
    import pymupdf as fitz
    import docx
    import openpyxl
    from pptx import Presentation
    import course_index as ci
    assert hasattr(ci, 'build'), 'offline build missing'
    with tempfile.TemporaryDirectory(dir=ROOT) as tmp:
        root = Path(tmp)
        archive, output = root / 'archive', root / 'output'
        archive.mkdir()
        courses = []
        for cid in ['1', '2']:
            folder = archive / cid
            folder.mkdir()
            (folder / 'page_bodies').mkdir()
            (folder / 'page_bodies' / 'policy.json').write_text(json.dumps({'page_id': 7, 'url': 'policy+extra', 'title': 'Rules', 'body': '<h2>Course rules</h2><p>Quasar collaboration</p>', 'updated_at': 'v1'}))
            (folder / 'assignments.json').write_text(json.dumps([{'id': 9, 'name': 'Task', 'description': '<p>Quasar assignment description</p>', 'submission': {'grade': 'SECRET'}}]))
            (folder / 'syllabus.json').write_text(json.dumps({'syllabus_body': '<p>Quasar syllabus</p>'}))
            (folder / 'modules.json').write_text(json.dumps([{'id': 3, 'name': 'Overview', 'state': 'completed'}]))
            (folder / 'module_items').mkdir()
            (folder / 'module_items' / '3.json').write_text(json.dumps([{'id': 4, 'title': 'Rules', 'type': 'Page', 'page_url': 'policy+extra'}]))
            files = []
            def add(name, writer):
                fid = len(files) + 10
                path = folder / str(fid) / name
                path.parent.mkdir()
                writer(path)
                files.append({'id': fid, 'name': name, 'path': str(path), 'updated_at': 'v1', 'status': 'downloaded'})
            def pdf(path):
                d = fitz.open()
                d.new_page().insert_text((40, 40), 'Quasar PDF page one')
                d.new_page().insert_text((40, 40), 'Quasar PDF page two')
                d.save(path)
                d.close()
            def word(path):
                d = docx.Document()
                d.add_heading('Quasar Word', 1)
                d.add_table(rows=1, cols=1).cell(0, 0).text = 'Quasar table'
                d.save(path)
            def excel(path):
                d = openpyxl.Workbook()
                d.active.append(['Quasar sheet', '=1+1'])
                d.save(path)
                d.close()
            def slides(path):
                d = Presentation()
                d.slides.add_slide(d.slide_layouts[0]).shapes.title.text = 'Quasar slide'
                d.save(path)
            add('lecture.pdf', pdf)
            add('lecture.docx', word)
            add('data.xlsx', excel)
            add('slides.pptx', slides)
            add('data.csv', lambda p: p.write_text('topic,value\nQuasar,1\n'))
            add('code.ipynb', lambda p: p.write_text(json.dumps({'cells': [{'cell_type': 'code', 'source': ['Quasar = 1\n', 'raise Exception("NEVER EXECUTE")'], 'outputs': [{'text': 'IGNORE OUTPUT'}]}]})))
            add('index.html', lambda p: p.write_text('<p>Quasar HTML</p><script>NEVER EXECUTE</script>'))
            add('SolutionsManual.docx', lambda p: p.write_bytes(b'not opened'))
            add('image.png', lambda p: p.write_bytes(b'unsupported'))
            add('bad.pdf', lambda p: p.write_bytes(b'broken'))
            courses.append({'id': cid, 'name': 'Course ' + cid, 'files': files})
        (archive / 'manifest.json').write_text(json.dumps({'courses': courses}))
        output.mkdir()
        import hashlib
        first_pdf = Path(courses[0]['files'][0]['path'])
        (output / 'supplements.json').write_text(json.dumps([
            {'record_key': '1:file:10', 'source_sha256': hashlib.sha256(first_pdf.read_bytes()).hexdigest(), 'segments': [{'locator': 'PDF page 1 / visually checked image', 'text': 'VisualSupplementToken'}]},
            {'record_key': '2:file:10', 'source_sha256': 'stale', 'segments': [{'locator': 'PDF page 1', 'text': 'StaleSupplementToken'}]}
        ]))
        catalog = ci.build(archive, output)
        assert ci.search(output, 'VisualSupplementToken', course='1'), 'verified supplements must be searchable'
        assert not ci.search(output, 'StaleSupplementToken', course='2')
        assert len(catalog['courses']) == 2
        assert all(not c['missing_module_targets'] for c in catalog['courses']), 'match original page URL slug, not sanitized filename'
        baseline_path = output / 'upload_baseline.json'
        assert baseline_path.exists(), 'successful build must publish upload baseline'
        baseline = json.loads(baseline_path.read_text())
        assert len(baseline['courses']) == 2
        assert all(len(c['files']) == 10 for c in baseline['courses'])
        assert all(set(f) == {'id', 'updated_at', 'size', 'name'} for c in baseline['courses'] for f in c['files'])
        records = catalog['records']
        assert len({r['key'] for r in records}) == len(records)
        assert all(r['review_status'] == 'not_fully_reviewed' for r in records)
        restricted = [r for r in records if 'SolutionsManual' in r['title']]
        assert all(r['extraction_status'] == 'policy-check-before-use' and r['segment_count'] == 0 for r in restricted)
        assert any(r['extraction_status'] == 'failed' for r in records)
        assert any(r['extraction_status'] == 'unsupported' for r in records)
        hits = ci.search(output, 'Quasar', course='1', limit=100)
        assert hits and all(h['course_id'] == '1' for h in hits)
        assert {'file', 'page', 'assignment', 'syllabus'} <= {h['kind'] for h in hits}
        assert any(h['locator'] == 'PDF page 2' for h in hits)
        assert all(h['source_path'] and h['source_url'] for h in hits)
        assert not ci.search(output, 'SECRET', limit=100)
        assert not ci.search(output, 'IGNORE OUTPUT', limit=100)
        assert (output / 'INDEX.md').exists()
        assert all((output / cid / 'INDEX.md').exists() for cid in ['1', '2'])
        assert not ci.search(output, 'SolutionsManual', limit=100)
        # Repeat builds must be stable with no phantom changes.
        second = ci.build(archive, output)
        assert second['changes'] == {'added': [], 'removed': [], 'changed': []}
        # Baseline remains unchanged if a subsequent rebuild aborts.
        stable_baseline = baseline_path.read_bytes()
        (archive / 'manifest.json').write_text('{not valid json')
        import pytest
        with pytest.raises(json.JSONDecodeError):
            ci.build(archive, output)
        assert baseline_path.read_bytes() == stable_baseline


def test_pdf_scan_blank_and_failure_are_not_silently_successful():
    import tempfile
    import pymupdf
    import course_index as ci
    with tempfile.TemporaryDirectory(dir=ROOT) as tmp:
        path = Path(tmp) / 'blank.pdf'
        d = pymupdf.open()
        d.new_page()
        d.save(path)
        d.close()
        segments, status, warnings = ci.extract_file(path)
        assert not segments and status == 'scanned-or-no-text'
        assert any('PDF page 1' in w for w in warnings)


def test_course_folder_rejects_escape_and_symlink(tmp_path):
    import pytest
    import course_index as ci

    archive = tmp_path / 'archive'
    archive.mkdir()
    (tmp_path / 'outside').mkdir()
    with pytest.raises(ValueError, match='Unsafe course folder'):
        ci.course_folder(archive, {'id': 1, 'folder': '../outside'})
    (archive / 'linked').symlink_to(tmp_path / 'outside', target_is_directory=True)
    with pytest.raises(ValueError, match='inside the archive'):
        ci.course_folder(archive, {'id': 1, 'folder': 'linked'})


