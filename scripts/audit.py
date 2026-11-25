#!/usr/bin/env python3
"""Mechanical checks and genuine presentation rendering; never a scientific score."""
import argparse
import json
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile

from prepare import versions, write_json


def normalized(text):
    return re.sub(r'\s+', '', text).casefold()


def evidence(args):
    ledger = json.loads(Path(args.ledger).read_text(encoding='utf-8'))
    errors, warnings, checked = [], [], []
    sources = ledger.get('sources', [])
    identifiers = [x.get('id') for x in sources]
    if not sources:
        errors.append('No evidence sources supplied')
    if len(set(identifiers)) != len(identifiers) or None in identifiers:
        errors.append('Source IDs must be present and unique')
    for source in sources:
        record = {'id': source.get('id'), 'path': source.get('path')}
        path = Path(source.get('path', ''))
        if not path.is_file():
            errors.append(f'{source.get("id")}: source file missing')
            checked.append(record)
            continue
        text = None
        if path.suffix.lower() == '.pdf':
            import fitz
            with fitz.open(path) as doc:
                number = source.get('pdf_page')
                if not isinstance(number, int) or isinstance(number, bool) or not 1 <= number <= len(doc):
                    errors.append(f'{source.get("id")}: physical PDF page missing/out of range')
                elif doc.needs_pass:
                    errors.append(f'{source.get("id")}: source PDF encrypted')
                else:
                    text = doc[number - 1].get_text()
                    record['pdf_page'] = number
        elif path.suffix.lower() in ['.txt', '.md', '.csv', '.json']:
            text = path.read_text(encoding='utf-8')
        else:
            warnings.append(f'{source.get("id")}: literal anchor validation unsupported for this format; inspect it directly')
        excerpt = source.get('excerpt')
        if excerpt and text is not None:
            record['excerpt_found'] = normalized(excerpt) in normalized(text)
            if not record['excerpt_found']:
                errors.append(f'{source.get("id")}: excerpt absent on specified source page')
        elif not excerpt:
            warnings.append(f'{source.get("id")}: no literal excerpt; needs direct visual/content verification')
        checked.append(record)
    slide_numbers = []
    for slide in ledger.get('slides', []):
        slide_numbers.append(slide.get('slide'))
        claims = slide.get('claims', [])
        if not claims and not slide.get('non_substantive'):
            errors.append(f'Slide {slide.get("slide")}: no claim-level evidence mapping')
        for claim in claims:
            refs = claim.get('source_ids', [])
            if not refs:
                errors.append(f'Slide {slide.get("slide")}: claim lacks evidence IDs')
            for ref in refs:
                if ref not in identifiers:
                    errors.append(f'Slide {slide.get("slide")}: unknown evidence ID {ref}')
    if not slide_numbers:
        errors.append('No slide evidence map supplied')
    if len(set(slide_numbers)) != len(slide_numbers):
        errors.append('Duplicate slide evidence entries')
    write_json(args.out, {'mechanical_status': 'failed' if errors else 'passed', 'errors': errors,
                         'warnings': warnings, 'sources': checked, 'scientific_entailment': 'not_checked'})
    return bool(errors)


def inspect_deck(args):
    from pptx import Presentation
    errors, warnings, slides = [], [], []
    with zipfile.ZipFile(args.input) as package:
        bad = package.testzip()
        if bad:
            errors.append('Corrupted ZIP member: ' + bad)
    prs = Presentation(args.input)
    mapped = {}
    if args.ledger:
        ledger = json.loads(Path(args.ledger).read_text(encoding='utf-8'))
        mapped = {x['slide']: x for x in ledger.get('slides', [])}
        invalid = [x for x in mapped if not isinstance(x, int) or not 1 <= x <= len(prs.slides)]
        if invalid:
            errors.append('Evidence map includes non-existent slide numbers: ' + str(invalid))
    def count_shape(shape, count, number, grouped=False):
        from pptx.enum.shapes import MSO_SHAPE_TYPE
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            for child in shape.shapes:
                count_shape(child, count, number, True)
            return
        if shape.has_text_frame and shape.text.strip():
            count['text'] += 1
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    if run.font.size is not None and run.font.size.pt < 12:
                        warnings.append(f'Slide {number} shape {shape.shape_id}: explicit font <12pt; assess template/footer context')
        if shape.has_table:
            count['tables'] += 1
        if shape.has_chart:
            count['charts'] += 1
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            count['pictures'] += 1
        # Group children have transformed coordinates: skip naive global bounds.
        if not grouped and (shape.left < -9144 or shape.top < -9144 or shape.left + shape.width > prs.slide_width + 9144 or shape.top + shape.height > prs.slide_height + 9144):
            warnings.append(f'Slide {number} shape {shape.shape_id}: coarse bounds extend outside slide; review actual render')
    for number, slide in enumerate(prs.slides, 1):
        count = {'text': 0, 'tables': 0, 'charts': 0, 'pictures': 0}
        for shape in slide.shapes:
            count_shape(shape, count, number)
        notes = slide.notes_slide.notes_text_frame.text if slide.has_notes_slide and slide.notes_slide.notes_text_frame else ''
        if not notes.strip():
            warnings.append(f'Slide {number}: no speaker notes')
        if count['text'] + count['tables'] + count['charts'] == 0:
            warnings.append(f'Slide {number}: no detected native editable content')
        if args.ledger and number not in mapped:
            errors.append(f'Slide {number}: no evidence-map entry')
        slides.append({'slide': number, 'layout': slide.slide_layout.name, 'native_objects': count, 'has_notes': bool(notes.strip())})
    if not len(prs.slides):
        errors.append('Deck contains no slides')
    write_json(args.out, {'mechanical_status': 'failed' if errors else 'passed', 'slide_count': len(prs.slides),
                         'master_count': len(prs.slide_masters), 'errors': errors, 'warnings': warnings,
                         'slides': slides, 'visual_fit': 'not_checked', 'scientific_accuracy': 'not_checked', 'target_playback': 'not_checked'})
    return bool(errors)


def render(args):
    out = Path(args.out).resolve()
    out.mkdir(parents=True, exist_ok=True)
    input_path = Path(args.input).resolve()
    record = {'status': 'partial', 'input': str(input_path), 'versions': versions(),
              'visual_review': 'not_performed_by_script', 'errors': [], 'outputs': []}
    write_json(out / 'render.json', record)
    binary = shutil.which('libreoffice') or shutil.which('soffice')
    if not binary:
        record['errors'].append('No actual LibreOffice/soffice renderer available; use another supplied renderer or return partial')
        write_json(out / 'render.json', record)
        return 1
    start = time.monotonic()
    try:
        version = subprocess.run([binary, '--version'], capture_output=True, text=True, timeout=15)
        record['renderer_version'] = (version.stdout + version.stderr).strip()
        # A fresh directory prevents stale PDFs from masquerading as new renders.
        with tempfile.TemporaryDirectory(prefix='evidence-render-') as temp:
            temp_path = Path(temp)
            destination = temp_path / 'export'
            destination.mkdir()
            command = [binary, '-env:UserInstallation=' + (temp_path / 'profile').as_uri(), '--headless',
                       '--convert-to', 'pdf', '--outdir', str(destination), str(input_path)]
            record['command'] = command
            completed = subprocess.run(command, capture_output=True, text=True, timeout=args.timeout)
            record['returncode'] = completed.returncode
            record['stdout'] = completed.stdout[-8000:]
            record['stderr'] = completed.stderr[-8000:]
            rendered = destination / (input_path.stem + '.pdf')
            if completed.returncode or not rendered.is_file() or rendered.stat().st_size == 0:
                raise RuntimeError('Actual conversion did not produce a nonempty PDF')
            final_pdf = out / (input_path.stem + '.pdf')
            shutil.copy2(rendered, final_pdf)
        import fitz
        from PIL import Image, ImageDraw
        thumbnails = []
        with fitz.open(final_pdf) as document:
            record['pdf_page_count'] = len(document)
            for number, page in enumerate(document, 1):
                target = out / f'slide-{number:03d}.png'
                page.get_pixmap(matrix=fitz.Matrix(1.4, 1.4), alpha=False).save(target)
                record['outputs'].append(str(target))
                with Image.open(target) as image:
                    image.thumbnail((400, 225))
                    thumb = Image.new('RGB', (420, 255), 'white')
                    thumb.paste(image, ((420 - image.width) // 2, 22))
                    ImageDraw.Draw(thumb).text((8, 4), f'Slide {number}', fill='black')
                    thumbnails.append(thumb)
        for batch_start in range(0, len(thumbnails), 12):
            batch = thumbnails[batch_start:batch_start + 12]
            cols = min(3, len(batch))
            sheet = Image.new('RGB', (420 * cols, 255 * ((len(batch) + cols - 1) // cols)), '#d9d9d9')
            for i, thumb in enumerate(batch):
                sheet.paste(thumb, (i % cols * 420, i // cols * 255))
            target = out / f'contact-{batch_start // 12 + 1:02d}.png'
            sheet.save(target)
            record['outputs'].append(str(target))
        from pptx import Presentation
        record['pptx_slide_count'] = len(Presentation(str(input_path)).slides)
        if record['pptx_slide_count'] != record['pdf_page_count']:
            raise RuntimeError('Actual PDF page count differs from PPTX slide count')
        record['outputs'].insert(0, str(final_pdf))
        record['status'] = 'rendered_unreviewed'
    except Exception as error:
        record['errors'].append(str(error))
    record['elapsed_seconds'] = round(time.monotonic() - start, 3)
    write_json(out / 'render.json', record)
    return record['status'] != 'rendered_unreviewed'


def package_check(args):
    root = Path(args.root).resolve()
    errors, files, total = [], [], 0
    for path in sorted(root.rglob('*')):
        if '.git' in path.relative_to(root).parts:
            continue
        if path.is_symlink():
            errors.append(str(path.relative_to(root)) + ': symlink forbidden')
            continue
        if not path.is_file():
            continue
        content = path.read_bytes()
        total += len(content)
        files.append({'path': str(path.relative_to(root)), 'bytes': len(content)})
        try:
            content.decode('utf-8')
            if b'\x00' in content:
                errors.append(str(path.relative_to(root)) + ': NUL/binary content')
        except UnicodeDecodeError:
            errors.append(str(path.relative_to(root)) + ': not UTF-8 text')
        if len(content) > 500_000:
            errors.append(str(path.relative_to(root)) + ': exceeds 500KB')
    if len(files) > 50 or total > 2_000_000:
        errors.append('Exceeded 50-file or 2MB resource limit')
    skill = root / 'SKILL.md'
    if not skill.is_file():
        errors.append('Root SKILL.md missing')
    else:
        content = skill.read_text(encoding='utf-8')
        if not content.startswith('---\n') or not re.search(r'^name: [a-z0-9-]+$', content, re.M) or not re.search(r'^description: .+', content, re.M):
            errors.append('Required frontmatter is absent/malformed; also run Skill-authoring validator')
    print(json.dumps({'status': 'failed' if errors else 'passed', 'file_count': len(files),
                      'total_bytes': total, 'errors': errors, 'files': files}, ensure_ascii=False, indent=2))
    return bool(errors)


def check_result(args):
    case = json.loads(Path(args.case).read_text(encoding='utf-8'))
    result = json.loads(Path(args.result).read_text(encoding='utf-8'))
    errors, warnings = [], []
    if not isinstance(result, dict):
        raise ValueError('Final result must be a JSON object')
    if 'case_id' not in case:
        warnings.append('case.json has no top-level case_id: read task for authoritative identifier')
    elif result.get('case_id') != case['case_id']:
        errors.append('case_id differs from original input')
    contract = case.get('output_contract')
    if isinstance(contract, dict):
        if isinstance(contract.get('properties'), dict):
            fields = contract['properties']
            required = contract.get('required', [])
        else:
            fields = contract
            required = list(fields)
        for key in required:
            if key not in result:
                errors.append(f'Original required output field missing: {key}')
        types = {'string': str, 'object': dict, 'array': list, 'boolean': bool, 'integer': int, 'number': (float, int)}
        for key, desc in fields.items():
            kind = desc.get('type') if isinstance(desc, dict) else desc
            if isinstance(kind, str) and kind in types and key in result:
                value = result[key]
                if not isinstance(value, types[kind]) or kind in ['integer', 'number'] and isinstance(value, bool):
                    errors.append(f'Original field has wrong top-level type: {key}')
    elif isinstance(contract, list) and all(isinstance(x, str) for x in contract):
        for key in contract:
            if key not in result:
                errors.append(f'Original required output field missing: {key}')
    else:
        warnings.append('No mechanically supported output_contract structure; inspect original task manually')
    write_json(args.out, {'mechanical_status': 'failed' if errors else 'passed', 'errors': errors, 'warnings': warnings,
                         'note': 'Checks top-level fields/types only. Review nested schema, scientific content and actual file paths separately.'})
    return bool(errors)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest='command', required=True)
    p = sub.add_parser('evidence'); p.add_argument('--ledger', required=True); p.add_argument('--out', required=True); p.set_defaults(func=evidence)
    p = sub.add_parser('deck'); p.add_argument('--input', required=True); p.add_argument('--ledger'); p.add_argument('--out', required=True); p.set_defaults(func=inspect_deck)
    p = sub.add_parser('render'); p.add_argument('--input', required=True); p.add_argument('--out', required=True)
    p.add_argument('--timeout', type=int, default=120); p.set_defaults(func=render)
    p = sub.add_parser('package'); p.add_argument('--root', required=True); p.set_defaults(func=package_check)
    p = sub.add_parser('result'); p.add_argument('--case', required=True); p.add_argument('--result', required=True)
    p.add_argument('--out', required=True); p.set_defaults(func=check_result)
    args = parser.parse_args()
    try:
        return int(bool(args.func(args)))
    except Exception as error:
        print(json.dumps({'status': 'execution_failed', 'error': str(error)}, ensure_ascii=False), file=sys.stderr)
        return 1


if __name__ == '__main__':
    sys.exit(main())
