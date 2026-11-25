#!/usr/bin/env python3
"""Local input inventories. No model calls or inferred scientific answers."""
import argparse
import hashlib
import importlib.metadata
import json
import os
from pathlib import Path
import platform
import shutil
import subprocess
import sys
import zipfile
import xml.etree.ElementTree as ET


def write_json(path, value):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + '.tmp')
    temporary.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding='utf-8')
    temporary.replace(path)


def versions():
    result = {'python': platform.python_version()}
    for name in ['python-pptx', 'PyMuPDF', 'Pillow']:
        try:
            result[name] = importlib.metadata.version(name)
        except importlib.metadata.PackageNotFoundError:
            result[name] = None
    return result


def probe(args):
    case_path = Path(args.case).resolve()
    case = json.loads(case_path.read_text(encoding='utf-8'))
    roots = [case_path.parent, Path('/app'), Path.cwd() / 'environment']
    found = {}
    for root in roots:
        if not root.is_dir():
            continue
        for current, dirs, files in os.walk(root):
            relative = Path(current).relative_to(root)
            dirs[:] = [x for x in dirs if not x.startswith('.') and x not in ['work', 'output', 'node_modules', '__pycache__', 'submission']]
            if len(relative.parts) > 3:
                dirs[:] = []
            for name in files:
                path = Path(current) / name
                if path.suffix.lower() in ['.pdf', '.pptx', '.potx', '.docx', '.json', '.yaml', '.yml', '.md']:
                    found[str(path.resolve())] = path.stat().st_size
            if len(found) > 1000:
                break
    fonts = []
    font_tool = shutil.which('fc-list')
    if font_tool:
        try:
            fonts = sorted(set(subprocess.run([font_tool, '--format', '%{family}\n'], capture_output=True, text=True, timeout=15).stdout.splitlines()))
        except subprocess.SubprocessError:
            pass
    write_json(args.out, {'case_path': str(case_path), 'case': case, 'files': found,
                         'versions': versions(), 'tools': {x: shutil.which(x) for x in ['libreoffice', 'soffice', 'pdftoppm', 'tesseract']},
                         'available_font_families': fonts, 'note': 'Discovery only; read original task/resources and files.'})


def pdf(args):
    import fitz
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    with fitz.open(args.input) as doc:
        if doc.needs_pass:
            raise ValueError('PDF requires an authorized password; no content extracted')
        pages, text = [], []
        for index, page in enumerate(doc):
            content = page.get_text('text')
            images = [{'xref': x[0], 'width': x[2], 'height': x[3]} for x in page.get_images(full=True)]
            pages.append({'pdf_page': index + 1, 'printed_label': page.get_label(),
                          'width_pt': page.rect.width, 'height_pt': page.rect.height,
                          'text': content, 'images': images, 'possible_scan': len(content.strip()) < 40})
            text.append(f'\n=== PHYSICAL PDF PAGE {index + 1}, LABEL {page.get_label()} ===\n{content}')
        write_json(out / 'pages.json', {'input': str(Path(args.input).resolve()), 'page_count': len(doc), 'pages': pages})
        (out / 'pages.txt').write_text('\n'.join(text), encoding='utf-8')


def page_image(args):
    import fitz
    with fitz.open(args.input) as doc:
        if doc.needs_pass:
            raise ValueError('PDF requires an authorized password')
        if not 1 <= args.page <= len(doc):
            raise ValueError('Page is outside the actual PDF')
        page = doc[args.page - 1]
        clip = fitz.Rect(args.bbox) if args.bbox else page.rect
        if clip.is_empty or clip.is_infinite or not page.rect.contains(clip):
            raise ValueError('Crop must be a positive rectangle inside the actual PDF page')
        output = Path(args.output)
        output.parent.mkdir(parents=True, exist_ok=True)
        page.get_pixmap(matrix=fitz.Matrix(args.scale, args.scale), clip=clip, alpha=False).save(output)


def shape_info(shape):
    item = {'id': shape.shape_id, 'name': shape.name, 'type': str(shape.shape_type),
            'box_inches': [round(float(v) / 914400, 4) for v in [shape.left, shape.top, shape.width, shape.height]]}
    if shape.is_placeholder:
        item['placeholder_idx'] = shape.placeholder_format.idx
        item['placeholder_type'] = str(shape.placeholder_format.type)
    if shape.has_text_frame:
        item['text'] = shape.text
        item['paragraphs'] = [{'level': p.level, 'paragraph_font_size': p.font.size.pt if p.font.size else None,
                               'paragraph_font_name': p.font.name, 'font_sizes': [r.font.size.pt if r.font.size else None for r in p.runs],
                               'font_names': [r.font.name for r in p.runs]} for p in shape.text_frame.paragraphs]
    if shape.has_table:
        item['table'] = [[c.text for c in row.cells] for row in shape.table.rows]
    return item


def deck(args):
    from pptx import Presentation
    prs = Presentation(args.input)
    masters = []
    for i, master in enumerate(prs.slide_masters):
        masters.append({'master': i, 'name': master.name, 'layouts': [
            {'layout': j, 'name': layout.name, 'placeholders': [shape_info(x) for x in layout.placeholders]}
            for j, layout in enumerate(master.slide_layouts)]})
    slides = [{'slide': i + 1, 'layout': s.slide_layout.name, 'shapes': [shape_info(x) for x in s.shapes],
               'notes': s.notes_slide.notes_text_frame.text if s.has_notes_slide and s.notes_slide.notes_text_frame else ''}
              for i, s in enumerate(prs.slides)]
    theme_info = []
    with zipfile.ZipFile(args.input) as package:
        ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
        for name in package.namelist():
            if name.startswith('ppt/theme/theme') and name.endswith('.xml'):
                theme = ET.fromstring(package.read(name))
                theme_info.append({'part': name, 'fonts': [node.attrib for node in theme.findall('.//a:fontScheme//*', ns)],
                                   'colors': [node.attrib for node in theme.findall('.//a:clrScheme//*', ns)]})
    write_json(args.out, {'input': str(Path(args.input).resolve()),
                         'size_inches': [prs.slide_width / 914400, prs.slide_height / 914400],
                         'masters': masters, 'themes': theme_info, 'slides': slides,
                         'note': 'Null font sizes/names may inherit from layout/master/theme; do not treat as missing style.'})


def manifest(args):
    items = []
    for name in args.paths:
        path = Path(name)
        if not path.is_file():
            items.append({'path': str(path), 'exists': False})
            continue
        digest = hashlib.sha256()
        with path.open('rb') as f:
            for data in iter(lambda: f.read(1024 * 1024), b''):
                digest.update(data)
        items.append({'path': str(path.resolve()), 'exists': True, 'bytes': path.stat().st_size, 'sha256': digest.hexdigest()})
    write_json(args.out, {'versions': versions(), 'files': items})


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest='command', required=True)
    p = sub.add_parser('probe'); p.add_argument('--case', required=True); p.add_argument('--out', required=True); p.set_defaults(func=probe)
    for name, function in [('pdf', pdf), ('deck', deck)]:
        p = sub.add_parser(name); p.add_argument('--input', required=True); p.add_argument('--out', required=True); p.set_defaults(func=function)
    p = sub.add_parser('page'); p.add_argument('--input', required=True); p.add_argument('--page', type=int, required=True)
    p.add_argument('--bbox', type=float, nargs=4); p.add_argument('--scale', type=float, default=2)
    p.add_argument('--output', required=True); p.set_defaults(func=page_image)
    p = sub.add_parser('manifest'); p.add_argument('--paths', nargs='+', required=True); p.add_argument('--out', required=True); p.set_defaults(func=manifest)
    args = parser.parse_args()
    try:
        args.func(args)
    except Exception as error:
        print(json.dumps({'status': 'execution_failed', 'error': str(error)}, ensure_ascii=False), file=sys.stderr)
        return 1
    return 0


if __name__ == '__main__':
    sys.exit(main())
