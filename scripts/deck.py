#!/usr/bin/env python3
"""Optional native PPTX adapter. Read a source-grounded blueprint; preserve template parts."""
import argparse
from copy import deepcopy
import json
from pathlib import Path
import sys

from prepare import write_json


def set_text(shape, content, font_size=None, font_name=None):
    from pptx.util import Pt
    if not shape.has_text_frame:
        raise ValueError(f'Shape {shape.shape_id} is not an editable text shape')
    tf = shape.text_frame
    # Retain paragraph properties and the first run style where present.
    initial = tf.paragraphs[0]
    run_style = deepcopy(initial.runs[0]._r.rPr) if initial.runs and initial.runs[0]._r.rPr is not None else None
    tf.clear()
    entries = content if isinstance(content, list) else [content]
    for i, entry in enumerate(entries or ['']):
        entry = {'text': entry} if isinstance(entry, str) else entry
        paragraph = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        paragraph.level = int(entry.get('level', 0))
        run = paragraph.add_run()
        run.text = str(entry.get('text', ''))
        if run_style is not None:
            run._r.insert(0, deepcopy(run_style))
        if font_size is not None:
            run.font.size = Pt(font_size)
        if font_name:
            run.font.name = font_name


def placeholder(slide, idx):
    try:
        return slide.placeholders[int(idx)]
    except KeyError:
        raise ValueError(f'Actual slide/layout has no placeholder idx={idx}')


def dimensions(item):
    from pptx.util import Inches
    box = item['box']
    if len(box) != 4 or box[2] <= 0 or box[3] <= 0:
        raise ValueError('box must contain [left, top, positive width, positive height]')
    return [Inches(float(x)) for x in box]


def add_objects(slide, spec, actions):
    from pptx.chart.data import CategoryChartData, XyChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.util import Pt
    for item in spec.get('textboxes', []):
        shape = slide.shapes.add_textbox(*dimensions(item))
        set_text(shape, item['text'], item.get('font_size'), item.get('font_name'))
        actions.append({'action': 'add_textbox', 'shape_id': shape.shape_id})
    for item in spec.get('pictures', []):
        from PIL import Image
        path = Path(item['path'])
        with Image.open(path) as image:
            width, height = image.size
        left, top, available_w, available_h = dimensions(item)
        ratio = min(available_w / width, available_h / height)
        final_w, final_h = int(width * ratio), int(height * ratio)
        shape = slide.shapes.add_picture(str(path), int(left + (available_w - final_w) / 2),
                                         int(top + (available_h - final_h) / 2), final_w, final_h)
        actions.append({'action': 'add_source_picture', 'shape_id': shape.shape_id, 'path': str(path)})
    for item in spec.get('tables', []):
        rows = item['rows']
        if not rows or not rows[0] or any(len(row) != len(rows[0]) for row in rows):
            raise ValueError('Native table requires nonempty rectangular rows')
        shape = slide.shapes.add_table(len(rows), len(rows[0]), *dimensions(item))
        for i, row in enumerate(rows):
            for j, value in enumerate(row):
                cell = shape.table.cell(i, j)
                cell.text = str(value)
                for paragraph in cell.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if item.get('font_size') is not None:
                            run.font.size = Pt(item['font_size'])
                        if item.get('font_name'):
                            run.font.name = item['font_name']
        actions.append({'action': 'add_native_table', 'shape_id': shape.shape_id})
    for item in spec.get('charts', []):
        kind = item['type']
        types = {'column': XL_CHART_TYPE.COLUMN_CLUSTERED, 'bar': XL_CHART_TYPE.BAR_CLUSTERED,
                 'line': XL_CHART_TYPE.LINE, 'scatter': XL_CHART_TYPE.XY_SCATTER}
        if kind not in types:
            raise ValueError(f'Unsupported chart type: {kind}')
        if kind == 'scatter':
            data = XyChartData()
            for series in item['series']:
                if len(series['x']) != len(series['y']):
                    raise ValueError('Scatter x/y lengths differ')
                target = data.add_series(series['name'])
                for x, y in zip(series['x'], series['y']):
                    target.add_data_point(x, y)
        else:
            data = CategoryChartData()
            data.categories = item['categories']
            for series in item['series']:
                if len(series['values']) != len(item['categories']):
                    raise ValueError('Category/series lengths differ')
                data.add_series(series['name'], series['values'])
        shape = slide.shapes.add_chart(types[kind], *dimensions(item), data)
        chart = shape.chart
        if item.get('chart_title'):
            chart.has_title = True
            chart.chart_title.text_frame.text = item['chart_title']
        if item.get('y_title'):
            chart.value_axis.has_title = True
            chart.value_axis.axis_title.text_frame.text = item['y_title']
        actions.append({'action': 'add_native_chart', 'shape_id': shape.shape_id})


def apply_footer(slide, rules, number, actions):
    from pptx.oxml.xmlchemy import OxmlElement
    from pptx.enum.shapes import PP_PLACEHOLDER
    kinds = {'text': PP_PLACEHOLDER.FOOTER, 'date': PP_PLACEHOLDER.DATE, 'slide_number': PP_PLACEHOLDER.SLIDE_NUMBER}
    for key, kind in kinds.items():
        value = rules.get(key)
        if value is None or value is False:
            continue
        layout_ph = next((p for p in slide.slide_layout.placeholders if p.placeholder_format.type == kind), None)
        if layout_ph is None:
            raise ValueError(f'Requested {key} requires an actual matching template layout placeholder')
        idx = layout_ph.placeholder_format.idx
        if not any(p.placeholder_format.idx == idx for p in slide.placeholders):
            slide.shapes.clone_placeholder(layout_ph)
        target = slide.placeholders[idx]
        if key == 'slide_number':
            set_text(target, '')
            field = OxmlElement('a:fld')
            field.set('id', '{B8326B5A-9D07-4440-AC3D-55004CF09999}')
            field.set('type', 'slidenum')
            text = OxmlElement('a:t'); text.text = str(number); field.append(text)
            target.text_frame.paragraphs[0]._p.append(field)
        else:
            set_text(target, str(value))
        actions.append({'action': 'set_template_footer', 'field': key, 'placeholder_idx': idx})


def build(args):
    from pptx import Presentation
    bp = json.loads(Path(args.blueprint).read_text(encoding='utf-8'))
    mode = bp.get('mode')
    if mode not in ['new', 'adapt']:
        raise ValueError('mode must explicitly be new or adapt')
    source = Path(bp['template'] if mode == 'new' else bp['input']).resolve()
    output = Path(bp['output']).resolve()
    if source == output:
        raise ValueError('Never overwrite the original template/starting deck')
    if not bp.get('slides'):
        raise ValueError('Blueprint needs at least one actual slide specification')
    prs = Presentation(str(source))
    if mode == 'new':
        for slide_id in list(prs.slides._sldIdLst):
            prs.part.drop_rel(slide_id.rId)
            prs.slides._sldIdLst.remove(slide_id)
    output.parent.mkdir(parents=True, exist_ok=True)
    log = {'mode': mode, 'source': str(source), 'output': str(output), 'changes': [], 'status': 'partial'}
    write_json(args.log, log)
    for spec in bp['slides']:
        if mode == 'new':
            master = prs.slide_masters[int(spec['master'])]
            layout = master.slide_layouts[int(spec['layout'])]
            slide = prs.slides.add_slide(layout)
            number = len(prs.slides)
        else:
            number = int(spec['slide'])
            if not 1 <= number <= len(prs.slides):
                raise ValueError(f'Adapt target slide {number} is outside the actual deck')
            slide = prs.slides[number - 1]
        actions = []
        if 'footer' in spec or 'footer' in bp:
            apply_footer(slide, spec.get('footer', bp.get('footer')), number, actions)
        for content_key, idx_key in [('title', 'title_placeholder'), ('bullets', 'body_placeholder')]:
            if content_key in spec:
                if idx_key not in spec:
                    raise ValueError(f'{content_key} requires explicit {idx_key}')
                target = placeholder(slide, spec[idx_key])
                set_text(target, spec[content_key], spec.get('font_size'), spec.get('font_name'))
                actions.append({'action': 'replace_placeholder_text', 'placeholder_idx': spec[idx_key], 'shape_id': target.shape_id})
        for item in spec.get('replace_shapes', []):
            target = next((s for s in slide.shapes if s.shape_id == int(item['shape_id'])), None)
            if target is None:
                raise ValueError(f'Target shape {item["shape_id"]} does not exist')
            set_text(target, item['text'], item.get('font_size'), item.get('font_name'))
            actions.append({'action': 'replace_shape_text', 'shape_id': target.shape_id})
        if spec.get('delete_shapes') and not spec.get('reason'):
            raise ValueError('Deleting shapes requires a recorded reason and task authorization')
        for shape_id in spec.get('delete_shapes', []):
            target = next((s for s in slide.shapes if s.shape_id == int(shape_id)), None)
            if target is None:
                raise ValueError(f'Delete target {shape_id} does not exist')
            # Drop relationships used exclusively by the deleted shape to avoid orphan media.
            element = target._element
            for rid in set(element.xpath('.//@r:embed') + element.xpath('.//@r:link')):
                uses = slide._element.xpath(f'.//*[@r:embed="{rid}" or @r:link="{rid}"]')
                if len(uses) == 1:
                    slide.part.drop_rel(rid)
            element.getparent().remove(element)
            actions.append({'action': 'delete_shape', 'shape_id': shape_id, 'reason': spec['reason']})
        add_objects(slide, spec, actions)
        notes = spec.get('notes', '')
        if spec.get('source_note'):
            notes += '\nSources: ' + spec['source_note']
        if spec.get('source_ids'):
            notes += '\nEvidence IDs: ' + ', '.join(spec['source_ids'])
        if notes:
            tf = slide.notes_slide.notes_text_frame
            if tf is None:
                raise ValueError('Notes master has no notes text placeholder; use another supported editor')
            previous = tf.text if mode == 'adapt' and not spec.get('replace_notes') else ''
            tf.text = (previous + '\n\n' + notes).strip()
            actions.append({'action': 'append_notes' if previous else 'set_notes'})
        if spec.get('font_name') or spec.get('font_size'):
            actions.append({'action': 'explicit_font_override', 'font_name': spec.get('font_name'), 'font_size': spec.get('font_size')})
        log['changes'].append({'slide': number, 'reason': spec.get('reason', 'case blueprint'), 'actions': actions})
        tmp = output.with_suffix('.partial.pptx')
        prs.save(str(tmp))
        tmp.replace(output)
        write_json(args.log, log)
    log['status'] = 'built_unreviewed'
    log['slide_count'] = len(prs.slides)
    write_json(args.log, log)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--blueprint', required=True)
    parser.add_argument('--log', required=True)
    args = parser.parse_args()
    try:
        build(args)
    except Exception as error:
        print(json.dumps({'status': 'execution_failed', 'error': str(error)}, ensure_ascii=False), file=sys.stderr)
        return 1
    return 0


if __name__ == '__main__':
    sys.exit(main())
