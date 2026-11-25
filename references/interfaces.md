# Helper interfaces

## Source ledger

Write JSON with `sources` and `slides`. Every source needs unique `id` and actual absolute file `path`. PDF sources also REQUIRE a 1-based physical `pdf_page`. Optional fields include `printed_label`, `section`, `figure`, `panel`, exact `excerpt`, and `ownership`. Every substantive slide needs a positive 1-based `slide`, a `purpose`, and a REQUIRED `claims` array. Every claim needs `text` and nonempty `source_ids`. Optional slide fields include `uncertainty` and `visuals`, whose sources should also be explicitly identified. Use exact short excerpts where searchable text is available. For image-only evidence use a locator and record actual visual inspection rather than inventing an excerpt. Include every slide in the ledger, marking only genuinely administrative slides as non_substantive.

Use the exact machine keys `pdf_page` and `excerpt` (not `page` or `quote`). Every substantive slide MUST have a `claims` array, and every claim MUST have `source_ids`. An interface example, not scientific output:

```json
{"sources":[{"id":"S1","path":"/input/paper.pdf","pdf_page":3,"printed_label":"125","section":"Results","figure":"Fig. 2a","excerpt":"An actual short excerpt read from physical PDF page 3"}],"slides":[{"slide":1,"purpose":"Explain the main result","claims":[{"text":"An actual source-supported result","source_ids":["S1"]}]}]}
```

For an image-only source omit `excerpt`, record `figure`/`pdf_page` and actual visual-review notes. The audit warns that direct verification is needed. A cover or closing slide with no substantive claim can explicitly use `non_substantive: true` and `claims: []`.

`audit.py evidence` validates existing paths, PDF page ranges, excerpt occurrence and source IDs. It cannot establish scientific entailment. Run it, then perform the scientific review.

## Blueprint

`deck.py --blueprint FILE --log FILE` reads:

```
{
  "mode": "new",
  "template": "/input/template.pptx",
  "output": "/app/output/deck.pptx",
  "slides": [
    {
      "master": 0, "layout": 1,
      "title_placeholder": 0, "title": "A source-supported title",
      "body_placeholder": 1, "bullets": ["A source-supported point"],
      "notes": "Explain the result and transition to the next slide.",
      "source_ids": ["S1"], "source_note": "Paper, PDF p. 3, Fig. 2a"
    }
  ]
}
```

This is an interface illustration, NOT a formal answer. Agent must write actual case-specific content after reading the actual source. Set explicit placeholder indices from the template inspection. Bullets may be strings or `{text, level}` objects. Text inherits template style; optional `font_size`, `font_name` are explicit overrides to log. `title` and `bullets` require their corresponding placeholder index.

Optional native objects: `textboxes` with `text`, `box`; `pictures` with `path`, `box`; `tables` with rectangular `rows`, `box`; `charts` with `type` (column/bar/line/scatter), `categories`, `series`, `box`. Series use `name`/`values`; scatter uses `name`/`x`/`y`. Boxes are [left, top, width, height] in inches, measured against actual template geometry. Pictures are fitted without stretching. Native charts require exact actual source values; include units/conditions in visible labels or notes. `chart_title` and `y_title` are optional chart properties. Use another supported tool for complex chart types.

For `mode: adapt`, give `input` with actual starting deck. Each slides entry must specify 1-based `slide`. The helper does not add/delete/reorder slides in this mode. It preserves untouched content. Use explicit placeholder targets, or `replace_shapes` with `{shape_id, text}` for existing text shapes. New picture/table/chart/textbox entries add objects. `delete_shapes` requires explicit shape IDs and a per-slide `reason`; use only if the task authorizes removal. Log replacement/removal reasons. Existing notes are appended to by default; `replace_notes: true` requires case authorization.

For `mode: new`, use template masters/layouts and remove template sample slides while retaining reusable masters. Every slide has explicit master/layout; no default layout guessing. Preserve the original template file.

Python-pptx does not activate latent layout footer/date/slide-number placeholders automatically. If the actual task/template requires them, add top-level or per-slide `footer` with `text`, `date` and/or `slide_number: true`. The adapter clones actual matching layout placeholders and preserves their inherited style. No footer is invented by default. Use another editor for footer controls this helper cannot preserve.

## Structured result

Preserve all ORIGINAL output_contract fields/types and exact case_id first. Suggested supplemental fields, if compatible:

```
{
  "case_id":"actual input ID",
  "status":"partial",
  "artifacts":{"editable":"/app/output/deck.pptx","evidence":"/app/output/evidence.json"},
  "sources":[], "uncertainty":[],
  "issues":[{"type":"unable_to_determine","detail":"Actual unresolved item"}],
  "checks":{"structural":"not_run","render":"not_run","visual_review":"not_run"},
  "changes":[], "versions":{}
}
```

Use the status enum required by the original case. State inspected pages/observations for visual review. No helper claims scientific approval or model improvement. Manifest hashes and dependency versions support process reproduction; do not promise byte-identical PPTX files across machines.

## Commands

* `prepare.py probe --case FILE --out FILE`: case, files, tool availability.
* `prepare.py pdf --input FILE --out DIR`: page text/image index, no pre-extracted answers.
* `prepare.py page --input FILE --page N [--bbox x0 y0 x1 y1] --output FILE`: actual page/crop PNG.
* `prepare.py deck --input FILE --out FILE`: master/layout/placeholder and slide inventory.
* `prepare.py manifest --paths FILE ... --out FILE`: hashes and actual dependency versions.
* `audit.py evidence --ledger FILE --out FILE`: mechanical source checks.
* `audit.py deck --input FILE [--ledger FILE] --out FILE`: native content/coarse geometry checks.
* `audit.py render --input FILE --out DIR [--timeout 120]`: actual conversion, PNGs, contact sheets and honest render.json.
* `audit.py package --root DIR`: competition text-only resource limits.
* `audit.py result --case FILE --result FILE --out FILE`: exact case_id and original top-level contract fields/types (nested schema/semantics need separate review).
