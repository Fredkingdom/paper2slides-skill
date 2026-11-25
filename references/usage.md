# Usage and minimum example

Version 1.0.1. Use this Skill for actual source-to-PPT work, including targeted adaptation. It supplements the Agent's scientific reading and tool use. It does not autonomously understand a paper or certify facts. Chinese operator instructions are in `usage-zh.md`.

## Loading and invocation

Load this directory through the target Agent's Skill loader. The competition evaluator installs it at `/skills/weekly-skill`. Prompt:

> Use evidence-paper-slides to complete the actual /input/case.json and task.md. Read the original papers, use the supplied template/starting deck, preserve the original output_contract and save artifacts in /app/output.

Read SKILL.md first. Resolve script paths from its directory. No private files, credentials, remote services or formal-case answers are included.

## Dependencies

Prefer organizer-installed versions listed in environment/resources.yaml. Author development uses Python 3.11+, python-pptx 1.0.2 and PyMuPDF 1.26.6; Pillow is used for contact sheets (actual version in the development report). Runtime scripts do not require PyYAML. If packages are absent and installing is permitted: `python -m pip install python-pptx==1.0.2 PyMuPDF==1.26.6 Pillow`. Do not reinstall compatible packages or wait on unavailable network access. Actual rendering needs external libreoffice/soffice; Python packages alone do not provide it. Use another supplied renderer if available.

## Minimal example: illustrative, not a formal result

Input: actual paper.pdf, template.pptx and a task requesting six total slides for an eight-minute talk. Run:

```
python /skills/weekly-skill/scripts/prepare.py probe --case /input/case.json --out /app/work/probe.json
python /skills/weekly-skill/scripts/prepare.py pdf --input /input/paper.pdf --out /app/work/paper
python /skills/weekly-skill/scripts/prepare.py deck --input /input/template.pptx --out /app/work/template.json
```

Read/inspect the source and create actual ledger.json and blueprint.json following interfaces.md. Then run:

```
python /skills/weekly-skill/scripts/audit.py evidence --ledger /app/work/ledger.json --out /app/work/evidence-check.json
python /skills/weekly-skill/scripts/deck.py --blueprint /app/work/blueprint.json --log /app/output/changes.json
python /skills/weekly-skill/scripts/audit.py deck --input /app/output/deck.pptx --ledger /app/work/ledger.json --out /app/work/deck-check.json
python /skills/weekly-skill/scripts/audit.py render --input /app/output/deck.pptx --out /app/output/preview
```

Inspect real previews, repair the actual deck and rerender as necessary. Example output paths are deck.pptx, evidence.json, preview/deck.pdf, preview/slide-001.png, changes.json, manifest.json and result.json. These paths illustrate structure only; no prepared scientific output is supplied. Complete the actual case's original output fields, not just this list.

## Maintenance and limits

Run `python scripts/audit.py package --root .` for text-resource limits. Rerun generation, targeted adaptation, real rendering and error handling on clearly synthetic fixtures outside the submission after changes. Tests establish tool behavior, not official scores. Scientific entailment, visual quality and target-application behavior still require actual review. Render absence, unreadable source, missing template, unsupported edits and missing fonts must be reported honestly.

GitHub is optional extra credit. GITHUB_URL.txt remains blank until a real public work URL is provided. Do not invent a URL or claim its accessibility. Official cases, whiteboards and benchmark results come from the platform.
