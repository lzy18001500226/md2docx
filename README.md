# md2word

`md2word` is a template-governed Markdown-to-DOCX delivery pipeline. It uses
Pandoc for Markdown conversion, then applies a profile's Word rules and checks
the resulting OOXML package before delivery.

It is intended for reports whose format is part of the acceptance criteria, not
for a best-effort Markdown preview.

## What it verifies

- Reference-DOCX styles remain the formatting authority.
- Figure and table captions use stable fields such as `图2-1 标题`.
- Generated documents do not retain the `SEQ Chapter` fields that can leak into
  first-level headings as values such as `20`.
- Source images are embedded unchanged when present.
- The DOCX package passes ZIP-integrity and structural checks.
- On Windows, an optional Microsoft Word pass updates fields and exports PDF.

## Quick start

Prerequisites:

- Python 3.11 or newer
- Pandoc on `PATH`
- Microsoft Word is optional and required only for `--word` or `--pdf`

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -e ".[dev]"
python scripts\create_demo_template.py
md2word build examples\demo-report.md --profile profiles\public-demo.toml --out artifacts\demo.docx --overwrite
pytest
```

Add `--word` to update fields through Word before validation, and `--pdf
artifacts\demo.pdf` to export a PDF in the same pass.

## Profiles

A profile supplies the reference DOCX and caption labels. The public example is
at `profiles/public-demo.toml`; project-specific templates belong in private
repositories unless their redistribution rights are confirmed.

## Formula backends

The public pipeline supports Pandoc's Office Math output. Native MathType is a
separate, opt-in integration because it requires a licensed Windows Word and
MathType installation plus a formula bank the user is entitled to use. It is not
bundled in this repository.

## Release boundary

Do not add course reports, competition templates, proprietary Word templates,
private images, MathType OLE objects, licenses, or formula banks without an
explicit redistribution review. See `docs/release-boundary.md`.

This repository currently has no license selected. Choose and add a license
before publishing code for reuse.
