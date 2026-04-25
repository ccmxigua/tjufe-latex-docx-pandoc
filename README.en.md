# TJUFE LaTeX-to-DOCX Thesis Toolkit

This repository is an open-source **LaTeX-to-DOCX conversion toolkit for Chinese academic thesis formatting**. It converts LaTeX thesis sources into editable Word/WPS `.docx` drafts that follow institution-specific thesis layout requirements.

The current implementation targets the management graduate thesis formatting rules of Tianjin University of Finance and Economics (TJUFE), but the engineering problem is broader: many Chinese universities require final thesis submissions in DOCX format, while researchers often prefer writing in LaTeX. This project provides a reproducible toolchain to bridge that gap and reduce manual reformatting.

> The Chinese README is the primary user guide because the target rules and users are Chinese. This English file is provided for international reviewers, maintainers, and open-source program evaluators.

## What it does

- Converts LaTeX thesis files to Word DOCX.
- Generates or normalizes thesis structures such as cover pages, title pages, Chinese/English abstracts, table of contents, chapters, appendices, acknowledgements, bibliography, figures, tables, equations, and footnotes.
- Provides a generated `reference.docx` as the Word style baseline.
- Uses a Pandoc Lua filter to normalize thesis structure during conversion.
- Uses Python-based OOXML post-processing to fix layout features that Pandoc alone does not handle reliably.
- Includes public sample inputs and generated sample output for reproducibility.
- Includes implementation notes comparing the official formatting specification with the current script behavior.

## Why this matters

Chinese academic thesis submission workflows often require final documents to satisfy detailed Word/WPS layout rules, even when the author writes the thesis in LaTeX. Manual conversion is slow, error-prone, and hard to audit. This project turns that workflow into a scripted conversion pipeline so that formatting rules can be tested, reviewed, and improved over time.

The project is local to TJUFE in its first target specification, but the technical approach is reusable for other Chinese university thesis templates and other institution-specific academic document workflows.

## Technical architecture

```text
LaTeX source
  ↓
scripts/preprocess_tjufe_tex.py
  ↓
pandoc + filters/tjufe-thesis.lua + reference.docx
  ↓
intermediate raw.docx
  ↓
scripts/postprocess_tjufe_docx.py
  ↓
final output.docx
  ↓
built-in audit checks from convert_tjufe.sh
```

Main components:

- `convert_tjufe.sh`
  - Main command-line entry point.
  - Parses source, output path, metadata, and extra Pandoc options.
  - Runs preprocessing, Pandoc conversion, Lua filtering, DOCX post-processing, and audit checks.

- `scripts/preprocess_tjufe_tex.py`
  - Light preprocessing before Pandoc.
  - Reduces conversion failures caused by complex LaTeX patterns.

- `filters/tjufe-thesis.lua`
  - Pandoc Lua filter.
  - Handles thesis structure, headings, abstracts, keywords, figures, tables, equations, and cross-reference markers.

- `scripts/build_reference_docx.py`
  - Generates `reference.docx`.
  - Defines baseline Word styles, page settings, headings, abstracts, footnotes, bibliography styles, and other layout defaults.

- `scripts/postprocess_tjufe_docx.py`
  - Performs OOXML-level corrections after Pandoc generates DOCX.
  - Handles page sections, headers and footers, page numbering, table-of-contents fields, captions, three-line tables, equation numbering, footnote numbering, cross-references, paragraph styles, bibliography layout, and appendix numbering.

## Typical use

```bash
python3 scripts/build_reference_docx.py
./convert_tjufe.sh samples/sample-thesis.tex out/sample.docx samples/sample-metadata.yaml
```

With bibliography and CSL support:

```bash
./convert_tjufe.sh thesis.tex out/thesis.docx thesis-metadata.yaml \
  --bibliography refs.bib \
  --citeproc \
  --csl /path/to/gb7714-2015-numeric.csl
```

## Dependencies

Required:

- `bash`
- `pandoc`
- `python3`
- Python packages: `python-docx`, `lxml`

Optional:

- Microsoft Word or WPS for final manual inspection and updating document fields such as the table of contents.
- A LaTeX distribution if you also need PDF generation. This toolkit focuses on LaTeX-to-DOCX conversion through Pandoc.

## Repository contents

```text
.
├── convert_tjufe.sh
├── reference.docx
├── filters/
│   └── tjufe-thesis.lua
├── scripts/
│   ├── build_reference_docx.py
│   ├── preprocess_tjufe_tex.py
│   └── postprocess_tjufe_docx.py
├── samples/
│   ├── sample-thesis.tex
│   ├── sample-metadata.yaml
│   └── sample.docx
├── assets/
│   └── sample.docx.png
├── docs/
│   ├── tjufe-management-thesis-format-spec.docx
│   └── spec-vs-implementation-management-full-clause.md
├── LICENSE
├── README.md
└── README.en.md
```

## Scope and limitations

This project aims to produce a strong DOCX draft that is suitable for final checking in Word/WPS. It does not claim to eliminate all manual review. The following items should still be inspected before official submission:

- Exact positioning of cover and title pages.
- Table-of-contents line breaks and leader dots after field updates.
- Complex table borders and page breaks.
- Complex equation layout.
- Continued captions for figures and tables across pages.
- Edge cases in mixed Chinese/English bibliography formatting.
- School-specific declaration or authorization pages that may vary by year or department.

## Open-source maintenance value

This is a document-engineering project with real maintenance challenges:

- It combines Pandoc, Lua filters, Python scripts, and OOXML manipulation.
- It must track institution-specific formatting rules.
- It needs regression tests and public sample documents.
- It benefits from code review, refactoring, documentation, and automated checks.
- It addresses a practical academic workflow for Chinese LaTeX users who must submit DOCX files.

Contributions that improve portability, tests, style coverage, error reporting, or support for additional university templates are welcome.

## License

See [`LICENSE`](LICENSE).
