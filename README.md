# md2docx

Generate a styled DOCX from a Markdown file and a DOCX style template.

All visual style — fonts, spacing, margins, header, footer — is inherited from
the template. Body content is replaced entirely by the rendered Markdown.

## Requirements

- Python 3.9+
- [pandoc](https://pandoc.org/) — `brew install pandoc`

No Python packages are needed at runtime.

## Usage

```
python generate.py <content.md> <template.docx> <output.docx>
```

**Example:**

```
python generate.py test/content.md test/TEMPLATE.docx test/output.docx
```

### Markdown features

| Markdown | DOCX output |
|---|---|
| `# Heading 1` | Heading 1 style |
| `## Heading 2` | Heading 2 style |
| `### Heading 3` | Heading 3 style |
| `- item` | Real Word bullet (numPr) |
| Plain paragraph | Body text style |
| `**bold**` | Bold run |
| `*italic*` | Italic run |
| Blank line | Empty paragraph (preserves vertical spacing) |

Bullets are real Word list items (not dash characters), inheriting the
template's list styling.

## Repository structure

```
.
├── generate.py           # main script — the tool
├── requirements-dev.txt  # test dependencies (pytest, python-docx)
├── specs/
│   └── generate.md       # behavioural specification
├── tests/
│   ├── conftest.py       # shared fixtures
│   ├── test_generate.py  # test suite
│   └── fixtures/
│       ├── simple.md     # minimal markdown with bullets
│       └── headings.md   # markdown with heading levels
```

## Running tests

Install test dependencies once:

```
pip install -r requirements-dev.txt
```

Run the suite:

```
pytest tests/ -v
```

**Test results (12 tests):**

```
test_missing_content_file          PASSED  — exits 1 when content file absent
test_missing_template_file         PASSED  — exits 1 when template absent
test_creates_output_file           PASSED  — output file is written
test_prints_generated_path         PASSED  — stdout confirms output path
test_output_is_valid_ooxml         PASSED  — output is a valid OOXML package
test_heading1_style                PASSED  — # maps to Heading1
test_heading2_style                PASSED  — ## maps to Heading2
test_heading3_style                PASSED  — ### maps to Heading3
test_bullets_are_real_word_lists   PASSED  — bullets have numPr (real lists)
test_bullet_text_present           PASSED  — bullet text is in output
test_plain_paragraphs_present      PASSED  — body paragraphs are in output
test_sample_letter                 PASSED  — full letter round-trip
```

> **Note on test inspection:** The template embeds DM Sans fonts whose MIME
> types pandoc does not register in `[Content_Types].xml`. Tests therefore
> inspect OOXML directly via `zipfile` + `xml.etree.ElementTree` rather than
> using python-docx (which fails to open such files). This is a test-only
> concern — the generated DOCX opens correctly in Word and Google Docs.

