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
python generate.py ref/template-content.md ref/template-style.docx output.docx
```

### Markdown features

| Markdown | DOCX output |
|---|---|
| `# Heading 1` through `###### Heading 6` | Heading 1–6 styles |
| `- item` | Real Word bullet (numPr) |
| Plain paragraph | Body text style |
| `**bold**` | Bold run |
| `*italic*` | Italic run |
| Single newline | New paragraph (no blank line between) |
| Blank line | Empty paragraph (visible vertical gap) |

#### Line breaks vs blank lines

This converter treats **every newline as a paragraph break**, not a soft wrap.
Standard markdown collapses a single newline into a space; this tool does not.

```
normal text       ← paragraph 1
**bold text**     ← paragraph 2 (immediately follows, no gap)
*italic text*     ← paragraph 3
```

An **explicit blank line** inserts a visible empty paragraph between content:

```
paragraph one

paragraph two     ← blank paragraph appears between these two
```

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

Run `pytest tests/ -v` — 33 tests, all passing.

> **Note on test inspection:** The template embeds DM Sans fonts whose MIME
> types pandoc does not register in `[Content_Types].xml`. Tests therefore
> inspect OOXML directly via `zipfile` + `xml.etree.ElementTree` rather than
> using python-docx (which fails to open such files). This is a test-only
> concern — the generated DOCX opens correctly in Word and Google Docs.

