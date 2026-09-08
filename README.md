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

## Style mapping

Styles are applied through two mechanisms: native mappings and extension mappings.

### Native mappings

These are handled automatically by pandoc using the named paragraph styles defined in the template. No configuration required.

| Markdown | Template style |
|---|---|
| `# Heading 1` through `###### Heading 6` | Heading 1–6 |
| `- item` | List style (real Word list items) |
| Plain paragraph | Body text |

### Extension mappings

Run-level styles (font, size, colour, etc.) are discovered by scanning the style template for labelled example paragraphs. Each label is a paragraph whose full text exactly matches one of the names below. The run formatting of that paragraph defines how the corresponding markdown construct is rendered.

| Markdown | Template label | Notes |
|---|---|---|
| `%Title text` | `Title text` | |
| `%%Subtitle text` | `Subtitle text` | |
| Plain text | `Normal text` | |
| `**bold**` | `Bold text` | |
| `*italic*` | `Italic text` | |

**Title and Subtitle convention:** prefix a line with `%` for Title or `%%` for Subtitle. Each can appear anywhere in the document independently. `%%` is always tested before `%` so double-percent lines are never misread as a title.

```
%My Document Title
%%My Document Subtitle

# Section One

%%An inline subtitle anywhere
```

**How to define a style:** add a paragraph to the style template with the exact label text, styled as you want that construct to appear. Each label must appear exactly once — duplicate labels are an error.

### Line breaks

This converter treats **every newline as a paragraph break**, not a soft wrap. Standard markdown collapses a single newline into a space; this tool does not.

```
normal text       ← paragraph 1
**bold text**     ← paragraph 2 (immediately follows, no gap)
*italic text*     ← paragraph 3
```

An **explicit blank line** inserts a visible empty paragraph:

```
paragraph one

paragraph two     ← blank paragraph appears between these two
```

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

> **Note on test inspection:** The template embeds DM Sans fonts whose MIME
> types pandoc does not register in `[Content_Types].xml`. Tests therefore
> inspect OOXML directly via `zipfile` + `xml.etree.ElementTree` rather than
> using python-docx (which fails to open such files). This is a test-only
> concern — the generated DOCX opens correctly in Word and Google Docs.

