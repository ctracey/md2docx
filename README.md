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

| Markdown content | DOCX style mapping | DOCX output |
|---|---|---|
| `# H1` – `###### H6` | Heading 1–6 style | Heading 1–6 style |
| `%Title` | Title style | Title style |
| `%%Subtitle` | Subtitle style | Subtitle style |
| Plain paragraph | Normal style | Normal style |
| `**bold**` | Bold text label | Bold run |
| `*italic*` | Italic text label | Italic run |
| `` `code` `` or ` ``` ` | Code block text label | Code run / Code paragraph |
| `- item` | List style | Real Word list item |
| Single newline | — | New paragraph (no gap) |
| Blank line | — | Visible empty paragraph |

**Native mappings** (headings, bullets, blank lines) are applied automatically by pandoc using the named paragraph styles in the template.

**Label mappings** (Bold text, Italic text, Code block text) are discovered by scanning the template body for a paragraph whose full text exactly matches the label name. The run formatting of that paragraph — font, size, colour, etc. — is applied to the corresponding markdown construct. Each label must appear exactly once; duplicates are an error. The Code block label applies to both inline backtick code and fenced code blocks.

**Title and Subtitle** are matched directly to the `Title` and `Subtitle` paragraph styles defined in the template. Prefix a line with `%` for Title or `%%` for Subtitle — each can appear anywhere in the document independently.

```
%My Document Title
%%My Document Subtitle

# Section One

%%An inline subtitle anywhere
```

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

