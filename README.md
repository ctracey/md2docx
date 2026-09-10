# md2docx

Generate a styled DOCX from a Markdown file and a DOCX style template.

All visual style — fonts, spacing, margins, header, footer — is inherited from
the template. Body content is replaced entirely by the rendered Markdown.

## Prerequisites

- [pandoc](https://pandoc.org/) ≥ 3.11 — `brew install pandoc`
- [uv](https://docs.astral.sh/uv/) — `brew install uv`

No Python packages are needed at runtime. Python itself is managed by uv (see [Development setup](#development-setup)).

## Usage

```
python generate.py <content.md> <template.docx> <output.docx> [--partial NAME=partial.docx ...]
```

**Example:**

```
python generate.py ref/template-content.md ref/template-style.docx output.docx
```

## Partials

Partials let you splice pre-built DOCX sections into a generated document. Place a `{{NAME}}` placeholder on its own line in the markdown, then pass the matching DOCX file with `--partial`:

```
python generate.py content.md template.docx output.docx \
  --partial INTRO=intro.docx \
  --partial TABLE=data-table.docx
```

The `--partial` flag can be repeated for as many named placeholders as needed.

**In the markdown:**

```markdown
## Pre-generated section

{{INTRO}}

## Another section

{{TABLE}}
```

Each placeholder line is replaced with the full body content of the named partial DOCX. The partial renders with its own font properties — font face, size, colour — exactly as it looks when opened in Word, independent of the output template's style definitions. The template does not override partial styling.

**Partial style template:** partials may use a different style template from the main document. A separate style DOCX for partials is included in `sample/sample-style-partials.docx`.

## Sample

The `sample/` folder contains working examples that exercise every supported feature:

| File | Purpose |
|---|---|
| `sample/sample-content.md` | Content file demonstrating all syntax conventions |
| `sample/sample-style.docx` | Matching style template with all required labels defined |
| `sample/sample-content2.md` | Content file demonstrating partial interpolation |
| `sample/sample-style-partials.docx` | Style template for the partials example |
| `sample/sample-partial1.docx` | Partial DOCX for `{{SAMPLE_PARTIAL-1}}` |
| `sample/sample-partial2.docx` | Partial DOCX for `{{SAMPLE_PARTIAL-2}}` |

Run the standard example:

```
python generate.py sample/sample-content.md sample/sample-style.docx sample/output.docx
```

Run the partials example:

```
python generate.py sample/sample-content2.md sample/sample-style-partials.docx sample/output2.docx \
  --partial "SAMPLE_PARTIAL-1=sample/sample-partial1.docx" \
  --partial "SAMPLE_PARTIAL-2=sample/sample-partial2.docx"
```

The standard content file covers: title (`%`), subtitle (`%%`), all six heading levels, bold, italic, inline code, fenced code blocks, bullets, right-aligned tab stops (`>>`), horizontal rule (`---`), and page break (`===`).

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
| `- item` | Bullet point text label | Real Word list item |
| `left >> right` or `>> right` | RightAlignedTabStop | Right tab stop on same line |
| `===` (exactly three) | — | Page break |
| `---` (three or more) | — | Horizontal rule |
| Single newline | — | New paragraph (no gap) |
| Blank line | — | Visible empty paragraph |

**Native mappings** (headings, bullets, blank lines) are applied automatically by pandoc using the named paragraph styles in the template.

**Label mappings** (Bold text, Italic text, Code block text, Bullet point text, RightAlignedTabStop) are discovered by scanning the template body for a paragraph whose full text exactly matches the label name. The run formatting of that paragraph — font, size, colour, etc. — is applied to the corresponding markdown construct. Each label must appear exactly once; duplicates are an error. The Code block label applies to both inline backtick code and fenced code blocks.

**Right-aligned tab stop (`>>`):** the `RightAlignedTabStop` label is special — it must be a paragraph demonstrating a right-aligned tab stop (validated on load; errors if the paragraph has no right-aligned tab in its pPr). Use `>>` anywhere on a line to split it: text before `>>` stays left, text after `>>` is pulled to the right tab stop position. Works in headings and body text alike. Nothing on the left is valid.

```
## Section heading >> 2024-01

Normal text >> right aligned note

>> purely right aligned
```

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

## Development setup

Python 3.13 is pinned in `.python-version`. uv reads this automatically.

```
# Install uv (once, system-wide)
brew install uv

# Install Python 3.13 and project dependencies
uv sync

# Run the tests
uv run pytest tests/ -v
```

`uv sync` creates a `.venv`, installs the pinned Python version if needed, and installs all dev dependencies from `uv.lock`. No separate pip or virtualenv step required.

## Repository structure

```
.
├── generate.py       # main script — the tool
├── pyproject.toml    # project metadata and dev dependencies
├── uv.lock           # locked dependency versions
├── .python-version   # pins Python 3.13 for uv
├── specs/
│   └── generate.md   # behavioural specification
├── tests/
│   ├── conftest.py   # shared fixtures
│   ├── test_generate.py
│   └── fixtures/
│       ├── simple.md
│       └── headings.md
```

## Running tests

```
uv run pytest tests/ -v
```

> **Note on test inspection:** The template embeds DM Sans fonts whose MIME
> types pandoc does not register in `[Content_Types].xml`. Tests therefore
> inspect OOXML directly via `zipfile` + `xml.etree.ElementTree` rather than
> using python-docx (which fails to open such files). This is a test-only
> concern — the generated DOCX opens correctly in Word and Google Docs.

