# md2docx

Generate a styled DOCX from a Markdown file and a DOCX style guide.

All visual style — fonts, spacing, margins, header, footer — is inherited from
the template. Body content is replaced entirely by the rendered Markdown.

Style is defined in the template DOCX and mapped to the output based on structured and formatted markdown content. To get started, open the [sample style guide](samples/sample-style-guide.docx) and read the [Mapping convention](#mapping-convention) section to understand the extended mapping convention.

Pre-built DOCX fragments can also be inserted directly into the output via [Partials](#partials) — useful for tables, cover pages, or any section better authored in Word than markdown.

## Contents

- [Usage](#usage)
  - [CLI](#cli)
  - [Library](#library)
- [Render method](#render-method)
  - [Mapping convention](#mapping-convention)
  - [Partials](#partials)
- [Samples](#samples)
- [Environment setup](docs/env-setup.md)
- [Architecture](docs/architecture.md)

---

## Usage

For prerequisites, dev setup, and running tests see [docs/env-setup.md](docs/env-setup.md).

### CLI

```
uv run md2docx <content.md> <template.docx> <output.docx> [--partial NAME=partial.docx ...]
```

**Example:**

```
uv run md2docx samples/sample-content.md samples/sample-style-guide.docx output.docx
```

### Library

Install the package and import directly — do not shell out to the CLI:

```python
from pathlib import Path
from md2docx import convert, ConversionError

try:
    convert(
        content=Path("doc.md"),
        template=Path("template.docx"),
        output=Path("out.docx"),
        partials={"SECTION": Path("section.docx")},  # optional
    )
except ConversionError as e:
    handle_error(str(e))
```

`convert()` raises `ConversionError` on any failure (missing file, missing pandoc, bad template, missing styles). It never calls `sys.exit()` — that stays in the CLI layer where it belongs.

Other Python projects should declare `md2docx` as a dependency:

```toml
# your project's pyproject.toml
dependencies = ["md2docx"]
```

## Render method

### Mapping convention

| Convention | Markdown content | DOCX style mapping | DOCX output |
|---|---|---|---|
| Markdown | `# H1` – `###### H6` | paragraph style: Heading 1–6 | Heading 1–6 style |
| Extended | `%Title` | paragraph style: Title | Title style |
| Extended | `%%Subtitle` | paragraph style: Subtitle | Subtitle style |
| Markdown | Plain paragraph | paragraph style: Normal | Normal style |
| Markdown | `**bold**` | styled literal: "Bold text" | Bold run |
| Markdown | `*italic*` | styled literal: "Italic text" | Italic run |
| Markdown | `` `code` `` or ` ``` ` | styled literal: "Code block text" | Code run / Code paragraph |
| Markdown | `- item` | styled literal: "Bullet point text" | Real Word list item |
| Extended | `left >> right` or `>> right` | styled literal: "RightAlignedTabStop" | Right tab stop on same line |
| Extended | `===` (exactly three) | — | Page break |
| Markdown | `---` (three or more) | — | Horizontal rule |
| Extended | Single newline | — | New paragraph (no gap) |
| Extended | Blank line | — | Visible empty paragraph |

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

#### Line breaks

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

### Partials

Partials let you splice pre-built DOCX sections into a generated document. Place a `{{NAME}}` placeholder on its own line in the markdown, then pass the matching DOCX file with `--partial`:

```
uv run md2docx content.md template.docx output.docx \
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

**Partial style guide:** partials may use a different style guide from the main document. A separate style DOCX for partials is included in `samples/sample-style-guide-partials.docx`.

## Samples

The `samples/` folder contains working examples that exercise every supported feature:

| File | Purpose |
|---|---|
| `samples/sample-content.md` | Content file demonstrating all syntax conventions |
| `samples/sample-style-guide.docx` | Matching style guide with all required labels defined |
| `samples/sample-content2.md` | Content file demonstrating partial interpolation |
| `samples/sample-style-guide-partials.docx` | Style guide for the partials example |
| `samples/sample-partial1.docx` | Partial DOCX for `{{SAMPLE_PARTIAL-1}}` |
| `samples/sample-partial2.docx` | Partial DOCX for `{{SAMPLE_PARTIAL-2}}` |

Run the standard example:

```
uv run md2docx samples/sample-content.md samples/sample-style-guide.docx samples/output.docx
```

Run the partials example:

```
uv run md2docx samples/sample-content2.md samples/sample-style-guide-partials.docx samples/output2.docx \
  --partial "SAMPLE_PARTIAL-1=samples/sample-partial1.docx" \
  --partial "SAMPLE_PARTIAL-2=samples/sample-partial2.docx"
```

The standard content file covers: title (`%`), subtitle (`%%`), all six heading levels, bold, italic, inline code, fenced code blocks, bullets, right-aligned tab stops (`>>`), horizontal rule (`---`), and page break (`===`).

