# Specification: generate.py

## Purpose

Convert a Markdown file to a styled DOCX, inheriting all visual style (fonts,
spacing, margins, header, footer) from a DOCX template.

## Inputs

| Argument   | Type | Description                                      |
|------------|------|--------------------------------------------------|
| `content`  | path | Markdown source file (`.md`)                     |
| `template` | path | DOCX style reference file (`.docx`)              |
| `output`   | path | Destination path for the generated DOCX (`.docx`) |

## Behaviour

### Markdown mapping

| Markdown element      | DOCX output                        |
|-----------------------|------------------------------------|
| `# Heading 1`         | Heading 1 style                    |
| `## Heading 2`        | Heading 2 style                    |
| `### Heading 3`       | Heading 3 style                    |
| `- bullet item`       | List Bullet style (real Word list) |
| Plain paragraph       | Normal / body text style           |
| `**bold**`            | Bold run                           |
| `*italic*`            | Italic run                         |

### Style inheritance

All paragraph and character styles — including fonts, spacing, margins, page
size, header, and footer — are drawn from the template DOCX. The template body
content is replaced entirely by the rendered Markdown.

### Error conditions

| Condition                    | Exit behaviour                                        |
|------------------------------|-------------------------------------------------------|
| `content` file not found     | Exit 1 with message `Error: content file not found: <path>` |
| `template` file not found    | Exit 1 with message `Error: template file not found: <path>` |
| `pandoc` not on PATH         | Exit 1 with message `Error: pandoc not found. Install with: brew install pandoc` |
| pandoc returns non-zero      | Exit 1 with pandoc stderr forwarded                   |

### Success

Writes the output DOCX to the specified path and prints:
```
Generated: <output path>
```

## Dependencies

- **pandoc** ≥ 3.x — document conversion engine (`brew install pandoc`)
- **Python** ≥ 3.9 — standard library only; no pip packages required at runtime
- **python-docx** — test dependency only, used to inspect output in tests
