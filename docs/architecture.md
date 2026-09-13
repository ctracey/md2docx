# Architecture

## Tech stack

| Layer | Tool |
|---|---|
| Language | Python 3.13 |
| Env / package manager | [uv](https://docs.astral.sh/uv/) |
| Build backend | [hatchling](https://hatch.pypa.io/) |
| Conversion engine | [pandoc](https://pandoc.org/) ≥ 3.11 |
| Testing | pytest + python-docx (dev only) |

## Pipeline

The pipeline has four stages:

```
Read template → StyleMap
Read markdown → Content (validate + transform)
Render        → pandoc produces raw DOCX
Post-process  → fix up the DOCX output
```

**StyleMap** (`style_map.py`) — scans the template body for labelled paragraphs (`Bold text`, `Italic text`, `Code block text`, `Bullet point text`, `RightAlignedTabStop`) and reads paragraph styles and tab stop definitions. The run properties of each label define how that markdown construct is rendered.

**Content** (`content.py`) — validates the markdown against the StyleMap (raises an error if required labels are missing), then transforms the raw text before pandoc sees it:

1. Title/Subtitle prefix lines (`%`, `%%`) → pandoc custom-style fenced divs
2. Page break markers (`===`) → sentinel character that survives the pipeline
3. Right-tab markers (`>>`) → sentinel so pandoc doesn't treat them as blockquotes
4. Single newlines → marked as soft breaks so every source line becomes its own paragraph
5. Blank lines → explicit `\ ` empty paragraphs
6. Sentinels resolved → page break sentinels become raw OpenXML fenced blocks

**Render** (`converter.py`) — passes the transformed markdown and template to pandoc, which produces a raw DOCX.

**Post-process** (`document.py`) — manipulates `word/document.xml` directly via `zipfile` + `xml.etree.ElementTree`:

1. Heading bookmarks stripped (pandoc inserts them; Word renders them as visible margin markers)
2. Headers synced from the template (pandoc ignores reference-doc headers)
3. Fonts synced from the template (pandoc may drop embedded font binaries)
4. Right tab stops applied — sentinel runs split and `<w:tab/>` elements inserted
5. Run styles applied — bold, italic, code, and bullet runs get font/colour/size from template labels
6. Partials spliced — `{{NAME}}` placeholder paragraphs replaced with content from named DOCX files

## Module structure

```
src/md2docx/
├── __init__.py      # public API: convert, ConversionError
├── exceptions.py    # ConversionError
├── style_map.py     # StyleMap + RunStyle — reads styles from a template
├── content.py       # Content — validates and transforms markdown
├── document.py      # post-processing operations on the output DOCX
└── converter.py     # pipeline orchestrator + CLI entry point
```

Test files mirror modules 1:1 under `tests/`.
