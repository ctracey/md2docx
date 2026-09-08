#!/usr/bin/env python3
"""Generate a styled DOCX from a Markdown file and a DOCX style template."""

import argparse
import io
import re
import shutil
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from style_map import RunStyle, read_style_map

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


_SOFT_NL = '\x00'


def mark_soft_newlines(text: str) -> str:
    """Replace single newlines with a sentinel so inject_blank_paragraphs ignores them.

    In standard markdown a single newline is a soft wrap (renders as a space).
    This converter treats each line as its own paragraph instead.  Explicit
    blank lines still produce a visible empty paragraph via inject_blank_paragraphs.
    """
    return re.sub(r'(?<!\n)\n(?!\n)', _SOFT_NL, text)


def restore_soft_newlines(text: str) -> str:
    """Convert sentinels back into pandoc paragraph breaks (double newline, no blank para)."""
    return text.replace(_SOFT_NL, '\n\n')


def inject_blank_paragraphs(text: str) -> str:
    r"""Insert one '\ ' paragraph per blank line so Word renders them as empty lines.

    Prepend \n so a single leading blank line matches the same \n{2,} pattern
    as blank lines in the middle of the document. N consecutive newlines means
    N-1 blank lines, so N-1 '\ ' paragraphs are inserted.
    """
    padded = '\n' + text.rstrip()

    def replace(m: re.Match) -> str:
        blanks = len(m.group(0)) - 1
        return '\n\n' + '\\ \n\n' * blanks

    result = re.sub(r'\n{2,}', replace, padded)
    return result.lstrip('\n') + '\n'


_BOOKMARK_RE = re.compile(rb'<w:bookmark(?:Start|End)\b[^>]*/>')
_HEADER_REF_RE = re.compile(rb'<w:headerReference\b[^>]*/>\s*')
_HEADER_REL_RE = re.compile(rb'<Relationship\b[^>]*/relationships/header[^>]*/>\s*')
_RELS_KEY = 'word/_rels/document.xml.rels'
_is_header_name = re.compile(r'word/(_rels/)?header\d+').match


def _write_zip(path: Path, names: list[str], files: dict[str, bytes]) -> None:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name in names:
            zout.writestr(name, files[name])
    path.write_bytes(buf.getvalue())


def _parse_header_rels(rels_bytes: bytes) -> dict[str, dict]:
    """Return {rId: {type, target}} for header relationships only."""
    result = {}
    for m in re.finditer(rb'<Relationship\b([^>]*)/?>', rels_bytes):
        attrs = m.group(1)
        if b'relationships/header' not in attrs:
            continue
        rid = re.search(rb'\bId="([^"]+)"', attrs)
        rtype = re.search(rb'\bType="([^"]+)"', attrs)
        target = re.search(rb'\bTarget="([^"]+)"', attrs)
        if rid and target:
            result[rid.group(1).decode()] = {
                'type': rtype.group(1).decode() if rtype else '',
                'target': target.group(1).decode(),
            }
    return result


def _parse_header_refs(doc_bytes: bytes) -> list[dict]:
    """Return [{rid, type}] for headerReference elements in sectPr."""
    result = []
    for m in re.finditer(rb'<w:headerReference\b([^>]*)/?>', doc_bytes):
        attrs = m.group(1)
        rid = re.search(rb'\br:id="([^"]+)"', attrs)
        wtype = re.search(rb'\bw:type="([^"]+)"', attrs)
        if rid and wtype:
            result.append({'rid': rid.group(1).decode(), 'type': wtype.group(1).decode()})
    return result


def _max_rid(rels_bytes: bytes) -> int:
    ids = [int(m.group(1)) for m in re.finditer(rb'\bId="rId(\d+)"', rels_bytes)]
    return max(ids, default=0)


def strip_bookmarks(docx_path: Path) -> None:
    """Remove all w:bookmarkStart and w:bookmarkEnd elements from document.xml.

    Pandoc inserts heading bookmarks for cross-reference support; Word renders
    them as visible blue bracket markers in the left margin.
    """
    with zipfile.ZipFile(docx_path, 'r') as zin:
        names = zin.namelist()
        files = {name: zin.read(name) for name in names}

    files['word/document.xml'] = _BOOKMARK_RE.sub(b'', files['word/document.xml'])

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name in names:
            zout.writestr(name, files[name])
    docx_path.write_bytes(buf.getvalue())


def sync_headers(docx_path: Path, template_path: Path) -> None:
    """Make output headers exactly match the template — no more, no less.

    Pandoc injects its own default headers regardless of the reference doc.
    This strips those and copies the template's headers (if any) verbatim,
    remapping relationship IDs to avoid conflicts with the output's existing rels.
    """
    with zipfile.ZipFile(template_path, 'r') as zf:
        tmpl_names = zf.namelist()
        tmpl_files = {n: zf.read(n) for n in tmpl_names}
    with zipfile.ZipFile(docx_path, 'r') as zf:
        out_names = list(zf.namelist())
        out_files = {n: zf.read(n) for n in out_names}

    # What headers does the template want?
    tmpl_hdr_rels = _parse_header_rels(tmpl_files.get(_RELS_KEY, b''))
    tmpl_hdr_refs = _parse_header_refs(tmpl_files.get('word/document.xml', b''))
    tmpl_hdr_files = {n: tmpl_files[n] for n in tmpl_names if _is_header_name(n)}

    # Strip all header content from the output
    out_files['word/document.xml'] = _HEADER_REF_RE.sub(b'', out_files['word/document.xml'])
    if _RELS_KEY in out_files:
        out_files[_RELS_KEY] = _HEADER_REL_RE.sub(b'', out_files[_RELS_KEY])
    out_names = [n for n in out_names if not _is_header_name(n)]

    if not tmpl_hdr_rels:
        _write_zip(docx_path, out_names, out_files)
        return

    # Remap template header rIds to fresh IDs that don't conflict with output
    base = _max_rid(out_files.get(_RELS_KEY, b''))
    rid_map = {old: f'rId{base + i + 1}' for i, old in enumerate(sorted(tmpl_hdr_rels))}

    # Copy template header XML files into output
    for name, content in tmpl_hdr_files.items():
        out_files[name] = content
        if name not in out_names:
            out_names.append(name)

    # Add relationship entries with remapped IDs
    new_rels = b''.join(
        (f'<Relationship Id="{rid_map[old]}" '
         f'Type="{tmpl_hdr_rels[old]["type"]}" '
         f'Target="{tmpl_hdr_rels[old]["target"]}"/>').encode()
        for old in sorted(rid_map)
    )
    out_files[_RELS_KEY] = out_files[_RELS_KEY].replace(
        b'</Relationships>', new_rels + b'</Relationships>'
    )

    # Add headerReference entries into sectPr
    new_refs = b''.join(
        (f'<w:headerReference r:id="{rid_map.get(ref["rid"], ref["rid"])}" '
         f'w:type="{ref["type"]}"/>').encode()
        for ref in tmpl_hdr_refs
        if ref['rid'] in rid_map
    )
    out_files['word/document.xml'] = out_files['word/document.xml'].replace(
        b'</w:sectPr>', new_refs + b'</w:sectPr>', 1
    )

    _write_zip(docx_path, out_names, out_files)


def _merge_run_style(rpr: ET.Element, style: RunStyle) -> None:
    """Add template-defined properties to an rPr element without overriding existing ones."""
    if style.font is not None and rpr.find(f'{{{W}}}rFonts') is None:
        el = ET.SubElement(rpr, f'{{{W}}}rFonts')
        for attr in ('ascii', 'hAnsi', 'cs', 'eastAsia'):
            el.set(f'{{{W}}}{attr}', style.font)
    if style.color is not None and rpr.find(f'{{{W}}}color') is None:
        el = ET.SubElement(rpr, f'{{{W}}}color')
        el.set(f'{{{W}}}val', style.color)
    if style.size_half_pt is not None:
        if rpr.find(f'{{{W}}}sz') is None:
            ET.SubElement(rpr, f'{{{W}}}sz').set(f'{{{W}}}val', str(style.size_half_pt))
        if rpr.find(f'{{{W}}}szCs') is None:
            ET.SubElement(rpr, f'{{{W}}}szCs').set(f'{{{W}}}val', str(style.size_half_pt))


def apply_run_styles(docx_path: Path, style_map: dict[str, RunStyle]) -> None:
    """Post-process output DOCX to apply template run styles to bold/italic runs in Normal paragraphs."""
    italic_style = style_map.get("Italic text")
    bold_style = style_map.get("Bold text")
    if not italic_style and not bold_style:
        return

    with zipfile.ZipFile(docx_path, 'r') as zin:
        names = zin.namelist()
        files = {name: zin.read(name) for name in names}

    doc_bytes = files['word/document.xml']

    for m in re.finditer(rb'xmlns:(\w+)="([^"]+)"', doc_bytes):
        ET.register_namespace(m.group(1).decode(), m.group(2).decode())

    root = ET.fromstring(doc_bytes)
    modified = False

    _SKIP = {f'Heading{i}' for i in range(1, 7)} | {'Title', 'Subtitle'}

    for p in root.findall(f'.//{{{W}}}p'):
        ppr = p.find(f'{{{W}}}pPr')
        style_el = ppr.find(f'{{{W}}}pStyle') if ppr is not None else None
        para_style = style_el.get(f'{{{W}}}val') if style_el is not None else ''
        if para_style in _SKIP:
            continue

        for r in p.findall(f'{{{W}}}r'):
            rpr = r.find(f'{{{W}}}rPr')
            if rpr is None:
                continue
            has_italic = rpr.find(f'{{{W}}}i') is not None
            has_bold = rpr.find(f'{{{W}}}b') is not None
            if has_italic and italic_style:
                _merge_run_style(rpr, italic_style)
                modified = True
            elif has_bold and bold_style:
                _merge_run_style(rpr, bold_style)
                modified = True

    if not modified:
        return

    new_xml = ET.tostring(root, encoding='unicode')
    files['word/document.xml'] = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        + new_xml.encode('utf-8')
    )
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name in names:
            zout.writestr(name, files[name])
    docx_path.write_bytes(buf.getvalue())


_SPECIAL_LINE_RE = re.compile(r'^(#{1,6}\s|[-*+]\s|>\s|```|\s*$)')


def inject_title_styles(text: str, style_map: dict[str, RunStyle]) -> str:
    """Wrap the first one or two plain lines in pandoc custom-style fenced divs.

    Convention: the first non-special line of the document becomes the Title
    paragraph; the immediately following non-special line becomes the Subtitle.
    Lines starting with #, -, >, ``` or blank are considered 'special' and
    break the title/subtitle detection.
    """
    title_entry = style_map.get("Title text")
    subtitle_entry = style_map.get("Subtitle text")
    if not title_entry:
        return text

    title_style = title_entry.para_style or "Title"
    subtitle_style = subtitle_entry.para_style if subtitle_entry else "Subtitle"

    lines = text.split('\n')
    out = []
    i = 0

    if i < len(lines) and not _SPECIAL_LINE_RE.match(lines[i]):
        out += [f'::: {{custom-style="{title_style}"}}', lines[i], ':::']
        i += 1
        if subtitle_entry and i < len(lines) and not _SPECIAL_LINE_RE.match(lines[i]):
            out += [f'::: {{custom-style="{subtitle_style}"}}', lines[i], ':::']
            i += 1

    out += lines[i:]
    return '\n'.join(out)


_EXTENSION_CONSTRUCTS = [
    (r'\*\*\S', "Bold text", "**bold**"),
    (r'(?<!\*)\*(?!\*)\S', "Italic text", "*italic*"),
]


def check_required_styles(content: str, style_map: dict[str, RunStyle]) -> None:
    """Exit with a clear error if the content uses constructs whose style labels are absent.

    See README — Style mapping — Extension mappings.
    """
    missing = [
        (display, label)
        for pattern, label, display in _EXTENSION_CONSTRUCTS
        if re.search(pattern, content) and label not in style_map
    ]
    if not missing:
        return
    details = "\n".join(
        f'  {display}  →  add a paragraph labelled "{label}" to the style template'
        for display, label in missing
    )
    sys.exit(
        f"Error: style template is missing required labels:\n{details}\n"
        "See README — Style mapping — Extension mappings."
    )


def main():
    parser = argparse.ArgumentParser(
        description="Generate a styled DOCX from Markdown and a DOCX template.",
        epilog="Requires pandoc to be installed (brew install pandoc).",
    )
    parser.add_argument("content", type=Path, help="Markdown input file (.md)")
    parser.add_argument("template", type=Path, help="DOCX style template (.docx)")
    parser.add_argument("output", type=Path, help="Output DOCX file (.docx)")
    args = parser.parse_args()

    if not args.content.exists():
        sys.exit(f"Error: content file not found: {args.content}")
    if not args.template.exists():
        sys.exit(f"Error: template file not found: {args.template}")
    if shutil.which("pandoc") is None:
        sys.exit("Error: pandoc not found. Install with: brew install pandoc")

    try:
        style_map = read_style_map(args.template)
    except ValueError as e:
        sys.exit(f"Error: {e}")

    raw = args.content.read_text()
    check_required_styles(raw, style_map)
    raw = inject_title_styles(raw, style_map)
    modified = restore_soft_newlines(inject_blank_paragraphs(mark_soft_newlines(raw)))

    with tempfile.NamedTemporaryFile(mode='w', suffix='.md', delete=False) as tmp:
        tmp.write(modified)
        tmp_path = Path(tmp.name)

    try:
        result = subprocess.run(
            [
                "pandoc",
                str(tmp_path),
                "--reference-doc", str(args.template),
                "-o", str(args.output),
            ],
            capture_output=True,
            text=True,
        )
    finally:
        tmp_path.unlink()

    if result.returncode != 0:
        sys.exit(f"Error: pandoc failed:\n{result.stderr}")

    strip_bookmarks(args.output)
    sync_headers(args.output, args.template)
    apply_run_styles(args.output, style_map)
    print(f"Generated: {args.output}")


if __name__ == "__main__":
    main()
