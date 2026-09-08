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

from style_map import RunStyle, read_style_map, read_paragraph_style_ids, read_right_tab_stop, validate_template

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


_SOFT_NL = '\x00'
_PAGE_BREAK_MARKER = ''   # U+E002 PUA — survives preprocessing pipeline
_PAGE_BREAK_XML = (
    '```{=openxml}\n'
    '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:r><w:br w:type="page"/></w:r></w:p>\n'
    '```'
)


def inject_page_break_markers(text: str) -> str:
    """Replace lines that are exactly '===' with a page-break sentinel.

    The sentinel travels through the preprocessing pipeline untouched.
    It is swapped for the pandoc raw OpenXML page-break block after
    restore_soft_newlines, so the fenced block syntax is never mangled.
    """
    lines = text.split('\n')
    return '\n'.join(_PAGE_BREAK_MARKER if line == '===' else line for line in lines)


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

_CONTENT_TYPES_KEY = '[Content_Types].xml'

_SKIP_REL_TARGETS = frozenset({
    'styles.xml', 'settings.xml', 'fontTable.xml', 'webSettings.xml',
    'endnotes.xml', 'footnotes.xml', 'numbering.xml', 'document.xml',
    'comments.xml', 'theme/theme1.xml',
})

_MEDIA_CONTENT_TYPES = {
    '.png': 'image/png',
    '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
    '.gif': 'image/gif',
    '.tiff': 'image/tiff', '.tif': 'image/tiff',
    '.bmp': 'image/bmp',
    '.svg': 'image/svg+xml',
    '.emf': 'image/x-emf',
    '.wmf': 'image/x-wmf',
}


def _resolve_style_rprs(styles_bytes: bytes, theme_bytes: bytes = b'') -> dict[str, ET.Element]:
    """Return {styleId: effective w:rPr element} for all paragraph styles.

    Follows basedOn chains starting from docDefaults so each entry holds the
    fully inherited run properties. Includes '__default__' for styles not
    explicitly defined in the document (uses docDefaults + theme body font).
    """
    WW = f'{{{W}}}'
    A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    styles_root = ET.fromstring(styles_bytes)

    # Base: document-level run defaults
    base_props: dict[str, ET.Element] = {}
    doc_def_rpr = styles_root.find(f'{WW}docDefaults/{WW}rPrDefault/{WW}rPr')
    if doc_def_rpr is not None:
        for prop in doc_def_rpr:
            base_props[prop.tag] = prop

    # Add the partial's theme body font so undefined styles use it explicitly
    if theme_bytes and f'{WW}rFonts' not in base_props:
        try:
            theme_root = ET.fromstring(theme_bytes)
            minor = theme_root.find(f'.//{{{A}}}fontScheme/{{{A}}}minorFont/{{{A}}}latin')
            if minor is not None and minor.get('typeface'):
                fonts_el = ET.Element(f'{WW}rFonts')
                for attr in ('ascii', 'hAnsi', 'cs', 'eastAsia'):
                    fonts_el.set(f'{WW}{attr}', minor.get('typeface'))
                base_props[f'{WW}rFonts'] = fonts_el
        except Exception:
            pass

    styles: dict[str, ET.Element] = {
        s.get(f'{WW}styleId'): s
        for s in styles_root.findall(f'{WW}style')
        if s.get(f'{WW}type') == 'paragraph' and s.get(f'{WW}styleId')
    }

    def _effective(style_id: str, visited: set) -> dict[str, ET.Element]:
        if style_id in visited or style_id not in styles:
            return dict(base_props)
        visited.add(style_id)
        style_el = styles[style_id]
        based_on = style_el.find(f'{WW}basedOn')
        parent_id = based_on.get(f'{WW}val') if based_on is not None else None
        props: dict[str, ET.Element] = _effective(parent_id, visited) if parent_id else dict(base_props)
        rpr = style_el.find(f'{WW}rPr')
        if rpr is not None:
            for prop in rpr:
                props[prop.tag] = prop
        return props

    result: dict[str, ET.Element] = {}
    for sid in styles:
        props = _effective(sid, set())
        if props:
            rpr_el = ET.Element(f'{WW}rPr')
            for prop in props.values():
                rpr_el.append(ET.fromstring(ET.tostring(prop)))
            result[sid] = rpr_el

    # Fallback for styles not in this document (e.g. built-in Word styles)
    if base_props:
        default_el = ET.Element(f'{WW}rPr')
        for prop in base_props.values():
            default_el.append(ET.fromstring(ET.tostring(prop)))
        result['__default__'] = default_el

    return result


def _inline_partial_styles(paras: list, styles_bytes: bytes, theme_bytes: bytes = b'') -> None:
    """Strip all paragraph style names and inline run properties from the partial.

    Every paragraph's pStyle is removed so the content is self-contained and
    immune to the output template's style definitions. For styles explicitly
    defined in the partial, the full effective rPr (following basedOn chains) is
    inlined. For styles not defined in the partial (built-in Word styles), the
    partial's document defaults + theme body font are inlined as a baseline.
    """
    WW = f'{{{W}}}'
    style_rprs = _resolve_style_rprs(styles_bytes, theme_bytes)
    default_rpr = style_rprs.get('__default__')

    for p in paras:
        ppr = p.find(f'{WW}pPr')
        if ppr is None:
            continue
        style_el = ppr.find(f'{WW}pStyle')
        if style_el is None:
            continue

        effective_rpr = style_rprs.get(style_el.get(f'{WW}val', ''), default_rpr)
        ppr.remove(style_el)

        if effective_rpr is None:
            continue

        for r in p.findall(f'{WW}r'):
            existing = r.find(f'{WW}rPr')
            if existing is None:
                r.insert(0, ET.fromstring(ET.tostring(effective_rpr)))
            else:
                for prop in effective_rpr:
                    if existing.find(prop.tag) is None:
                        existing.append(ET.fromstring(ET.tostring(prop)))


def _parse_rels(rels_bytes: bytes) -> dict[str, dict]:
    """Return {rId: {type, target}} for all relationships."""
    result = {}
    for m in re.finditer(rb'<Relationship\b([^>]*)/?>', rels_bytes):
        attrs = m.group(1)
        rid = re.search(rb'\bId="([^"]+)"', attrs)
        rtype = re.search(rb'\bType="([^"]+)"', attrs)
        target = re.search(rb'\bTarget="([^"]+)"', attrs)
        if rid and rtype and target:
            result[rid.group(1).decode()] = {
                'type': rtype.group(1).decode(),
                'target': target.group(1).decode(),
            }
    return result


def _ensure_content_type(ct_bytes: bytes, ext: str) -> bytes:
    """Add a Default content type entry for the extension if not already declared."""
    ct = _MEDIA_CONTENT_TYPES.get(ext)
    if ct is None:
        return ct_bytes
    ext_bare = ext.lstrip('.').lower()
    if ext_bare.encode() in ct_bytes:
        return ct_bytes
    entry = f'<Default Extension="{ext_bare}" ContentType="{ct}"/>'.encode()
    return ct_bytes.replace(b'</Types>', entry + b'</Types>')


def _collect_num_ids(paras: list) -> set[str]:
    """Return all w:numId val= values referenced in the given paragraph elements."""
    ids = set()
    for p in paras:
        for el in p.findall(f'.//{{{W}}}numId'):
            val = el.get(f'{{{W}}}val')
            if val and val != '0':
                ids.add(val)
    return ids


def _merge_numbering(out_files: dict, p_files: dict, p_num_ids: set) -> dict[str, str]:
    """Merge partial numbering definitions into the output's numbering.xml.

    Returns {old_numId: new_numId} for any IDs that were remapped to avoid conflicts.
    """
    if not p_num_ids or 'word/numbering.xml' not in p_files:
        return {}

    WW = f'{{{W}}}'
    for m in re.finditer(rb'xmlns:(\w+)="([^"]+)"', p_files['word/numbering.xml']):
        ET.register_namespace(m.group(1).decode(), m.group(2).decode())

    p_num_root = ET.fromstring(p_files['word/numbering.xml'])
    p_nums = {n.get(f'{WW}numId'): n for n in p_num_root.findall(f'{WW}num')}
    p_abs_nums = {a.get(f'{WW}abstractNumId'): a for a in p_num_root.findall(f'{WW}abstractNum')}

    needed_nums = {nid: p_nums[nid] for nid in p_num_ids if nid in p_nums}
    needed_abs_ids: set[str] = set()
    for num_el in needed_nums.values():
        abs_ref = num_el.find(f'{WW}abstractNumId')
        if abs_ref is not None:
            needed_abs_ids.add(abs_ref.get(f'{WW}val', ''))
    needed_abs_nums = {aid: p_abs_nums[aid] for aid in needed_abs_ids if aid in p_abs_nums}

    if not needed_nums:
        return {}

    out_num_bytes = out_files.get('word/numbering.xml')
    if out_num_bytes:
        out_num_root = ET.fromstring(out_num_bytes)
    else:
        out_num_root = ET.fromstring(
            f'<w:numbering xmlns:w="{W}"/>'
        )

    out_num_ids = {n.get(f'{WW}numId') for n in out_num_root.findall(f'{WW}num')}
    out_abs_ids = {a.get(f'{WW}abstractNumId') for a in out_num_root.findall(f'{WW}abstractNum')}
    max_abs = max((int(x) for x in out_abs_ids if x and x.isdigit()), default=0)
    max_num = max((int(x) for x in out_num_ids if x and x.isdigit()), default=0)

    abs_id_map: dict[str, str] = {}
    for old_aid in needed_abs_ids:
        if old_aid in out_abs_ids:
            max_abs += 1
            abs_id_map[old_aid] = str(max_abs)
        else:
            abs_id_map[old_aid] = old_aid

    num_id_map: dict[str, str] = {}
    for old_nid in needed_nums:
        if old_nid in out_num_ids:
            max_num += 1
            num_id_map[old_nid] = str(max_num)
        else:
            num_id_map[old_nid] = old_nid

    # abstractNum elements must precede num elements — insert before first existing num
    first_num_idx = next(
        (i for i, c in enumerate(out_num_root) if c.tag == f'{WW}num'),
        None,
    )
    for old_aid, abs_el in needed_abs_nums.items():
        copy = ET.fromstring(ET.tostring(abs_el))
        copy.set(f'{WW}abstractNumId', abs_id_map[old_aid])
        if first_num_idx is not None:
            out_num_root.insert(first_num_idx, copy)
            first_num_idx += 1
        else:
            out_num_root.append(copy)

    for old_nid, num_el in needed_nums.items():
        copy = ET.fromstring(ET.tostring(num_el))
        copy.set(f'{WW}numId', num_id_map[old_nid])
        abs_ref = copy.find(f'{WW}abstractNumId')
        if abs_ref is not None:
            abs_ref.set(f'{WW}val', abs_id_map.get(abs_ref.get(f'{WW}val', ''), abs_ref.get(f'{WW}val', '')))
        out_num_root.append(copy)

    new_xml = ET.tostring(out_num_root, encoding='unicode')
    out_files['word/numbering.xml'] = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        + new_xml.encode('utf-8')
    )

    if 'word/numbering.xml' not in out_files.get('_names_', []):
        pass  # already in out_files dict

    # Ensure numbering rel exists when we created numbering.xml from scratch
    if out_num_bytes is None and _RELS_KEY in out_files:
        if b'relationships/numbering' not in out_files[_RELS_KEY]:
            new_rid = f'rId{_max_rid(out_files[_RELS_KEY]) + 1}'
            entry = (
                f'<Relationship Id="{new_rid}" '
                f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" '
                f'Target="numbering.xml"/>'
            ).encode()
            out_files[_RELS_KEY] = out_files[_RELS_KEY].replace(b'</Relationships>', entry + b'</Relationships>')

    return num_id_map


def _remap_rids(root: ET.Element, rid_map: dict[str, str]) -> None:
    """Replace all rId attribute values in the element tree according to rid_map."""
    for el in root.iter():
        for attr, val in list(el.attrib.items()):
            if val in rid_map:
                el.set(attr, rid_map[val])


def _remap_numids(root: ET.Element, num_id_map: dict[str, str]) -> None:
    """Replace w:numId w:val values in the element tree according to num_id_map."""
    for el in root.findall(f'.//{{{W}}}numId'):
        val = el.get(f'{{{W}}}val')
        if val in num_id_map:
            el.set(f'{{{W}}}val', num_id_map[val])


def splice_partials(docx_path: Path, partials: dict[str, Path]) -> None:
    """Replace {{NAME}} placeholder paragraphs with content from named partial DOCX files."""
    if not partials:
        return

    with zipfile.ZipFile(docx_path, 'r') as zin:
        out_names = list(zin.namelist())
        out_files = {n: zin.read(n) for n in out_names}

    doc_bytes = out_files['word/document.xml']
    for m in re.finditer(rb'xmlns:(\w+)="([^"]+)"', doc_bytes):
        ET.register_namespace(m.group(1).decode(), m.group(2).decode())

    root = ET.fromstring(doc_bytes)
    body = root.find(f'{{{W}}}body')
    modified = False

    # Warn about any {{...}} placeholders with no matching partial
    all_placeholders = {
        ''.join(t.text or '' for t in p.findall(f'.//{{{W}}}t')).strip()
        for p in body.findall(f'{{{W}}}p')
    }
    unclaimed = {
        ph[2:-2] for ph in all_placeholders
        if ph.startswith('{{') and ph.endswith('}}') and ph[2:-2] not in partials
    }
    for name in sorted(unclaimed):
        print(f"Warning: placeholder {{{{{name}}}}} has no matching --partial", file=sys.stderr)

    for name, partial_path in partials.items():
        placeholder = '{{' + name + '}}'

        with zipfile.ZipFile(partial_path, 'r') as pzip:
            p_files = {n: pzip.read(n) for n in pzip.namelist()}

        p_doc_bytes = p_files.get('word/document.xml', b'')
        for m in re.finditer(rb'xmlns:(\w+)="([^"]+)"', p_doc_bytes):
            ET.register_namespace(m.group(1).decode(), m.group(2).decode())

        # Remap external resource relationships (images, hyperlinks — not core XML parts)
        p_rels = _parse_rels(p_files.get(_RELS_KEY, b''))
        rid_map: dict[str, str] = {}

        for old_rid, rel in p_rels.items():
            target = rel['target']
            if target in _SKIP_REL_TARGETS:
                continue
            is_external = target.startswith(('http://', 'https://', 'mailto:'))
            new_rid = f'rId{_max_rid(out_files.get(_RELS_KEY, b"")) + len(rid_map) + 1}'
            rid_map[old_rid] = new_rid

            if not is_external:
                src_key = 'word/' + target
                if src_key in p_files:
                    dst_key = src_key
                    if dst_key in out_files:
                        ext = Path(src_key).suffix
                        stem = Path(src_key).stem
                        parent = str(Path(src_key).parent)
                        dst_key = f'{parent}/{stem}_{new_rid}{ext}'
                        target = str(Path(target).parent / f'{stem}_{new_rid}{ext}')
                    out_files[dst_key] = p_files[src_key]
                    if dst_key not in out_names:
                        out_names.append(dst_key)
                    ext_lower = Path(dst_key).suffix.lower()
                    if _CONTENT_TYPES_KEY in out_files:
                        out_files[_CONTENT_TYPES_KEY] = _ensure_content_type(
                            out_files[_CONTENT_TYPES_KEY], ext_lower
                        )

            target_mode = ' TargetMode="External"' if is_external else ''
            rel_entry = (
                f'<Relationship Id="{new_rid}" Type="{rel["type"]}" Target="{target}"{target_mode}/>'
            ).encode()
            if _RELS_KEY in out_files:
                out_files[_RELS_KEY] = out_files[_RELS_KEY].replace(
                    b'</Relationships>', rel_entry + b'</Relationships>'
                )

        # Parse partial body, collect needed numIds, merge numbering
        p_root = ET.fromstring(p_doc_bytes)
        p_body = p_root.find(f'{{{W}}}body')
        p_paras = [el for el in list(p_body) if el.tag != f'{{{W}}}sectPr']

        p_num_ids = _collect_num_ids(p_paras)
        num_id_map = _merge_numbering(out_files, p_files, p_num_ids)

        # Apply remapping directly on element tree
        if rid_map:
            _remap_rids(p_root, rid_map)
            p_paras = [el for el in list(p_body) if el.tag != f'{{{W}}}sectPr']
        if num_id_map:
            _remap_numids(p_root, num_id_map)
            p_paras = [el for el in list(p_body) if el.tag != f'{{{W}}}sectPr']

        # Strip paragraph styles and inline run properties so the content renders
        # correctly regardless of what styles the output template defines.
        _inline_partial_styles(
            p_paras,
            p_files.get('word/styles.xml', b''),
            p_files.get('word/theme/theme1.xml', b''),
        )

        # Replace each occurrence of the placeholder paragraph
        while True:
            found = False
            for i, el in enumerate(list(body)):
                if el.tag != f'{{{W}}}p':
                    continue
                text = ''.join(t.text or '' for t in el.findall(f'.//{{{W}}}t')).strip()
                if text == placeholder:
                    body.remove(el)
                    for j, para in enumerate(p_paras):
                        body.insert(i + j, para)
                    modified = True
                    found = True
                    break
            if not found:
                break

    if not modified:
        return

    new_xml = ET.tostring(root, encoding='unicode')
    out_files['word/document.xml'] = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        + new_xml.encode('utf-8')
    )
    _write_zip(docx_path, out_names, out_files)


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


def _build_rpr(style: RunStyle) -> ET.Element:
    """Build a fresh w:rPr element from a RunStyle."""
    rpr = ET.Element(f'{{{W}}}rPr')
    if style.font:
        fonts = ET.SubElement(rpr, f'{{{W}}}rFonts')
        for attr in ('ascii', 'hAnsi', 'cs', 'eastAsia'):
            fonts.set(f'{{{W}}}{attr}', style.font)
    if style.bold:
        ET.SubElement(rpr, f'{{{W}}}b')
    if style.italic:
        ET.SubElement(rpr, f'{{{W}}}i')
    if style.color:
        ET.SubElement(rpr, f'{{{W}}}color').set(f'{{{W}}}val', style.color)
    if style.size_half_pt:
        ET.SubElement(rpr, f'{{{W}}}sz').set(f'{{{W}}}val', str(style.size_half_pt))
        ET.SubElement(rpr, f'{{{W}}}szCs').set(f'{{{W}}}val', str(style.size_half_pt))
    return rpr


def _merge_para_indent(ppr: ET.Element, style: RunStyle) -> bool:
    """Apply indentation from style to a pPr element. Returns True if changed."""
    if style.ind_left is None and style.ind_hanging is None:
        return False
    ind_el = ppr.find(f'{{{W}}}ind')
    if ind_el is None:
        ind_el = ET.SubElement(ppr, f'{{{W}}}ind')
    if style.ind_left is not None:
        ind_el.set(f'{{{W}}}left', str(style.ind_left))
    if style.ind_hanging is not None:
        ind_el.set(f'{{{W}}}hanging', str(style.ind_hanging))
    return True


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
    """Post-process output DOCX to apply template run styles to bold, italic, and code runs."""
    italic_style = style_map.get("Italic text")
    bold_style = style_map.get("Bold text")
    code_style = style_map.get("Code block text")
    bullet_style = style_map.get("Bullet point text")
    if not italic_style and not bold_style and not code_style and not bullet_style:
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

        is_code_para = para_style == 'SourceCode'
        ppr_el = p.find(f'{{{W}}}pPr')
        is_bullet_para = bullet_style and ppr_el is not None and ppr_el.find(f'{{{W}}}numPr') is not None

        for r in p.findall(f'{{{W}}}r'):
            rpr = r.find(f'{{{W}}}rPr')
            if rpr is None:
                if is_code_para and code_style:
                    rpr = ET.Element(f'{{{W}}}rPr')
                    r.insert(0, rpr)  # rPr must be first child of w:r
                    _merge_run_style(rpr, code_style)
                    modified = True
                elif is_bullet_para:
                    rpr = ET.Element(f'{{{W}}}rPr')
                    r.insert(0, rpr)  # rPr must be first child of w:r
                    _merge_run_style(rpr, bullet_style)
                    modified = True
                continue
            rStyle_el = rpr.find(f'{{{W}}}rStyle')
            is_verbatim = (
                rStyle_el is not None
                and rStyle_el.get(f'{{{W}}}val') == 'VerbatimChar'
            )
            if (is_verbatim or is_code_para) and code_style:
                _merge_run_style(rpr, code_style)
                modified = True
            elif is_bullet_para:
                _merge_run_style(rpr, bullet_style)
                modified = True
            elif rpr.find(f'{{{W}}}i') is not None and italic_style:
                _merge_run_style(rpr, italic_style)
                modified = True
            elif rpr.find(f'{{{W}}}b') is not None and bold_style:
                _merge_run_style(rpr, bold_style)
                modified = True

        if is_bullet_para and _merge_para_indent(ppr_el, bullet_style):
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


_RT_MARKER = ''


def inject_right_tab_markers(text: str) -> str:
    """Replace >> with a sentinel so pandoc does not interpret it as a blockquote.

    Convention: left text >> right text
    Any >> on a line becomes a right-aligned tab stop separator.
    Nothing to the left of >> is valid (purely right-aligned).
    """
    lines = text.split('\n')
    out = []
    for line in lines:
        if '>>' in line:
            line = re.sub(r'>>\s*', _RT_MARKER, line, count=1)
        out.append(line)
    return '\n'.join(out)


def apply_right_tab_stops(docx_path: Path, tab_stop: dict) -> None:
    """Split runs at RT_MARKER, insert <w:tab/>, and apply tab stop definition to pPr."""
    with zipfile.ZipFile(docx_path, 'r') as zin:
        names = zin.namelist()
        files = {name: zin.read(name) for name in names}

    doc_bytes = files['word/document.xml']
    for m in re.finditer(rb'xmlns:(\w+)="([^"]+)"', doc_bytes):
        ET.register_namespace(m.group(1).decode(), m.group(2).decode())

    root = ET.fromstring(doc_bytes)
    modified = False

    for p in root.findall(f'.//{{{W}}}p'):
        all_text = ''.join(t.text or '' for t in p.findall(f'.//{{{W}}}t'))
        if _RT_MARKER not in all_text:
            continue

        for r in list(p.findall(f'{{{W}}}r')):
            t_el = r.find(f'{{{W}}}t')
            if t_el is None or not t_el.text or _RT_MARKER not in t_el.text:
                continue

            left_text, right_text = t_el.text.split(_RT_MARKER, 1)
            rpr = r.find(f'{{{W}}}rPr')
            right_style = tab_stop.get('right_style')
            run_idx = list(p).index(r)
            p.remove(r)

            def _make_run(txt, use_rpr=None):
                rn = ET.Element(f'{{{W}}}r')
                if use_rpr is not None:
                    rn.append(ET.fromstring(ET.tostring(use_rpr)))
                te = ET.SubElement(rn, f'{{{W}}}t')
                te.text = txt
                if txt and (txt[0] == ' ' or txt[-1] == ' '):
                    te.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                return rn

            right_rpr = _build_rpr(right_style) if right_style else rpr

            # Insert in reverse order at run_idx so final order is left, tab, right
            if right_text:
                p.insert(run_idx, _make_run(right_text, right_rpr))
            r_tab = ET.Element(f'{{{W}}}r')
            if rpr is not None:
                r_tab.append(ET.fromstring(ET.tostring(rpr)))
            ET.SubElement(r_tab, f'{{{W}}}tab')
            p.insert(run_idx, r_tab)
            if left_text:
                p.insert(run_idx, _make_run(left_text))

            # Inject tab stop into paragraph pPr
            ppr_el = p.find(f'{{{W}}}pPr')
            if ppr_el is None:
                ppr_el = ET.SubElement(p, f'{{{W}}}pPr')
            tabs_el = ppr_el.find(f'{{{W}}}tabs')
            if tabs_el is None:
                tabs_el = ET.SubElement(ppr_el, f'{{{W}}}tabs')
            tab_el = ET.SubElement(tabs_el, f'{{{W}}}tab')
            tab_el.set(f'{{{W}}}val', tab_stop['val'])
            tab_el.set(f'{{{W}}}pos', tab_stop['pos'])
            tab_el.set(f'{{{W}}}leader', tab_stop['leader'])

            modified = True
            break

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


def inject_title_styles(text: str, para_styles: set[str]) -> str:
    """Replace % and %% prefix lines with pandoc custom-style fenced divs.

    Convention:
      %<text>   → Title paragraph  (any line starting with a single %)
      %%<text>  → Subtitle paragraph (any line starting with %%)
    Each can appear anywhere in the document, independently of the other.
    %% is tested first so it is not misread as a single-% line.
    Silently passes through if the template does not define the style.
    """
    has_title = "Title" in para_styles
    has_subtitle = "Subtitle" in para_styles
    if not has_title and not has_subtitle:
        return text

    out = []
    for line in text.split('\n'):
        if has_subtitle and line.startswith('%%'):
            out += ['::: {custom-style="Subtitle"}', line[2:].lstrip(), ':::']
        elif has_title and line.startswith('%'):
            out += ['::: {custom-style="Title"}', line[1:].lstrip(), ':::']
        else:
            out.append(line)
    return '\n'.join(out)


_EXTENSION_CONSTRUCTS = [
    (r'\*\*\S', "Bold text", "**bold**"),
    (r'(?<!\*)\*(?!\*)\S', "Italic text", "*italic*"),
    (r'`', "Code block text", "`code`"),
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
    parser.add_argument(
        "--partial", metavar="NAME=FILE.docx", action="append", default=[],
        help="Splice partial DOCX at {{NAME}} placeholder (repeatable)",
    )
    args = parser.parse_args()

    if not args.content.exists():
        sys.exit(f"Error: content file not found: {args.content}")
    if not args.template.exists():
        sys.exit(f"Error: template file not found: {args.template}")
    if shutil.which("pandoc") is None:
        sys.exit("Error: pandoc not found. Install with: brew install pandoc")

    partials: dict[str, Path] = {}
    for spec in args.partial:
        if '=' not in spec:
            sys.exit(f"Error: --partial must be NAME=FILE.docx, got: {spec!r}")
        pname, ppath_str = spec.split('=', 1)
        ppath = Path(ppath_str)
        if not ppath.exists():
            sys.exit(f"Error: partial file not found: {ppath}")
        partials[pname] = ppath

    issues = validate_template(args.template)
    if issues:
        details = "\n".join(f"  • {issue}" for issue in issues)
        sys.exit(
            f"Error: style template is incomplete:\n{details}\n"
            "See README — Style mapping for template requirements."
        )

    style_map = read_style_map(args.template)
    right_tab = read_right_tab_stop(args.template)
    para_styles = read_paragraph_style_ids(args.template)

    raw = args.content.read_text()
    check_required_styles(raw, style_map)
    raw = inject_title_styles(raw, para_styles)
    raw = inject_page_break_markers(raw)
    if right_tab:
        raw = inject_right_tab_markers(raw)
    modified = restore_soft_newlines(inject_blank_paragraphs(mark_soft_newlines(raw)))
    modified = modified.replace(_PAGE_BREAK_MARKER, _PAGE_BREAK_XML)

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
    if right_tab:
        apply_right_tab_stops(args.output, right_tab)
    apply_run_styles(args.output, style_map)
    splice_partials(args.output, partials)
    print(f"Generated: {args.output}")


if __name__ == "__main__":
    main()
