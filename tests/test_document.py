"""Tests for document.py — DOCX post-processing functions."""

import io
import shutil
import subprocess
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import pytest

from md2docx.content import _RT_MARKER, inject_right_tab_markers, mark_soft_newlines, inject_blank_paragraphs, restore_soft_newlines
from md2docx.document import apply_run_styles, apply_right_tab_stops, sync_headers
from md2docx.style_map import RunStyle, read_right_tab_stop
from tests.conftest import run_generator

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _pandoc_only(md, template, out):
    subprocess.run(
        [shutil.which("pandoc"), str(md), "--reference-doc", str(template), "-o", str(out)],
        check=True,
    )


def _generate(md, template, out):
    run_generator([str(md), str(template), str(out)])


def _italic_runs(docx_path: Path) -> list[ET.Element]:
    with zipfile.ZipFile(str(docx_path)) as zf:
        root = ET.parse(zf.open('word/document.xml')).getroot()
    return [
        r for r in root.findall(f'.//{{{W}}}r')
        if r.find(f'{{{W}}}rPr/{{{W}}}i') is not None
    ]


def _bold_runs(docx_path: Path) -> list[ET.Element]:
    with zipfile.ZipFile(str(docx_path)) as zf:
        root = ET.parse(zf.open('word/document.xml')).getroot()
    return [
        r for r in root.findall(f'.//{{{W}}}r')
        if r.find(f'{{{W}}}rPr/{{{W}}}b') is not None
    ]


# ---------------------------------------------------------------------------
# sync_headers
# ---------------------------------------------------------------------------


def test_headers_copied_from_template_when_present(simple_md, tmp_path):
    HEADER_XML = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        b'<w:p><w:r><w:t>TEMPLATE HEADER</w:t></w:r></w:p>'
        b'</w:hdr>'
    )
    DOC_XML = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        b' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        b'<w:body><w:p><w:r><w:t>body</w:t></w:r></w:p>'
        b'<w:sectPr>'
        b'<w:headerReference r:id="rId2" w:type="default"/>'
        b'</w:sectPr></w:body></w:document>'
    )
    RELS_XML = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        b'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        b'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
        b'</Relationships>'
    )
    CT_XML = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        b'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        b'<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        b'</Types>'
    )

    tmpl_path = tmp_path / "tmpl_with_header.docx"
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w') as zf:
        zf.writestr('word/document.xml', DOC_XML)
        zf.writestr('word/header1.xml', HEADER_XML)
        zf.writestr('word/_rels/document.xml.rels', RELS_XML)
        zf.writestr('[Content_Types].xml', CT_XML)
        zf.writestr('_rels/.rels', b'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>')
    tmpl_path.write_bytes(buf.getvalue())

    out = tmp_path / "out.docx"
    out.write_bytes(buf.getvalue())

    sync_headers(out, tmpl_path)

    with zipfile.ZipFile(str(out)) as zf:
        names = zf.namelist()
        out_doc = zf.read('word/document.xml')
        header_content = zf.read('word/header1.xml')

    assert 'word/header1.xml' in names
    assert b'TEMPLATE HEADER' in header_content
    assert b'headerReference' in out_doc


# ---------------------------------------------------------------------------
# apply_run_styles
# ---------------------------------------------------------------------------


def test_apply_adds_font_to_italic_runs(inline_formatting_md, template, tmp_path):
    out = tmp_path / "out.docx"
    _pandoc_only(inline_formatting_md, template, out)
    style_map = {"Italic text": RunStyle(font="TestFont", size_half_pt=20, color="FF0000", italic=True)}
    apply_run_styles(out, style_map)
    for r in _italic_runs(out):
        fonts = r.find(f'{{{W}}}rPr/{{{W}}}rFonts')
        assert fonts is not None
        assert fonts.get(f'{{{W}}}ascii') == "TestFont"


def test_apply_adds_font_to_bold_runs(inline_formatting_md, template, tmp_path):
    out = tmp_path / "out.docx"
    _pandoc_only(inline_formatting_md, template, out)
    style_map = {"Bold text": RunStyle(font="BoldFont", bold=True)}
    apply_run_styles(out, style_map)
    for r in _bold_runs(out):
        fonts = r.find(f'{{{W}}}rPr/{{{W}}}rFonts')
        assert fonts is not None
        assert fonts.get(f'{{{W}}}ascii') == "BoldFont"


def test_apply_does_not_modify_heading_runs(headings_md, template, tmp_path):
    out = tmp_path / "out.docx"
    _generate(headings_md, template, out)
    style_map = {"Italic text": RunStyle(font="ShouldNotAppear", italic=True)}
    apply_run_styles(out, style_map)
    with zipfile.ZipFile(str(out)) as zf:
        root = ET.parse(zf.open('word/document.xml')).getroot()
    for p in root.findall(f'.//{{{W}}}p'):
        style_el = p.find(f'.//{{{W}}}pPr/{{{W}}}pStyle')
        if style_el is None:
            continue
        if 'Heading' in style_el.get(f'{{{W}}}val', ''):
            for r in p.findall(f'{{{W}}}r'):
                fonts = r.find(f'{{{W}}}rPr/{{{W}}}rFonts')
                if fonts is not None:
                    assert fonts.get(f'{{{W}}}ascii') != "ShouldNotAppear"


def test_apply_adds_font_to_inline_code_runs(code_blocks_md, template, tmp_path):
    out = tmp_path / "out.docx"
    _pandoc_only(code_blocks_md, template, out)
    style_map = {"Code block text": RunStyle(font="CodeFont", color="333333")}
    apply_run_styles(out, style_map)
    with zipfile.ZipFile(str(out)) as zf:
        root = ET.parse(zf.open('word/document.xml')).getroot()
    verbatim_runs = [
        r for r in root.findall(f'.//{{{W}}}r')
        if r.find(f'{{{W}}}rPr/{{{W}}}rStyle') is not None
        and r.find(f'{{{W}}}rPr/{{{W}}}rStyle').get(f'{{{W}}}val') == 'VerbatimChar'
    ]
    assert verbatim_runs
    for r in verbatim_runs:
        fonts = r.find(f'{{{W}}}rPr/{{{W}}}rFonts')
        assert fonts is not None
        assert fonts.get(f'{{{W}}}ascii') == "CodeFont"


def test_apply_adds_font_to_code_block_runs(code_blocks_md, template, tmp_path):
    out = tmp_path / "out.docx"
    _pandoc_only(code_blocks_md, template, out)
    style_map = {"Code block text": RunStyle(font="CodeFont")}
    apply_run_styles(out, style_map)
    with zipfile.ZipFile(str(out)) as zf:
        root = ET.parse(zf.open('word/document.xml')).getroot()
    source_paras = [
        p for p in root.findall(f'.//{{{W}}}p')
        if p.find(f'.//{{{W}}}pStyle') is not None
        and p.find(f'.//{{{W}}}pStyle').get(f'{{{W}}}val') == 'SourceCode'
    ]
    assert source_paras
    for p in source_paras:
        for r in p.findall(f'{{{W}}}r'):
            fonts = r.find(f'{{{W}}}rPr/{{{W}}}rFonts')
            assert fonts is not None
            assert fonts.get(f'{{{W}}}ascii') == "CodeFont"


def test_apply_adds_font_to_bullet_runs(simple_md, template, tmp_path):
    out = tmp_path / "out.docx"
    _pandoc_only(simple_md, template, out)
    style_map = {"Bullet point text": RunStyle(font="BulletFont", color="444444")}
    apply_run_styles(out, style_map)
    with zipfile.ZipFile(str(out)) as zf:
        root = ET.parse(zf.open('word/document.xml')).getroot()
    bullet_paras = [
        p for p in root.findall(f'.//{{{W}}}p')
        if p.find(f'{{{W}}}pPr/{{{W}}}numPr') is not None
    ]
    assert bullet_paras
    for p in bullet_paras:
        for r in p.findall(f'{{{W}}}r'):
            fonts = r.find(f'{{{W}}}rPr/{{{W}}}rFonts')
            assert fonts is not None
            assert fonts.get(f'{{{W}}}ascii') == "BulletFont"


def test_apply_noop_when_style_map_empty(inline_formatting_md, template, tmp_path):
    out = tmp_path / "out.docx"
    _generate(inline_formatting_md, template, out)
    mtime_before = out.stat().st_mtime
    apply_run_styles(out, {})
    assert out.stat().st_mtime == mtime_before


def test_bullet_rpr_is_first_child_of_run(simple_md, template, tmp_path):
    out = tmp_path / "out.docx"
    _pandoc_only(simple_md, template, out)
    style_map = {"Bullet point text": RunStyle(font="BulletFont")}
    apply_run_styles(out, style_map)
    with zipfile.ZipFile(str(out)) as zf:
        root = ET.parse(zf.open('word/document.xml')).getroot()
    for p in root.findall(f'.//{{{W}}}p'):
        if p.find(f'.//{{{W}}}numPr') is None:
            continue
        for r in p.findall(f'{{{W}}}r'):
            children = list(r)
            tags = [c.tag.split('}')[-1] for c in children]
            if 'rPr' in tags and 't' in tags:
                assert tags.index('rPr') < tags.index('t')


def test_code_rpr_is_first_child_of_run(code_blocks_md, template, tmp_path):
    out = tmp_path / "out.docx"
    _pandoc_only(code_blocks_md, template, out)
    style_map = {"Code block text": RunStyle(font="CodeFont")}
    apply_run_styles(out, style_map)
    with zipfile.ZipFile(str(out)) as zf:
        root = ET.parse(zf.open('word/document.xml')).getroot()
    for p in root.findall(f'.//{{{W}}}p'):
        ppr = p.find(f'.//{{{W}}}pStyle')
        if ppr is None or ppr.get(f'{{{W}}}val') != 'SourceCode':
            continue
        for r in p.findall(f'{{{W}}}r'):
            children = list(r)
            tags = [c.tag.split('}')[-1] for c in children]
            if 'rPr' in tags and 't' in tags:
                assert tags.index('rPr') < tags.index('t')


def test_bullet_style_applied_after_right_tab_processing(template, tmp_path):
    right_tab = read_right_tab_stop(template)
    assert right_tab is not None, "Template must define RightAlignedTabStop for this test"

    raw = "Normal text >> right side\n\n- bullet one\n- bullet two\n\n"
    preprocessed = restore_soft_newlines(
        inject_blank_paragraphs(mark_soft_newlines(inject_right_tab_markers(raw)))
    )
    assert _RT_MARKER in preprocessed

    md = tmp_path / "content.md"
    md.write_text(preprocessed)
    out = tmp_path / "out.docx"

    subprocess.run(
        [shutil.which("pandoc"), str(md), "--reference-doc", str(template), "-o", str(out)],
        check=True,
    )

    apply_right_tab_stops(out, right_tab)

    style_map = {"Bullet point text": RunStyle(font="BulletFont")}
    apply_run_styles(out, style_map)

    with zipfile.ZipFile(str(out)) as zf:
        root = ET.parse(zf.open('word/document.xml')).getroot()
    bullet_paras = [
        p for p in root.findall(f'.//{{{W}}}p')
        if p.find(f'.//{{{W}}}numPr') is not None
    ]
    assert bullet_paras
    for p in bullet_paras:
        for r in p.findall(f'{{{W}}}r'):
            fonts = r.find(f'{{{W}}}rPr/{{{W}}}rFonts')
            assert fonts is not None
            assert fonts.get(f'{{{W}}}ascii') == "BulletFont"
