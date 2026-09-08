"""Tests for style_map.py — StyleMap reader and run-style application."""

import io
import shutil
import subprocess
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import pytest

from style_map import KNOWN_LABELS, RunStyle, read_style_map
from generate import apply_run_styles
from tests.conftest import run_generator

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# ---------------------------------------------------------------------------
# Minimal DOCX builder for controlled tests
# ---------------------------------------------------------------------------

_CONTENT_TYPES = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

_RELS = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>"""


def _para_xml(text: str, bold: bool = False, italic: bool = False,
              font: str | None = None, size_half_pt: int | None = None,
              color: str | None = None) -> str:
    parts = []
    if bold:
        parts.append('<w:b/>')
    if italic:
        parts.append('<w:i/>')
    if font:
        parts.append(f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}"/>')
    if size_half_pt:
        parts.append(f'<w:sz w:val="{size_half_pt}"/>')
    if color:
        parts.append(f'<w:color w:val="{color}"/>')
    rpr = f'<w:rPr>{"".join(parts)}</w:rPr>' if parts else ''
    return f'<w:p><w:r>{rpr}<w:t>{text}</w:t></w:r></w:p>'


def _make_docx(paragraphs_xml: list[str], tmp_path: Path) -> Path:
    body = ''.join(paragraphs_xml) + '<w:sectPr/>'
    doc_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:body>{body}</w:body>'
        '</w:document>'
    )
    path = tmp_path / "template.docx"
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w') as zf:
        zf.writestr('word/document.xml', doc_xml)
        zf.writestr('[Content_Types].xml', _CONTENT_TYPES)
        zf.writestr('_rels/.rels', _RELS)
    path.write_bytes(buf.getvalue())
    return path


# ---------------------------------------------------------------------------
# read_style_map — unit tests
# ---------------------------------------------------------------------------


def test_reads_known_label(tmp_path):
    tmpl = _make_docx([_para_xml("Bold text", bold=True)], tmp_path)
    result = read_style_map(tmpl)
    assert "Bold text" in result


def test_extracts_bold_flag(tmp_path):
    tmpl = _make_docx([_para_xml("Bold text", bold=True)], tmp_path)
    assert read_style_map(tmpl)["Bold text"].bold is True


def test_extracts_italic_flag(tmp_path):
    tmpl = _make_docx([_para_xml("Italic text", italic=True)], tmp_path)
    assert read_style_map(tmpl)["Italic text"].italic is True


def test_extracts_font(tmp_path):
    tmpl = _make_docx([_para_xml("Italic text", italic=True, font="DM Sans ExtraLight")], tmp_path)
    assert read_style_map(tmpl)["Italic text"].font == "DM Sans ExtraLight"


def test_extracts_size(tmp_path):
    tmpl = _make_docx([_para_xml("Italic text", italic=True, size_half_pt=16)], tmp_path)
    style = read_style_map(tmpl)["Italic text"]
    assert style.size_half_pt == 16
    assert style.size_pt == 8.0


def test_extracts_color(tmp_path):
    tmpl = _make_docx([_para_xml("Italic text", italic=True, color="666666")], tmp_path)
    assert read_style_map(tmpl)["Italic text"].color == "666666"


def test_unknown_label_ignored(tmp_path):
    tmpl = _make_docx([_para_xml("Something else entirely")], tmp_path)
    assert read_style_map(tmpl) == {}


def test_absent_label_not_in_result(tmp_path):
    tmpl = _make_docx([_para_xml("Bold text", bold=True)], tmp_path)
    result = read_style_map(tmpl)
    assert "Italic text" not in result


def test_raises_on_duplicate_label(tmp_path):
    tmpl = _make_docx([
        _para_xml("Italic text", italic=True),
        _para_xml("Italic text", italic=True, font="Other Font"),
    ], tmp_path)
    with pytest.raises(ValueError, match="Italic text"):
        read_style_map(tmpl)


def test_no_rpr_returns_empty_run_style(tmp_path):
    tmpl = _make_docx([_para_xml("Normal text")], tmp_path)
    style = read_style_map(tmpl)["Normal text"]
    assert style == RunStyle()


def test_all_known_labels_constant():
    assert "Normal text" in KNOWN_LABELS
    assert "Bold text" in KNOWN_LABELS
    assert "Italic text" in KNOWN_LABELS


# ---------------------------------------------------------------------------
# apply_run_styles — integration tests
# ---------------------------------------------------------------------------


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


def _generate(md, template, out):
    run_generator([str(md), str(template), str(out)])


def _pandoc_only(md, template, out):
    """Call pandoc directly — bypasses our pipeline so apply_run_styles hasn't run yet."""
    subprocess.run(
        [shutil.which("pandoc"), str(md), "--reference-doc", str(template), "-o", str(out)],
        check=True,
    )


def test_apply_adds_font_to_italic_runs(inline_formatting_md, template, tmp_path):
    out = tmp_path / "out.docx"
    _pandoc_only(inline_formatting_md, template, out)

    style_map = {"Italic text": RunStyle(font="TestFont", size_half_pt=20, color="FF0000", italic=True)}
    apply_run_styles(out, style_map)

    for r in _italic_runs(out):
        fonts = r.find(f'{{{W}}}rPr/{{{W}}}rFonts')
        assert fonts is not None, "rFonts not added to italic run"
        assert fonts.get(f'{{{W}}}ascii') == "TestFont"


def test_apply_adds_font_to_bold_runs(inline_formatting_md, template, tmp_path):
    out = tmp_path / "out.docx"
    _pandoc_only(inline_formatting_md, template, out)

    style_map = {"Bold text": RunStyle(font="BoldFont", bold=True)}
    apply_run_styles(out, style_map)

    for r in _bold_runs(out):
        fonts = r.find(f'{{{W}}}rPr/{{{W}}}rFonts')
        assert fonts is not None, "rFonts not added to bold run"
        assert fonts.get(f'{{{W}}}ascii') == "BoldFont"


def test_apply_does_not_modify_heading_runs(headings_md, template, tmp_path):
    """Run styles must not touch heading paragraphs."""
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


def test_apply_noop_when_style_map_empty(inline_formatting_md, template, tmp_path):
    out = tmp_path / "out.docx"
    _generate(inline_formatting_md, template, out)
    mtime_before = out.stat().st_mtime
    apply_run_styles(out, {})
    assert out.stat().st_mtime == mtime_before, "File must not be modified for empty style map"
