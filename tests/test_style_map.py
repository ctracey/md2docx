"""Tests for style_map.py — StyleMap reader and RunStyle extraction."""

import io
import zipfile
from pathlib import Path

import pytest

from md2docx.style_map import KNOWN_LABELS, RunStyle, StyleMap, read_style_map

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


# ---------------------------------------------------------------------------
# Minimal DOCX builder
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
# read_style_map
# ---------------------------------------------------------------------------


def test_reads_known_label(tmp_path):
    tmpl = _make_docx([_para_xml("Bold text", bold=True)], tmp_path)
    assert "Bold text" in read_style_map(tmpl)


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
    assert "Italic text" not in read_style_map(tmpl)


def test_raises_on_duplicate_label(tmp_path):
    tmpl = _make_docx([
        _para_xml("Italic text", italic=True),
        _para_xml("Italic text", italic=True, font="Other Font"),
    ], tmp_path)
    with pytest.raises(ValueError, match="Italic text"):
        read_style_map(tmpl)


def test_all_known_labels_constant():
    assert "Bold text" in KNOWN_LABELS
    assert "Italic text" in KNOWN_LABELS
    assert "Code block text" in KNOWN_LABELS
    assert "Bullet point text" in KNOWN_LABELS
    assert "Normal text" not in KNOWN_LABELS
    assert "Title text" not in KNOWN_LABELS


# ---------------------------------------------------------------------------
# StyleMap.from_template
# ---------------------------------------------------------------------------


def test_style_map_from_template_returns_stylemap(template):
    sm = StyleMap.from_template(template)
    assert isinstance(sm, StyleMap)


def test_style_map_run_styles_populated(template):
    sm = StyleMap.from_template(template)
    assert isinstance(sm.run_styles, dict)
    assert len(sm.run_styles) > 0


def test_style_map_para_styles_is_set(template):
    sm = StyleMap.from_template(template)
    assert isinstance(sm.para_styles, set)
    assert len(sm.para_styles) > 0


def test_style_map_right_tab_present_in_sample_template(template):
    sm = StyleMap.from_template(template)
    assert sm.right_tab is not None
