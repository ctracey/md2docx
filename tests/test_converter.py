"""Tests for converter.py — end-to-end pipeline and CLI."""

import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import pytest

from md2docx.converter import convert
from md2docx.exceptions import ConversionError
from tests.conftest import run_generator

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def run(args: list[str]) -> subprocess.CompletedProcess:
    return run_generator(args)


def _document_xml(path: Path) -> ET.Element:
    with zipfile.ZipFile(str(path)) as zf:
        with zf.open("word/document.xml") as f:
            return ET.parse(f).getroot()


def _paragraph_styles(path: Path) -> list[str]:
    root = _document_xml(path)
    styles = []
    for p in root.findall(f".//{{{W}}}p"):
        style_el = p.find(f".//{{{W}}}pStyle")
        styles.append(style_el.get(f"{{{W}}}val") if style_el is not None else "Normal")
    return styles


def _paragraph_texts(path: Path) -> list[str]:
    root = _document_xml(path)
    texts = []
    for p in root.findall(f".//{{{W}}}p"):
        runs = p.findall(f".//{{{W}}}t")
        texts.append("".join(t.text or "" for t in runs))
    return texts


def _has_real_bullets(path: Path) -> bool:
    root = _document_xml(path)
    return any(
        p.find(f".//{{{W}}}numPr") is not None
        for p in root.findall(f".//{{{W}}}p")
    )


def _has_bold_run(path: Path) -> bool:
    root = _document_xml(path)
    return any(
        r.find(f"{{{W}}}rPr/{{{W}}}b") is not None
        for r in root.findall(f".//{{{W}}}r")
    )


def _has_italic_run(path: Path) -> bool:
    root = _document_xml(path)
    return any(
        r.find(f"{{{W}}}rPr/{{{W}}}i") is not None
        for r in root.findall(f".//{{{W}}}r")
    )


# ---------------------------------------------------------------------------
# convert() — library interface
# ---------------------------------------------------------------------------


def test_convert_raises_on_missing_content(template, tmp_path):
    with pytest.raises(ConversionError, match="content file not found"):
        convert(Path("nonexistent.md"), template, tmp_path / "out.docx")


def test_convert_raises_on_missing_template(simple_md, tmp_path):
    with pytest.raises(ConversionError, match="template file not found"):
        convert(simple_md, Path("nonexistent.docx"), tmp_path / "out.docx")


def test_convert_produces_output_file(simple_md, template, tmp_path):
    out = tmp_path / "out.docx"
    convert(simple_md, template, out)
    assert out.exists()


def test_convert_output_is_valid_ooxml(simple_md, template, tmp_path):
    out = tmp_path / "out.docx"
    convert(simple_md, template, out)
    with zipfile.ZipFile(out) as zf:
        names = zf.namelist()
    assert "word/document.xml" in names
    assert "[Content_Types].xml" in names


def test_convert_does_not_print(simple_md, template, tmp_path, capsys):
    convert(simple_md, template, tmp_path / "out.docx")
    captured = capsys.readouterr()
    assert captured.out == ""
    assert captured.err == ""


# ---------------------------------------------------------------------------
# CLI — error handling
# ---------------------------------------------------------------------------


def test_missing_content_file(template, tmp_path):
    result = run(["nonexistent.md", str(template), str(tmp_path / "out.docx")])
    assert result.returncode == 1
    assert "content file not found" in result.stderr


def test_missing_template_file(simple_md, tmp_path):
    result = run([str(simple_md), "nonexistent.docx", str(tmp_path / "out.docx")])
    assert result.returncode == 1
    assert "template file not found" in result.stderr


# ---------------------------------------------------------------------------
# CLI — successful generation
# ---------------------------------------------------------------------------


def test_creates_output_file(simple_md, template, tmp_path):
    out = tmp_path / "out.docx"
    result = run([str(simple_md), str(template), str(out)])
    assert result.returncode == 0, result.stderr
    assert out.exists()


def test_prints_generated_path(simple_md, template, tmp_path):
    out = tmp_path / "out.docx"
    result = run([str(simple_md), str(template), str(out)])
    assert "Generated:" in result.stdout
    assert str(out) in result.stdout


def test_no_headers_when_template_has_none(simple_md, template, tmp_path):
    out = tmp_path / "out.docx"
    run([str(simple_md), str(template), str(out)])
    with zipfile.ZipFile(str(out)) as zf:
        names = zf.namelist()
        doc_xml = zf.read("word/document.xml")
    assert not any("header" in n.lower() for n in names)
    assert b"headerReference" not in doc_xml


def test_no_bookmarks_in_output(headings_md, template, tmp_path):
    out = tmp_path / "out.docx"
    run([str(headings_md), str(template), str(out)])
    with zipfile.ZipFile(str(out)) as zf:
        doc_xml = zf.read("word/document.xml")
    assert b"bookmarkStart" not in doc_xml
    assert b"bookmarkEnd" not in doc_xml


def test_output_is_valid_ooxml(simple_md, template, tmp_path):
    out = tmp_path / "out.docx"
    run([str(simple_md), str(template), str(out)])
    with zipfile.ZipFile(str(out)) as zf:
        names = zf.namelist()
    assert "word/document.xml" in names
    assert "[Content_Types].xml" in names
    assert "word/styles.xml" in names


# ---------------------------------------------------------------------------
# Markdown → DOCX mapping
# ---------------------------------------------------------------------------


def test_heading1_style(headings_md, template, tmp_path):
    out = tmp_path / "out.docx"
    run([str(headings_md), str(template), str(out)])
    assert any("Heading1" in s for s in _paragraph_styles(out))


def test_heading2_style(headings_md, template, tmp_path):
    out = tmp_path / "out.docx"
    run([str(headings_md), str(template), str(out)])
    assert any("Heading2" in s for s in _paragraph_styles(out))


def test_heading3_style(headings_md, template, tmp_path):
    out = tmp_path / "out.docx"
    run([str(headings_md), str(template), str(out)])
    assert any("Heading3" in s for s in _paragraph_styles(out))


def test_heading4_style(headings_md, template, tmp_path):
    out = tmp_path / "out.docx"
    run([str(headings_md), str(template), str(out)])
    assert any("Heading4" in s for s in _paragraph_styles(out))


def test_heading5_style(headings_md, template, tmp_path):
    out = tmp_path / "out.docx"
    run([str(headings_md), str(template), str(out)])
    assert any("Heading5" in s for s in _paragraph_styles(out))


def test_heading6_style(headings_md, template, tmp_path):
    out = tmp_path / "out.docx"
    run([str(headings_md), str(template), str(out)])
    assert any("Heading6" in s for s in _paragraph_styles(out))


def test_bullets_are_real_word_lists(simple_md, template, tmp_path):
    out = tmp_path / "out.docx"
    run([str(simple_md), str(template), str(out)])
    assert _has_real_bullets(out)


def test_bullet_text_present(simple_md, template, tmp_path):
    out = tmp_path / "out.docx"
    run([str(simple_md), str(template), str(out)])
    texts = _paragraph_texts(out)
    assert any("First bullet item" in t for t in texts)
    assert any("Second bullet item" in t for t in texts)
    assert any("Third bullet item" in t for t in texts)


def test_plain_paragraphs_present(simple_md, template, tmp_path):
    out = tmp_path / "out.docx"
    run([str(simple_md), str(template), str(out)])
    texts = _paragraph_texts(out)
    assert any("plain paragraph" in t for t in texts)


# ---------------------------------------------------------------------------
# Blank line rendering
# ---------------------------------------------------------------------------


def test_blank_lines_between_paragraphs_produce_empty_paragraphs(simple_md, template, tmp_path):
    out = tmp_path / "out.docx"
    run([str(simple_md), str(template), str(out)])
    texts = _paragraph_texts(out)
    assert any(t.strip() == "" for t in texts)


def test_two_leading_blank_lines_produce_two_empty_paragraphs(leading_blank_md, template, tmp_path):
    out = tmp_path / "out.docx"
    result = run([str(leading_blank_md), str(template), str(out)])
    assert result.returncode == 0, result.stderr
    texts = _paragraph_texts(out)
    first_nonempty = next(i for i, t in enumerate(texts) if t.strip())
    assert first_nonempty >= 2


# ---------------------------------------------------------------------------
# Line-per-paragraph behaviour
# ---------------------------------------------------------------------------


def test_consecutive_lines_produce_separate_paragraphs(consecutive_lines_md, template, tmp_path):
    out = tmp_path / "out.docx"
    run([str(consecutive_lines_md), str(template), str(out)])
    texts = _paragraph_texts(out)
    non_empty = [t for t in texts if t.strip()]
    assert any("normal text" in t for t in non_empty)
    assert any("bold text" in t for t in non_empty)
    assert any("italic text" in t for t in non_empty)


def test_consecutive_lines_no_blank_paragraph_between(consecutive_lines_md, template, tmp_path):
    out = tmp_path / "out.docx"
    run([str(consecutive_lines_md), str(template), str(out)])
    texts = _paragraph_texts(out)
    non_empty_indices = [i for i, t in enumerate(texts) if t.strip()]
    assert non_empty_indices == list(range(non_empty_indices[0], non_empty_indices[0] + 3))


# ---------------------------------------------------------------------------
# Inline formatting
# ---------------------------------------------------------------------------


def test_bold_produces_bold_run(inline_formatting_md, template, tmp_path):
    out = tmp_path / "out.docx"
    run([str(inline_formatting_md), str(template), str(out)])
    assert _has_bold_run(out)


def test_italic_produces_italic_run(inline_formatting_md, template, tmp_path):
    out = tmp_path / "out.docx"
    run([str(inline_formatting_md), str(template), str(out)])
    assert _has_italic_run(out)


# ---------------------------------------------------------------------------
# Page breaks and right tabs
# ---------------------------------------------------------------------------


def test_page_break_produces_br_in_output(page_break_md, template, tmp_path):
    out = tmp_path / "out.docx"
    run([str(page_break_md), str(template), str(out)])
    with zipfile.ZipFile(str(out)) as zf:
        doc_xml = zf.read("word/document.xml").decode()
    assert 'w:type="page"' in doc_xml


def test_right_tab_applied_in_output(right_tab_md, template, tmp_path):
    out = tmp_path / "out.docx"
    run([str(right_tab_md), str(template), str(out)])
    with zipfile.ZipFile(str(out)) as zf:
        root = ET.parse(zf.open('word/document.xml')).getroot()
    tabs = root.findall(f'.//{{{W}}}tab')
    assert tabs


def test_right_tab_stop_in_ppr(right_tab_md, template, tmp_path):
    out = tmp_path / "out.docx"
    run([str(right_tab_md), str(template), str(out)])
    with zipfile.ZipFile(str(out)) as zf:
        root = ET.parse(zf.open('word/document.xml')).getroot()
    right_tabs = [
        t for t in root.findall(f'.//{{{W}}}pPr/{{{W}}}tabs/{{{W}}}tab')
        if t.get(f'{{{W}}}val') == 'right'
    ]
    assert right_tabs


def test_title_subtitle_paragraph_styles_in_output(title_subtitle_md, template, tmp_path):
    out = tmp_path / "out.docx"
    run([str(title_subtitle_md), str(template), str(out)])
    styles = _paragraph_styles(out)
    assert any("Title" == s for s in styles)
    assert any("Subtitle" == s for s in styles)


# ---------------------------------------------------------------------------
# Font integrity
# ---------------------------------------------------------------------------


def test_font_relationship_targets_all_resolve(simple_md, template, tmp_path):
    out = tmp_path / "out.docx"
    result = run([str(simple_md), str(template), str(out)])
    assert result.returncode == 0, result.stderr

    import re
    with zipfile.ZipFile(out) as zf:
        zip_entries = set(zf.namelist())
        font_rels = zf.read("word/_rels/fontTable.xml.rels") if "word/_rels/fontTable.xml.rels" in zip_entries else b""

    targets = re.findall(rb'Target="([^"]+)"', font_rels)
    font_targets = [t.decode() for t in targets if b"://" not in t]
    missing = [t for t in font_targets if f"word/{t}" not in zip_entries]
    assert not missing


def test_font_content_types_present_when_fonts_embedded(simple_md, template, tmp_path):
    out = tmp_path / "out.docx"
    result = run([str(simple_md), str(template), str(out)])
    assert result.returncode == 0, result.stderr

    import re
    with zipfile.ZipFile(out) as zf:
        zip_entries = set(zf.namelist())
        font_rels = zf.read("word/_rels/fontTable.xml.rels") if "word/_rels/fontTable.xml.rels" in zip_entries else b""
        ct = zf.read("[Content_Types].xml")

    targets = re.findall(rb'Target="([^"]+)"', font_rels)
    font_targets = [t.decode() for t in targets if b"://" not in t]
    if not font_targets:
        pytest.skip("template has no embedded fonts")

    extensions = {t.rsplit(".", 1)[-1].lower() for t in font_targets}
    for ext in extensions:
        assert f'Extension="{ext}"'.encode() in ct
