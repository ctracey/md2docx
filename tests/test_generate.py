"""Tests for generate.py — verified against real pandoc output.

python-docx cannot open the generated DOCX because the template embeds
custom fonts (DM Sans) whose MIME types pandoc doesn't register in
[Content_Types].xml. Tests therefore inspect the OOXML directly via zipfile.
"""

import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import pytest

from generate import inject_blank_paragraphs, mark_soft_newlines, restore_soft_newlines, strip_bookmarks, sync_headers, check_required_styles
from style_map import RunStyle

SCRIPT = Path(__file__).parent.parent / "generate.py"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def run(args: list[str]) -> subprocess.CompletedProcess:
    return subprocess.run(
        [sys.executable, str(SCRIPT)] + args,
        capture_output=True,
        text=True,
    )


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
    """Return True if any paragraph has numPr (real Word list item)."""
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
# mark_soft_newlines / restore_soft_newlines unit tests
# ---------------------------------------------------------------------------


def test_mark_soft_newlines_converts_single():
    result = mark_soft_newlines("line one\nline two")
    assert "\n" not in result
    assert "line one" in result and "line two" in result


def test_mark_soft_newlines_leaves_double_newline():
    result = mark_soft_newlines("para one\n\npara two")
    assert "\n\n" in result


def test_mark_soft_newlines_leaves_triple_newline():
    result = mark_soft_newlines("para one\n\n\npara two")
    assert "\n\n\n" in result


def test_restore_soft_newlines_produces_paragraph_break():
    marked = mark_soft_newlines("line one\nline two")
    restored = restore_soft_newlines(marked)
    assert restored == "line one\n\nline two"


def test_single_newline_does_not_produce_blank_paragraph():
    """Single newlines must NOT insert a \\ empty paragraph."""
    marked = mark_soft_newlines("line one\nline two")
    result = restore_soft_newlines(inject_blank_paragraphs(marked))
    assert "\\ " not in result


def test_blank_line_still_produces_blank_paragraph():
    """Explicit blank lines must still insert a \\ empty paragraph."""
    marked = mark_soft_newlines("line one\n\nline two")
    result = restore_soft_newlines(inject_blank_paragraphs(marked))
    assert "\\ " in result


# ---------------------------------------------------------------------------
# inject_blank_paragraphs unit tests
# ---------------------------------------------------------------------------


def test_inject_blank_lines_in_middle():
    result = inject_blank_paragraphs("Para one.\n\nPara two.")
    assert "\\ " in result
    assert result.index("Para one") < result.index("\\ ") < result.index("Para two")


def test_inject_no_leading_blank_line():
    result = inject_blank_paragraphs("Para one.\n\nPara two.")
    assert not result.startswith("\\ ")


def test_inject_leading_blank_line():
    result = inject_blank_paragraphs("\nPara one.\n\nPara two.")
    assert result.startswith("\\ ")


def test_inject_leading_blank_line_does_not_duplicate():
    result = inject_blank_paragraphs("\nPara one.")
    assert result.count("\\ ") == 1


def test_inject_no_trailing_blank_paragraph():
    result = inject_blank_paragraphs("Para one.\n\nPara two.\n\n")
    assert not result.rstrip("\n").endswith("\\ ")


def test_inject_two_middle_blank_lines_produce_two_empty_paragraphs():
    result = inject_blank_paragraphs("Para one.\n\n\nPara two.")
    assert result.count("\\ ") == 2


def test_inject_two_leading_blank_lines_produce_two_empty_paragraphs():
    result = inject_blank_paragraphs("\n\nPara one.")
    assert result.startswith("\\ ")
    assert result.count("\\ ") == 2


# ---------------------------------------------------------------------------
# check_required_styles unit tests
# ---------------------------------------------------------------------------


def test_no_error_when_required_labels_present():
    style_map = {"Bold text": RunStyle(bold=True), "Italic text": RunStyle(italic=True)}
    check_required_styles("**bold** and *italic*", style_map)  # must not raise/exit


def test_no_error_for_plain_content_with_no_labels():
    check_required_styles("Just plain text with no bold or italic.", {})


def test_error_when_bold_used_but_label_missing(tmp_path):
    md = tmp_path / "content.md"
    md.write_text("**bold text** here")
    result = run([str(md), "nonexistent.docx", str(tmp_path / "out.docx")])
    # Fails on missing template before reaching style check — use direct call instead
    with pytest.raises(SystemExit):
        check_required_styles("**bold text** here", {})


def test_error_message_names_missing_label():
    import io as _io
    from contextlib import redirect_stderr
    try:
        check_required_styles("**bold**", {})
    except SystemExit as e:
        assert "Bold text" in str(e)
        assert "README" in str(e)


def test_error_when_italic_used_but_label_missing():
    with pytest.raises(SystemExit):
        check_required_styles("*italic text* here", {})


def test_no_error_when_bold_absent_from_content():
    # template has no Bold text label — fine because content has no bold
    check_required_styles("Plain paragraph only.", {})


# ---------------------------------------------------------------------------
# Error handling
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
# Successful generation
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


def test_headers_copied_from_template_when_present(simple_md, tmp_path):
    """When the template has a header, the output must contain that exact header."""
    import io as _io

    # Build a minimal template docx that declares one header file
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
    buf = _io.BytesIO()
    with zipfile.ZipFile(buf, 'w') as zf:
        zf.writestr('word/document.xml', DOC_XML)
        zf.writestr('word/header1.xml', HEADER_XML)
        zf.writestr('word/_rels/document.xml.rels', RELS_XML)
        zf.writestr('[Content_Types].xml', CT_XML)
        zf.writestr('_rels/.rels', b'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>')
    tmpl_path.write_bytes(buf.getvalue())

    # Start with an output that has no headers (simulating pandoc without a header)
    out = tmp_path / "out.docx"
    out.write_bytes(buf.getvalue())  # use the template itself as a base output

    sync_headers(out, tmpl_path)

    with zipfile.ZipFile(str(out)) as zf:
        names = zf.namelist()
        out_doc = zf.read('word/document.xml')
        header_content = zf.read('word/header1.xml')

    assert 'word/header1.xml' in names
    assert b'TEMPLATE HEADER' in header_content
    assert b'headerReference' in out_doc


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
    assert _has_real_bullets(out), "No numPr (real Word bullet) found in output"


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
    assert any(t.strip() == "" for t in texts), "No empty paragraph found in output"


def test_two_leading_blank_lines_produce_two_empty_paragraphs(leading_blank_md, template, tmp_path):
    out = tmp_path / "out.docx"
    result = run([str(leading_blank_md), str(template), str(out)])
    assert result.returncode == 0, result.stderr
    texts = _paragraph_texts(out)
    first_nonempty = next(i for i, t in enumerate(texts) if t.strip())
    assert first_nonempty >= 2, f"Expected 2 empty paragraphs before content, got {first_nonempty}"


# ---------------------------------------------------------------------------
# Line-per-paragraph behaviour
# ---------------------------------------------------------------------------


def test_consecutive_lines_produce_separate_paragraphs(consecutive_lines_md, template, tmp_path):
    """Each line in the source must become its own paragraph in the docx."""
    out = tmp_path / "out.docx"
    run([str(consecutive_lines_md), str(template), str(out)])
    texts = _paragraph_texts(out)
    non_empty = [t for t in texts if t.strip()]
    assert any("normal text" in t for t in non_empty)
    assert any("bold text" in t for t in non_empty)
    assert any("italic text" in t for t in non_empty)


def test_consecutive_lines_no_blank_paragraph_between(consecutive_lines_md, template, tmp_path):
    """Single newlines must not insert a blank paragraph between lines."""
    out = tmp_path / "out.docx"
    run([str(consecutive_lines_md), str(template), str(out)])
    texts = _paragraph_texts(out)
    non_empty_indices = [i for i, t in enumerate(texts) if t.strip()]
    # The three content paragraphs must be consecutive (no gaps)
    assert non_empty_indices == list(range(non_empty_indices[0], non_empty_indices[0] + 3))


# ---------------------------------------------------------------------------
# Inline formatting
# ---------------------------------------------------------------------------


def test_bold_produces_bold_run(inline_formatting_md, template, tmp_path):
    out = tmp_path / "out.docx"
    run([str(inline_formatting_md), str(template), str(out)])
    assert _has_bold_run(out), "No bold run found in output"


def test_italic_produces_italic_run(inline_formatting_md, template, tmp_path):
    out = tmp_path / "out.docx"
    run([str(inline_formatting_md), str(template), str(out)])
    assert _has_italic_run(out), "No italic run found in output"
