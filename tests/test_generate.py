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

from generate import inject_blank_paragraphs

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
# Sample letter round-trip
# ---------------------------------------------------------------------------


def test_sample_letter(sample_md, template, tmp_path):
    out = tmp_path / "letter.docx"
    result = run([str(sample_md), str(template), str(out)])
    assert result.returncode == 0, result.stderr
    assert out.exists()
    texts = " ".join(_paragraph_texts(out))
    assert "Kaluza" in texts
    assert "Chris Tracey" in texts
    assert _has_real_bullets(out)
