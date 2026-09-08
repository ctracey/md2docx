"""Tests for validate_template() — style template completeness checks."""

import io
import zipfile
from pathlib import Path

import pytest

from style_map import validate_template, KNOWN_LABELS
from tests.conftest import SAMPLE, run_generator

# ---------------------------------------------------------------------------
# Minimal DOCX builders
# ---------------------------------------------------------------------------

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_CT = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""

_RELS = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>"""


def _styles_xml(style_ids: list[str]) -> str:
    styles = "".join(
        f'<w:style w:type="paragraph" w:styleId="{sid}"><w:name w:val="{sid}"/></w:style>'
        for sid in style_ids
    )
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:styles xmlns:w="{_W}">{styles}</w:styles>'
    )


def _para(text: str, tab_stop: bool = False) -> str:
    tabs = (
        f'<w:tabs><w:tab w:val="right" w:pos="10800"/></w:tabs>'
        if tab_stop else ''
    )
    return (
        f'<w:p><w:pPr>{tabs}</w:pPr>'
        f'<w:r><w:t>{text}</w:t></w:r></w:p>'
    )


def _make_template(paragraphs: list[str], style_ids: list[str], tmp_path: Path) -> Path:
    body = "".join(paragraphs) + "<w:sectPr/>"
    doc_xml = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{_W}"><w:body>{body}</w:body></w:document>'
    )
    path = tmp_path / "template.docx"
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("word/document.xml", doc_xml)
        zf.writestr("word/styles.xml", _styles_xml(style_ids))
        zf.writestr("[Content_Types].xml", _CT)
        zf.writestr("_rels/.rels", _RELS)
    path.write_bytes(buf.getvalue())
    return path


def _complete_template(tmp_path: Path) -> Path:
    """Minimal template with all required elements present."""
    label_paras = [_para(label) for label in KNOWN_LABELS]
    right_tab_para = _para("NormalRightAlignedTabStop", tab_stop=True)
    return _make_template(
        label_paras + [right_tab_para],
        ["Title", "Subtitle", "Heading1"],
        tmp_path,
    )


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------


def test_real_template_is_complete():
    """The committed template-style.docx must pass validation with no issues."""
    issues = validate_template(SAMPLE / "template-style.docx")
    assert issues == [], f"Real template has issues:\n" + "\n".join(issues)


def test_complete_minimal_template_passes(tmp_path):
    tmpl = _complete_template(tmp_path)
    assert validate_template(tmpl) == []


def test_missing_body_label_reported(tmp_path):
    # Template with everything except 'Bold text'
    labels = [l for l in KNOWN_LABELS if l != "Bold text"]
    tmpl = _make_template(
        [_para(l) for l in labels] + [_para("NormalRightAlignedTabStop", tab_stop=True)],
        ["Title", "Subtitle"],
        tmp_path,
    )
    issues = validate_template(tmpl)
    assert any('"Bold text"' in i for i in issues)


def test_all_missing_labels_reported(tmp_path):
    # Template with no body labels at all
    tmpl = _make_template(
        [_para("NormalRightAlignedTabStop", tab_stop=True)],
        ["Title", "Subtitle"],
        tmp_path,
    )
    issues = validate_template(tmpl)
    for label in KNOWN_LABELS:
        assert any(label in i for i in issues), f'"{label}" not reported'


def test_missing_right_aligned_tab_stop_reported(tmp_path):
    tmpl = _make_template(
        [_para(l) for l in KNOWN_LABELS],
        ["Title", "Subtitle"],
        tmp_path,
    )
    issues = validate_template(tmpl)
    assert any("RightAlignedTabStop" in i for i in issues)


def test_invalid_right_aligned_tab_stop_reported(tmp_path):
    # RightAlignedTabStop present but no right tab stop in pPr
    tmpl = _make_template(
        [_para(l) for l in KNOWN_LABELS] + [_para("NormalRightAlignedTabStop")],
        ["Title", "Subtitle"],
        tmp_path,
    )
    issues = validate_template(tmpl)
    assert any("RightAlignedTabStop" in i for i in issues)


def test_missing_title_style_reported(tmp_path):
    tmpl = _make_template(
        [_para(l) for l in KNOWN_LABELS] + [_para("NormalRightAlignedTabStop", tab_stop=True)],
        ["Subtitle"],  # No Title
        tmp_path,
    )
    issues = validate_template(tmpl)
    assert any('"Title"' in i for i in issues)


def test_missing_subtitle_style_reported(tmp_path):
    tmpl = _make_template(
        [_para(l) for l in KNOWN_LABELS] + [_para("NormalRightAlignedTabStop", tab_stop=True)],
        ["Title"],  # No Subtitle
        tmp_path,
    )
    issues = validate_template(tmpl)
    assert any('"Subtitle"' in i for i in issues)


def test_all_issues_reported_at_once(tmp_path):
    # Completely empty template — all issues should appear together
    tmpl = _make_template([], [], tmp_path)
    issues = validate_template(tmpl)
    assert len(issues) >= len(KNOWN_LABELS) + 3  # labels + RightAligned + Title + Subtitle


def test_cli_exits_with_all_issues_on_incomplete_template(simple_md, tmp_path):
    tmpl = _make_template([], [], tmp_path)
    result = run_generator([str(simple_md), str(tmpl), str(tmp_path / "out.docx")])
    assert result.returncode == 1
    assert "style template is incomplete" in result.stderr
    assert "README" in result.stderr
    # Multiple issues should be listed
    assert result.stderr.count("•") >= len(KNOWN_LABELS) + 3


def test_cli_proceeds_when_template_is_complete(simple_md, tmp_path):
    tmpl = _complete_template(tmp_path)
    # This minimal template won't work with pandoc, but validation should pass
    # (pandoc will fail — that's fine, we just confirm we get past validation)
    result = run_generator([str(simple_md), str(tmpl), str(tmp_path / "out.docx")])
    assert "style template is incomplete" not in result.stderr
