"""Tests for content.py — markdown validation and transformation."""

import pytest

from md2docx.content import (
    Content,
    _PAGE_BREAK_MARKER,
    _PAGE_BREAK_XML,
    _RT_MARKER,
    inject_blank_paragraphs,
    inject_page_break_markers,
    inject_right_tab_markers,
    inject_title_styles,
    mark_soft_newlines,
    restore_soft_newlines,
    validate_content,
)
from md2docx.exceptions import ConversionError
from md2docx.style_map import RunStyle


# ---------------------------------------------------------------------------
# mark_soft_newlines / restore_soft_newlines
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
    marked = mark_soft_newlines("line one\nline two")
    result = restore_soft_newlines(inject_blank_paragraphs(marked))
    assert "\\ " not in result


def test_blank_line_still_produces_blank_paragraph():
    marked = mark_soft_newlines("line one\n\nline two")
    result = restore_soft_newlines(inject_blank_paragraphs(marked))
    assert "\\ " in result


# ---------------------------------------------------------------------------
# inject_blank_paragraphs
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
# inject_page_break_markers
# ---------------------------------------------------------------------------


def test_page_break_replaces_exactly_three_equals():
    result = inject_page_break_markers("===")
    assert result == _PAGE_BREAK_MARKER


def test_page_break_not_triggered_by_more_equals():
    for s in ("====", "==", "=", "====="):
        assert inject_page_break_markers(s) == s


def test_page_break_not_triggered_by_equals_with_extra():
    assert inject_page_break_markers("=== ") == "=== "
    assert inject_page_break_markers(" ===") == " ==="


def test_page_break_marker_survives_pipeline():
    text = "before\n===\nafter"
    marked = inject_page_break_markers(text)
    processed = restore_soft_newlines(inject_blank_paragraphs(mark_soft_newlines(marked)))
    assert _PAGE_BREAK_MARKER in processed


def test_page_break_xml_inserted_after_pipeline():
    text = "before\n===\nafter"
    marked = inject_page_break_markers(text)
    processed = restore_soft_newlines(inject_blank_paragraphs(mark_soft_newlines(marked)))
    final = processed.replace(_PAGE_BREAK_MARKER, _PAGE_BREAK_XML)
    assert _PAGE_BREAK_XML in final
    assert _PAGE_BREAK_MARKER not in final


# ---------------------------------------------------------------------------
# inject_right_tab_markers
# ---------------------------------------------------------------------------


def test_right_tab_replaces_double_chevron():
    result = inject_right_tab_markers("left >> right")
    assert ">>" not in result
    assert _RT_MARKER in result


def test_right_tab_splits_correctly():
    result = inject_right_tab_markers("left >> right")
    assert result == f"left {_RT_MARKER}right"


def test_right_tab_no_left_text():
    result = inject_right_tab_markers(">> right only")
    assert result == f"{_RT_MARKER}right only"


def test_right_tab_in_heading():
    result = inject_right_tab_markers("## Heading >> Meta")
    assert f"## Heading {_RT_MARKER}Meta" == result


def test_right_tab_strips_space_after_chevron():
    result = inject_right_tab_markers("left >>   right")
    assert result == f"left {_RT_MARKER}right"


def test_right_tab_only_first_occurrence():
    result = inject_right_tab_markers("a >> b >> c")
    assert result.count(_RT_MARKER) == 1


def test_right_tab_unchanged_when_no_chevron():
    result = inject_right_tab_markers("normal line")
    assert result == "normal line"


# ---------------------------------------------------------------------------
# inject_title_styles
# ---------------------------------------------------------------------------


_TITLE_STYLES = {"Title", "Subtitle"}
_TITLE_ONLY = {"Title"}


def test_inject_title_wraps_percent_line():
    result = inject_title_styles("%My Title\n", _TITLE_STYLES)
    assert '::: {custom-style="Title"}' in result
    assert "My Title" in result


def test_inject_subtitle_wraps_double_percent_line():
    result = inject_title_styles("%%My Subtitle\n", _TITLE_STYLES)
    assert '::: {custom-style="Subtitle"}' in result
    assert "My Subtitle" in result


def test_inject_title_and_subtitle_independently():
    result = inject_title_styles("%Title\n\n# Heading\n\n%%Subtitle\n", _TITLE_STYLES)
    assert '::: {custom-style="Title"}' in result
    assert '::: {custom-style="Subtitle"}' in result


def test_inject_double_percent_not_matched_as_title():
    result = inject_title_styles("%%Subtitle\n", _TITLE_STYLES)
    assert result.count(":::") == 2
    assert '::: {custom-style="Title"}' not in result


def test_inject_title_strips_leading_space_after_sigil():
    result = inject_title_styles("% My Title\n", _TITLE_STYLES)
    assert "My Title" in result
    assert " My Title" not in result


def test_inject_title_skips_when_no_styles_defined():
    result = inject_title_styles("%Title\n%%Subtitle\n", set())
    assert ":::" not in result


def test_inject_title_only_when_subtitle_style_absent():
    result = inject_title_styles("%Title\n%%Subtitle\n", _TITLE_ONLY)
    assert '::: {custom-style="Title"}' in result
    assert '::: {custom-style="Subtitle"}' not in result


def test_non_percent_lines_are_unchanged():
    result = inject_title_styles("# Heading\nBody\n", _TITLE_STYLES)
    assert ":::" not in result


# ---------------------------------------------------------------------------
# validate_content
# ---------------------------------------------------------------------------


def test_no_error_when_required_labels_present():
    run_styles = {"Bold text": RunStyle(bold=True), "Italic text": RunStyle(italic=True)}
    validate_content("**bold** and *italic*", run_styles)


def test_no_error_for_plain_content_with_no_labels():
    validate_content("Just plain text with no bold or italic.", {})


def test_error_when_bold_used_but_label_missing():
    with pytest.raises(ConversionError, match="Bold text"):
        validate_content("**bold text** here", {})


def test_error_message_names_missing_label():
    with pytest.raises(ConversionError) as exc:
        validate_content("**bold**", {})
    assert "Bold text" in str(exc.value)
    assert "README" in str(exc.value)


def test_error_when_italic_used_but_label_missing():
    with pytest.raises(ConversionError):
        validate_content("*italic text* here", {})


def test_no_error_when_bold_absent_from_content():
    validate_content("Plain paragraph only.", {})
