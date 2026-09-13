"""Markdown content domain object — validates and transforms raw markdown for rendering."""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING

from md2docx.exceptions import ConversionError
from md2docx.style_map import RunStyle

if TYPE_CHECKING:
    from md2docx.style_map import StyleMap

_SOFT_NL = '\x00'
_PAGE_BREAK_MARKER = ''   # U+E002 PUA — survives preprocessing pipeline
_PAGE_BREAK_XML = (
    '```{=openxml}\n'
    '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:r><w:br w:type="page"/></w:r></w:p>\n'
    '```'
)
_RT_MARKER = ''   # U+E000 PUA — right-tab sentinel

_EXTENSION_CONSTRUCTS = [
    (r'\*\*\S', "Bold text", "**bold**"),
    (r'(?<!\*)\*(?!\*)\S', "Italic text", "*italic*"),
    (r'`', "Code block text", "`code`"),
]


@dataclass
class Content:
    """Validated and transformed markdown ready to pass to the renderer."""
    text: str

    @classmethod
    def from_markdown(cls, path: Path, style_map: StyleMap) -> Content:
        raw = path.read_text()
        validate_content(raw, style_map.run_styles)
        return cls(text=_transform(raw, style_map))


def validate_content(text: str, run_styles: dict[str, RunStyle]) -> None:
    """Raise ConversionError if text uses constructs whose style labels are absent."""
    missing = [
        (display, label)
        for pattern, label, display in _EXTENSION_CONSTRUCTS
        if re.search(pattern, text) and label not in run_styles
    ]
    if not missing:
        return
    details = "\n".join(
        f'  {display}  →  add a paragraph labelled "{label}" to the style template'
        for display, label in missing
    )
    raise ConversionError(
        f"style template is missing required labels:\n{details}\n"
        "See README — Style mapping — Extension mappings."
    )


def _transform(raw: str, style_map: StyleMap) -> str:
    raw = inject_title_styles(raw, style_map.para_styles)
    raw = inject_page_break_markers(raw)
    if style_map.right_tab:
        raw = inject_right_tab_markers(raw)
    text = restore_soft_newlines(inject_blank_paragraphs(mark_soft_newlines(raw)))
    return text.replace(_PAGE_BREAK_MARKER, _PAGE_BREAK_XML)


def mark_soft_newlines(text: str) -> str:
    """Replace single newlines with a sentinel so inject_blank_paragraphs ignores them.

    Each line becomes its own paragraph. Explicit blank lines still produce
    a visible empty paragraph via inject_blank_paragraphs.
    """
    return re.sub(r'(?<!\n)\n(?!\n)', _SOFT_NL, text)


def restore_soft_newlines(text: str) -> str:
    """Convert sentinels back into pandoc paragraph breaks (double newline, no blank para)."""
    return text.replace(_SOFT_NL, '\n\n')


def inject_blank_paragraphs(text: str) -> str:
    r"""Insert one '\ ' paragraph per blank line so Word renders them as empty lines."""
    padded = '\n' + text.rstrip()

    def replace(m: re.Match) -> str:
        blanks = len(m.group(0)) - 1
        return '\n\n' + '\\ \n\n' * blanks

    result = re.sub(r'\n{2,}', replace, padded)
    return result.lstrip('\n') + '\n'


def inject_page_break_markers(text: str) -> str:
    """Replace lines that are exactly '===' with a page-break sentinel."""
    lines = text.split('\n')
    return '\n'.join(_PAGE_BREAK_MARKER if line == '===' else line for line in lines)


def inject_title_styles(text: str, para_styles: set[str]) -> str:
    """Replace % and %% prefix lines with pandoc custom-style fenced divs.

    %<text>   → Title paragraph
    %%<text>  → Subtitle paragraph
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


def inject_right_tab_markers(text: str) -> str:
    """Replace >> with a sentinel so pandoc does not interpret it as a blockquote.

    Convention: left text >> right text
    """
    lines = text.split('\n')
    out = []
    for line in lines:
        if '>>' in line:
            line = re.sub(r'>>\s*', _RT_MARKER, line, count=1)
        out.append(line)
    return '\n'.join(out)
