"""Discover run-level styles from labelled example paragraphs in a DOCX template.

Each known label must appear exactly once in the template body. The run
properties of its first run define how that markdown construct is rendered.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import xml.etree.ElementTree as ET
import zipfile

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

KNOWN_LABELS = {
    "Bold text",
    "Italic text",
    "Code block text",
    "Bullet point text",
}


@dataclass
class RunStyle:
    """Formatting properties extracted from a style label paragraph."""
    font: str | None = None
    size_half_pt: int | None = None   # OOXML sz (half-points; divide by 2 for pt)
    color: str | None = None          # hex colour without #
    bold: bool = False
    italic: bool = False
    para_style: str | None = None     # paragraph style ID (e.g. "Title", "Subtitle")
    ind_left: int | None = None       # w:ind w:left (twips)
    ind_hanging: int | None = None    # w:ind w:hanging (twips)

    @property
    def size_pt(self) -> float | None:
        return self.size_half_pt / 2 if self.size_half_pt is not None else None


def _parse_rpr(rpr: ET.Element) -> RunStyle:
    s = RunStyle()

    fonts_el = rpr.find(f'{{{W}}}rFonts')
    if fonts_el is not None:
        s.font = (
            fonts_el.get(f'{{{W}}}ascii')
            or fonts_el.get(f'{{{W}}}hAnsi')
            or fonts_el.get(f'{{{W}}}cs')
        )

    sz_el = rpr.find(f'{{{W}}}sz')
    if sz_el is not None:
        val = sz_el.get(f'{{{W}}}val')
        if val:
            s.size_half_pt = int(val)

    color_el = rpr.find(f'{{{W}}}color')
    if color_el is not None:
        c = color_el.get(f'{{{W}}}val')
        if c and c != 'auto':
            s.color = c

    s.bold = rpr.find(f'{{{W}}}b') is not None
    s.italic = rpr.find(f'{{{W}}}i') is not None

    return s


def read_style_map(template_path: Path) -> dict[str, RunStyle]:
    """Scan the template body for known style labels and return their run properties.

    Returns a partial dict — labels absent from the template are not included.
    Raises ValueError if any label appears more than once.
    """
    with zipfile.ZipFile(str(template_path)) as zf:
        with zf.open('word/document.xml') as f:
            root = ET.parse(f).getroot()

    found: dict[str, RunStyle] = {}

    for p in root.findall(f'.//{{{W}}}p'):
        text = ''.join(t.text or '' for t in p.findall(f'.//{{{W}}}t'))
        if text not in KNOWN_LABELS:
            continue
        if text in found:
            raise ValueError(
                f"Style label '{text}' appears more than once in the template. "
                "Each label must appear exactly once."
            )
        ppr = p.find(f'{{{W}}}pPr')
        style_el = ppr.find(f'{{{W}}}pStyle') if ppr is not None else None
        para_style = style_el.get(f'{{{W}}}val') if style_el is not None else None

        first_run = next((r for r in p.findall(f'{{{W}}}r')), None)
        rpr = first_run.find(f'{{{W}}}rPr') if first_run is not None else None
        run_style = _parse_rpr(rpr) if rpr is not None else RunStyle()
        run_style.para_style = para_style

        if ppr is not None:
            ind_el = ppr.find(f'{{{W}}}ind')
            if ind_el is not None:
                left = ind_el.get(f'{{{W}}}left')
                hanging = ind_el.get(f'{{{W}}}hanging')
                if left: run_style.ind_left = int(left)
                if hanging: run_style.ind_hanging = int(hanging)

        found[text] = run_style

    return found


def read_right_tab_stop(template_path: Path) -> dict | None:
    """Find 'RightAlignedTabStop' in the template body and return its tab stop definition.

    Returns a dict with 'val', 'pos', and 'leader' keys, or None if the label
    is absent (feature disabled). Raises ValueError if the label is present but
    the paragraph has no right-aligned tab stop in its pPr.
    """
    with zipfile.ZipFile(str(template_path)) as zf:
        raw = zf.read('word/document.xml').decode()

    import re as _re
    paras = _re.findall(r'<w:p[ >].*?</w:p>', raw, _re.DOTALL)
    for p in paras:
        texts = _re.findall(r'<w:t[^>]*>([^<]*)</w:t>', p)
        if 'RightAlignedTabStop' not in ''.join(texts):
            continue

        # Validate: must have a right-aligned tab stop in pPr
        tab_def = None
        for m in _re.finditer(r'<w:tab\b([^>]*)/?>', p):
            attrs = m.group(1)
            val = _re.search(r'w:val="([^"]+)"', attrs)
            pos = _re.search(r'w:pos="([^"]+)"', attrs)
            leader = _re.search(r'w:leader="([^"]+)"', attrs)
            if val and val.group(1) == 'right' and pos:
                tab_def = {
                    'val': 'right',
                    'pos': pos.group(1),
                    'leader': leader.group(1) if leader else 'none',
                }
                break
        if tab_def is None:
            raise ValueError(
                "The 'RightAlignedTabStop' paragraph was found in the template but "
                "does not define a right-aligned tab stop (w:val=\"right\") in its pPr. "
                "See README — Style mapping."
            )

        # Extract run properties from the run containing "RightAlignedTabStop" text
        right_style = None
        runs = _re.findall(r'<w:r[ >].*?</w:r>', p, _re.DOTALL)
        for run in runs:
            run_texts = _re.findall(r'<w:t[^>]*>([^<]*)</w:t>', run)
            if 'RightAlignedTabStop' in ''.join(run_texts):
                rpr_match = _re.search(r'<w:rPr>(.*?)</w:rPr>', run, _re.DOTALL)
                if rpr_match:
                    rpr_el = ET.fromstring(
                        f'<w:rPr xmlns:w="{W}">{rpr_match.group(1)}</w:rPr>'
                    )
                    right_style = _parse_rpr(rpr_el)
                break

        tab_def['right_style'] = right_style
        return tab_def
    return None


def validate_template(template_path: Path) -> list[str]:
    """Check the style template for completeness.

    Returns a list of human-readable issue strings.
    An empty list means the template is complete.
    """
    issues: list[str] = []

    # Body label paragraphs
    try:
        style_map = read_style_map(template_path)
    except ValueError as e:
        issues.append(str(e))
        style_map = {}

    for label in sorted(KNOWN_LABELS):
        if label not in style_map:
            issues.append(f'Body label paragraph not found: "{label}"')

    # RightAlignedTabStop
    try:
        if read_right_tab_stop(template_path) is None:
            issues.append('Body label paragraph not found: "RightAlignedTabStop"')
    except ValueError as e:
        issues.append(str(e))

    # Title and Subtitle paragraph styles
    para_styles = read_paragraph_style_ids(template_path)
    for style_id in ('Title', 'Subtitle'):
        if style_id not in para_styles:
            issues.append(f'Paragraph style not defined in styles.xml: "{style_id}"')

    return issues


def read_paragraph_style_ids(template_path: Path) -> set[str]:
    """Return the set of paragraph style IDs defined in the template's styles.xml."""
    with zipfile.ZipFile(str(template_path)) as zf:
        with zf.open('word/styles.xml') as f:
            root = ET.parse(f).getroot()
    return {
        style.get(f'{{{W}}}styleId')
        for style in root.findall(f'{{{W}}}style')
        if style.get(f'{{{W}}}type') == 'paragraph'
        and style.get(f'{{{W}}}styleId')
    }


@dataclass
class StyleMap:
    """All style information extracted from a DOCX template."""
    run_styles: dict[str, RunStyle]
    para_styles: set[str]
    right_tab: dict | None

    @classmethod
    def from_template(cls, template: Path) -> StyleMap:
        return cls(
            run_styles=read_style_map(template),
            para_styles=read_paragraph_style_ids(template),
            right_tab=read_right_tab_stop(template),
        )
