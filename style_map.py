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
    "Normal text",
    "Bold text",
    "Italic text",
}


@dataclass
class RunStyle:
    """Run-level formatting properties extracted from a style label paragraph."""
    font: str | None = None
    size_half_pt: int | None = None   # OOXML sz (half-points; divide by 2 for pt)
    color: str | None = None          # hex colour without #
    bold: bool = False
    italic: bool = False

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
        first_run = next((r for r in p.findall(f'{{{W}}}r')), None)
        rpr = first_run.find(f'{{{W}}}rPr') if first_run is not None else None
        found[text] = _parse_rpr(rpr) if rpr is not None else RunStyle()

    return found
