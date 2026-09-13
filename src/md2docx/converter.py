"""Pipeline orchestrator and CLI entry point."""

import argparse
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from md2docx import document
from md2docx.content import Content
from md2docx.exceptions import ConversionError
from md2docx.style_map import StyleMap, validate_template


def convert(
    content: Path,
    template: Path,
    output: Path,
    partials: dict[str, Path] | None = None,
) -> None:
    """Convert a Markdown file to a styled DOCX.

    Raises ConversionError on any failure.
    """
    partials = partials or {}

    if not content.exists():
        raise ConversionError(f"content file not found: {content}")
    if not template.exists():
        raise ConversionError(f"template file not found: {template}")
    if shutil.which("pandoc") is None:
        raise ConversionError("pandoc not found. Install with: brew install pandoc")

    issues = validate_template(template)
    if issues:
        details = "\n".join(f"  • {issue}" for issue in issues)
        raise ConversionError(
            f"style template is incomplete:\n{details}\n"
            "See README — Style mapping for template requirements."
        )

    style_map = StyleMap.from_template(template)
    prepared = Content.from_markdown(content, style_map)

    _render(prepared, template, output)
    document.apply(output, template, style_map, partials)


def _render(content: Content, template: Path, output: Path) -> None:
    with tempfile.NamedTemporaryFile(mode='w', suffix='.md', delete=False) as tmp:
        tmp.write(content.text)
        tmp_path = Path(tmp.name)

    try:
        result = subprocess.run(
            [
                "pandoc",
                str(tmp_path),
                "--reference-doc", str(template),
                "-o", str(output),
            ],
            capture_output=True,
            text=True,
        )
    finally:
        tmp_path.unlink()

    if result.returncode != 0:
        raise ConversionError(f"pandoc failed:\n{result.stderr}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Generate a styled DOCX from Markdown and a DOCX template.",
        epilog="Requires pandoc to be installed (brew install pandoc).",
    )
    parser.add_argument("content", type=Path, help="Markdown input file (.md)")
    parser.add_argument("template", type=Path, help="DOCX style template (.docx)")
    parser.add_argument("output", type=Path, help="Output DOCX file (.docx)")
    parser.add_argument(
        "--partial", metavar="NAME=FILE.docx", action="append", default=[],
        help="Splice partial DOCX at {{NAME}} placeholder (repeatable)",
    )
    args = parser.parse_args()

    partials: dict[str, Path] = {}
    for spec in args.partial:
        if '=' not in spec:
            sys.exit(f"Error: --partial must be NAME=FILE.docx, got: {spec!r}")
        pname, ppath_str = spec.split('=', 1)
        ppath = Path(ppath_str)
        if not ppath.exists():
            sys.exit(f"Error: partial file not found: {ppath}")
        partials[pname] = ppath

    try:
        convert(args.content, args.template, args.output, partials)
    except ConversionError as e:
        sys.exit(f"Error: {e}")

    print(f"Generated: {args.output}")


if __name__ == "__main__":
    main()
