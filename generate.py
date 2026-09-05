#!/usr/bin/env python3
"""Generate a styled DOCX from a Markdown file and a DOCX style template."""

import argparse
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


def inject_blank_paragraphs(text: str) -> str:
    r"""Insert one '\ ' paragraph per blank line so Word renders them as empty lines.

    Prepend \n so a single leading blank line matches the same \n{2,} pattern
    as blank lines in the middle of the document. N consecutive newlines means
    N-1 blank lines, so N-1 '\ ' paragraphs are inserted.
    """
    padded = '\n' + text.rstrip()

    def replace(m: re.Match) -> str:
        blanks = len(m.group(0)) - 1
        return '\n\n' + '\\ \n\n' * blanks

    result = re.sub(r'\n{2,}', replace, padded)
    return result.lstrip('\n') + '\n'


def main():
    parser = argparse.ArgumentParser(
        description="Generate a styled DOCX from Markdown and a DOCX template.",
        epilog="Requires pandoc to be installed (brew install pandoc).",
    )
    parser.add_argument("content", type=Path, help="Markdown input file (.md)")
    parser.add_argument("template", type=Path, help="DOCX style template (.docx)")
    parser.add_argument("output", type=Path, help="Output DOCX file (.docx)")
    args = parser.parse_args()

    if not args.content.exists():
        sys.exit(f"Error: content file not found: {args.content}")
    if not args.template.exists():
        sys.exit(f"Error: template file not found: {args.template}")
    if shutil.which("pandoc") is None:
        sys.exit("Error: pandoc not found. Install with: brew install pandoc")

    modified = inject_blank_paragraphs(args.content.read_text())

    with tempfile.NamedTemporaryFile(mode='w', suffix='.md', delete=False) as tmp:
        tmp.write(modified)
        tmp_path = Path(tmp.name)

    try:
        result = subprocess.run(
            [
                "pandoc",
                str(tmp_path),
                "--reference-doc", str(args.template),
                "-o", str(args.output),
            ],
            capture_output=True,
            text=True,
        )
    finally:
        tmp_path.unlink()

    if result.returncode != 0:
        sys.exit(f"Error: pandoc failed:\n{result.stderr}")

    print(f"Generated: {args.output}")


if __name__ == "__main__":
    main()
