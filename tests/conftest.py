import subprocess
import sys
import pytest
from pathlib import Path

FIXTURES = Path(__file__).parent / "fixtures"
SAMPLE = Path(__file__).parent.parent / "ref"
SCRIPT = Path(__file__).parent.parent / "generate.py"


def run_generator(args: list[str]) -> subprocess.CompletedProcess:
    return subprocess.run(
        [sys.executable, str(SCRIPT)] + args,
        capture_output=True,
        text=True,
    )


@pytest.fixture
def template(tmp_path):
    """Copy the sample template into a temp dir to avoid mutation."""
    src = SAMPLE / "template-style.docx"
    dst = tmp_path / "TEMPLATE.docx"
    dst.write_bytes(src.read_bytes())
    return dst


@pytest.fixture
def simple_md():
    return FIXTURES / "simple.md"


@pytest.fixture
def headings_md():
    return FIXTURES / "headings.md"


@pytest.fixture
def leading_blank_md():
    return FIXTURES / "leading_blank.md"


@pytest.fixture
def inline_formatting_md():
    return FIXTURES / "inline_formatting.md"


@pytest.fixture
def consecutive_lines_md():
    return FIXTURES / "consecutive_lines.md"


@pytest.fixture
def title_subtitle_md():
    return FIXTURES / "title_subtitle.md"
