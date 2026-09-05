import pytest
from pathlib import Path

FIXTURES = Path(__file__).parent / "fixtures"
SAMPLE = Path(__file__).parent.parent / "ref"


@pytest.fixture
def template(tmp_path):
    """Copy the sample template into a temp dir to avoid mutation."""
    src = SAMPLE / "TEMPLATE.docx"
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
def sample_md():
    return SAMPLE / "content.md"


@pytest.fixture
def leading_blank_md():
    return FIXTURES / "leading_blank.md"
