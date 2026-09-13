"""Tests for __init__.py — public package API."""

from md2docx import ConversionError, convert


def test_convert_is_importable():
    assert callable(convert)


def test_conversion_error_is_importable():
    assert issubclass(ConversionError, Exception)


def test_public_api_complete():
    import md2docx
    assert hasattr(md2docx, "convert")
    assert hasattr(md2docx, "ConversionError")
