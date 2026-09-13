import pytest

from md2docx.exceptions import ConversionError


def test_conversion_error_is_exception_subclass():
    assert issubclass(ConversionError, Exception)


def test_conversion_error_is_catchable_specifically():
    with pytest.raises(ConversionError):
        raise ConversionError("something went wrong")


def test_conversion_error_does_not_catch_plain_exception():
    with pytest.raises(Exception):
        raise ConversionError("narrower than Exception is fine")
