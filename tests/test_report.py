"""Tests for the report and error recovery utilities."""

from markdownpick.errors import ConversionError
from markdownpick.report import ConversionReport, safe_convert


def test_report_initially_clean():
    r = ConversionReport()
    assert r.has_issues is False


def test_report_warn():
    r = ConversionReport()
    r.warn("test warning")
    assert r.has_issues is True
    assert len(r.warnings) == 1
    assert len(r.errors) == 0


def test_report_error():
    r = ConversionReport()
    r.error("test error")
    assert r.has_issues is True
    assert len(r.errors) == 1


def test_safe_convert_success():
    result = safe_convert(lambda: 42, fallback=0)
    assert result == 42


def test_safe_convert_conversion_error():
    def failing():
        raise ConversionError("bad thing", element_type="cell")

    result = safe_convert(failing, fallback=0, label="test")
    assert result == 0


def test_safe_convert_generic_error():
    def failing():
        raise ValueError("oops")

    result = safe_convert(failing, fallback="fallback", label="test")
    assert result == "fallback"


def test_safe_convert_populates_report():
    r = ConversionReport()
    safe_convert(lambda: (_ for _ in ()).throw(ValueError("boom")), fallback=0, report=r, label="x")
    assert r.has_issues is True
