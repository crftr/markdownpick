"""Tests for the Excel converter."""

from __future__ import annotations

from pathlib import Path

from markdownpick.errors import CorruptFileError
from markdownpick.excel.converter import convert_excel_to_markdown, validate_xlsx
from markdownpick.report import ConversionReport


def test_convert_formulas(xlsx_formulas: Path, tmp_path: Path):
    out = tmp_path / "formulas.md"
    convert_excel_to_markdown(xlsx_formulas, out)
    text = out.read_text()
    assert "---" in text
    assert "generator: markdownpick/xlsx" in text
    assert "`=A2*2`" in text
    assert "Formula Dependency Map" in text


def test_convert_multi_sheet(xlsx_multi_sheet: Path, tmp_path: Path):
    out = tmp_path / "multi.md"
    convert_excel_to_markdown(xlsx_multi_sheet, out, include_dependency_map=False, include_formatting_map=False)
    text = out.read_text()
    assert "Sheet: Sheet1" in text
    assert "Sheet: Sheet2" in text


def test_convert_empty(xlsx_empty: Path, tmp_path: Path):
    out = tmp_path / "empty.md"
    convert_excel_to_markdown(xlsx_empty, out, include_dependency_map=False, include_formatting_map=False)
    text = out.read_text()
    assert "---" in text


def test_merged_cells(xlsx_formulas: Path, tmp_path: Path):
    out = tmp_path / "merged.md"
    convert_excel_to_markdown(xlsx_formulas, out, include_dependency_map=False, include_formatting_map=False)
    text = out.read_text()
    assert "\u2197" in text  # span marker


def test_corrupt_xlsx(corrupt_file: Path):
    try:
        validate_xlsx(corrupt_file)
    except CorruptFileError:
        return
    raise AssertionError("Should have raised CorruptFileError")


def test_convert_with_report(xlsx_formulas: Path, tmp_path: Path):
    out = tmp_path / "report.md"
    report = ConversionReport()
    convert_excel_to_markdown(xlsx_formulas, out, report=report)
    assert out.exists()
