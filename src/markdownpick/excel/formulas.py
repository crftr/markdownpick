"""Formula parsing and dependency mapping for Excel workbooks."""

from __future__ import annotations

import re

from markdownpick.tables import markdown_table


def _ordered_unique(items) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for item in items:
        if item not in seen:
            seen.add(item)
            out.append(item)
    return out


def _md_escape(value: str) -> str:
    """Escape markdown table-breaking characters."""
    return value.replace("|", "\\|").replace("\n", "<br>")


def extract_formula_references(formula: str) -> list[str]:
    """Extract direct cell/range references from an Excel formula.

    Handles:
    - A1, $A$1, Sheet1!A1
    - Ranges: A1:B10, Sheet1!A1:Sheet1!B10
    - Named ranges with single quotes: 'Sheet Name'!A1
    """
    text = formula[1:] if formula.startswith("=") else formula
    sheet_prefix = r"(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!"
    cell_ref = r"\$?[A-Z]{1,3}\$?\d+"
    range_ref = f"{cell_ref}(?::{cell_ref})?"
    sheet_ref = f"{sheet_prefix}{range_ref}(?::{sheet_prefix}{range_ref})?"
    pattern = re.compile(f"{sheet_ref}|{range_ref}")
    return _ordered_unique(match.group(0) for match in pattern.finditer(text))


def build_formula_dependency_map(sheet, max_row: int, max_col: int) -> str:
    """Build a markdown dependency map of formula cells and their references."""
    dep_rows: list[list[str]] = []

    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            cell = sheet.cell(row=r, column=c)
            if cell.data_type != "f" or cell.value is None:
                continue

            formula_text = str(cell.value)
            if not formula_text.startswith("="):
                formula_text = f"={formula_text}"

            refs = extract_formula_references(formula_text)
            dep_rows.append(
                [
                    cell.coordinate,
                    _md_escape(f"`{formula_text}`"),
                    _md_escape(", ".join(refs) if refs else "(no direct cell refs found)"),
                ]
            )

    if not dep_rows:
        return "_No formulas found on this sheet._"

    return markdown_table(["Formula Cell", "Formula", "References"], dep_rows)
