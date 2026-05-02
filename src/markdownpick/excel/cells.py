"""Cell rendering: value formatting, merged cells, comments, hyperlinks."""

from __future__ import annotations

from datetime import date, datetime, time

from openpyxl.cell.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet


def _md_escape(value: str) -> str:
    """Escape markdown table-breaking characters."""
    return value.replace("|", "\\|").replace("\n", "<br>")


def excel_col_label(col_num: int) -> str:
    """Convert 1-based column number to Excel letters (1 -> A, 27 -> AA)."""
    label = ""
    n = col_num
    while n > 0:
        n, r = divmod(n - 1, 26)
        label = chr(65 + r) + label
    return label


def format_cell_value(cell: Cell, value_cell: Cell | None, show_formula_values: bool) -> str:
    """Render a cell's value for markdown output.

    Handles formulas, dates, numbers, and rich text.
    """
    raw_value = cell.value

    if raw_value is None:
        return ""

    if cell.data_type == "f":
        formula_text = str(raw_value)
        if not formula_text.startswith("="):
            formula_text = f"={formula_text}"

        if show_formula_values and value_cell is not None and value_cell.value is not None:
            rendered = f"`{formula_text}` \u21d2 {value_cell.value}"
        else:
            rendered = f"`{formula_text}`"
        return _md_escape(rendered)

    # Date/time formatting
    if isinstance(raw_value, datetime):
        fmt = cell.number_format if cell.number_format else "%Y-%m-%d %H:%M:%S"
        if isinstance(fmt, str) and "%" in fmt:
            try:
                return _md_escape(raw_value.strftime(fmt))
            except Exception:
                pass
        return _md_escape(raw_value.strftime("%Y-%m-%d"))

    if isinstance(raw_value, date):
        return _md_escape(raw_value.strftime("%Y-%m-%d"))

    if isinstance(raw_value, time):
        return _md_escape(raw_value.strftime("%H:%M:%S"))

    # Number formatting (currency, percentage, etc.)
    if isinstance(raw_value, (int, float)) and cell.number_format:
        fmt = cell.number_format
        if isinstance(fmt, str) and fmt != "General":
            if "%" in fmt:
                try:
                    pct = raw_value * 100
                    return _md_escape(f"{pct:.1f}%")
                except Exception:
                    pass
            if "$" in fmt or "#" in fmt:
                try:
                    decimals = fmt.count("0") - fmt.split(".")[-1].count("0") if "." in fmt else 0
                    decimals = max(0, min(2, decimals if decimals > 0 else 0))
                    if "$" in fmt:
                        return _md_escape(f"${raw_value:,.{decimals}f}")
                    return _md_escape(f"{raw_value:,.{decimals}f}")
                except Exception:
                    pass

    return _md_escape(str(raw_value))


def detect_merged_cells(sheet: Worksheet) -> dict[str, bool]:
    """Detect merged cell ranges.

    Returns dict mapping cell coordinate to:
    - "merged" if the cell is part of a merged range but NOT the top-left
    - "span_start" if the cell is the top-left of a merged range
    """
    result: dict[str, bool] = {}
    for merged_range in sheet.merged_cells.ranges:
        min_row, min_col, max_row, max_col = (
            merged_range.min_row,
            merged_range.min_col,
            merged_range.max_row,
            merged_range.max_col,
        )
        # Top-left cell is the span_start
        top_left = f"{excel_col_label(min_col)}{min_row}"
        result[top_left] = True

        # All other cells in the range are merged
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                coord = f"{excel_col_label(c)}{r}"
                if coord != top_left:
                    result[coord] = False

    return result


def get_cell_comment(cell: Cell) -> str | None:
    """Extract a cell comment/note, or None."""
    try:
        if cell.comment and cell.comment.text:
            return cell.comment.text.strip()
    except Exception:
        pass
    return None


def stringify_cell(
    raw_cell: Cell,
    value_cell: Cell | None,
    show_formula_values: bool,
    merged_map: dict[str, bool] | None = None,
    show_comments: bool = False,
) -> str:
    """Render one cell for markdown output.

    Handles merged cells, formulas, comments.
    """
    coord = raw_cell.coordinate

    if merged_map and coord in merged_map:
        if merged_map[coord] is False:
            return ""
        prefix = "\u2197 " if merged_map[coord] is True else ""
    else:
        prefix = ""

    value = format_cell_value(raw_cell, value_cell, show_formula_values)
    if not value:
        return ""

    result = prefix + value

    if show_comments:
        comment = get_cell_comment(raw_cell)
        if comment:
            result += " <sup>[note]</sup>"

    return _md_escape(result)
