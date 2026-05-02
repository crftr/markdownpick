"""Cell formatting maps for Excel workbooks."""

from __future__ import annotations

from collections import defaultdict

from openpyxl.cell.cell import Cell

from markdownpick.colors import xlsx_color_name
from markdownpick.tables import markdown_table


def _color_label(rgb: str | None) -> str | None:
    """Resolve an ARGB color string to a human-readable name."""
    if rgb is None:
        return None
    rgb_str = str(rgb).upper()
    # Validate: should be 6 or 8 hex chars
    import re
    if not re.match(r"^[0-9A-F]{6,8}$", rgb_str):
        return None
    return xlsx_color_name(rgb_str)


def _cell_format_descriptor(cell: Cell) -> str | None:
    """Build a concise human-readable format descriptor for a cell."""
    parts: list[str] = []

    fill = cell.fill
    if fill.patternType and fill.patternType != "none" and fill.fgColor:
        fg_rgb = getattr(fill.fgColor, "rgb", None)
        if fg_rgb is not None:
            bg = _color_label(str(fg_rgb))
            if bg and bg != "white":
                parts.append(f"{bg} background")

    font = cell.font
    fc_rgb = getattr(font.color, "rgb", None) if font.color else None
    if fc_rgb is not None:
        fc = _color_label(str(fc_rgb))
        if fc:
            parts.append(f"{fc} text")
    if font.bold:
        parts.append("bold")
    if font.italic:
        parts.append("italic")
    if font.strikethrough:
        parts.append("strikethrough")
    if font.underline:
        parts.append("underline")

    return ", ".join(parts) if parts else None


def _collapse_ranges(coords: list[str]) -> str:
    """Collapse a sorted list of cell coordinates into compact range notation.

    Groups cells sharing the same column into contiguous row ranges:
    ["B6", "B7", "B8", "B13", "D6"] -> "B6:B8, B13, D6"
    """
    import re

    col_rows: dict[str, list[int]] = defaultdict(list)
    coord_pattern = re.compile(r"^([A-Z]+)(\d+)$")
    non_standard: list[str] = []

    for coord in coords:
        m = coord_pattern.match(coord)
        if m:
            col_rows[m.group(1)].append(int(m.group(2)))
        else:
            non_standard.append(coord)

    parts: list[str] = []
    for col in sorted(col_rows):
        rows = sorted(col_rows[col])
        i = 0
        while i < len(rows):
            start = rows[i]
            end = start
            while i + 1 < len(rows) and rows[i + 1] == end + 1:
                i += 1
                end = rows[i]
            if start == end:
                parts.append(f"{col}{start}")
            else:
                parts.append(f"{col}{start}:{col}{end}")
            i += 1

    parts.extend(non_standard)
    return ", ".join(parts)


def build_formatting_map(sheet, max_row: int, max_col: int) -> str:
    """Build a compact markdown map grouping cells by their visual formatting."""
    fmt_cells: dict[str, list[str]] = defaultdict(list)

    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            cell = sheet.cell(row=r, column=c)
            if cell.value is None:
                continue
            desc = _cell_format_descriptor(cell)
            if desc:
                fmt_cells[desc].append(cell.coordinate)

    if not fmt_cells:
        return "_No cell formatting found on this sheet._"

    rows: list[list[str]] = []
    for desc, cells in sorted(fmt_cells.items(), key=lambda kv: kv[0]):
        rows.append([desc, _collapse_ranges(cells)])

    return markdown_table(["Format", "Cells"], rows)
