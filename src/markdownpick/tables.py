"""Shared GFM table renderer."""

from __future__ import annotations

from collections.abc import Iterable


def markdown_table(headers: Iterable[str], rows: Iterable[Iterable[str]]) -> str:
    """Generate a GitHub-flavored markdown table.

    Pads short rows, escapes pipe characters in content.
    """
    header_list = [h.replace("|", "\\|") for h in headers]
    row_lists = [[c.replace("|", "\\|") for c in r] for r in rows]

    lines: list[str] = []
    lines.append("| " + " | ".join(header_list) + " |")
    lines.append("| " + " | ".join("---" for _ in header_list) + " |")

    width = len(header_list)
    for row in row_lists:
        safe = (row + [""] * max(0, width - len(row)))[:width]
        lines.append("| " + " | ".join(safe) + " |")

    return "\n".join(lines)
