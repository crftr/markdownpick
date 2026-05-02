"""Shared YAML frontmatter builder."""

from __future__ import annotations

from datetime import UTC, datetime


def build_frontmatter(
    source: str,
    generator: str,
    **metadata: str | list[str],
) -> list[str]:
    """Build YAML frontmatter lines.

    Automatically adds source, generated timestamp, and generator fields.
    Additional metadata is appended in sorted key order.
    """
    lines = [
        "---",
        f"source: {source}",
        f"generated: {datetime.now(UTC).strftime('%Y-%m-%dT%H:%M:%SZ')}",
        f"generator: {generator}",
    ]

    for key in sorted(metadata.keys()):
        value = metadata[key]
        if isinstance(value, list):
            items = ", ".join(value)
            lines.append(f"{key}: [{items}]")
        else:
            lines.append(f"{key}: {value}")

    lines.append("---")
    return lines
