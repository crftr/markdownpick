#!/usr/bin/env python3
"""Convert Word documents (.docx) to Markdown.

Features:
- Style-aware conversion (headings, lists, quotes, code)
- Run-level formatting (bold, italic, strikethrough, underline, sup, sub)
- Hyperlink preservation
- Image extraction
- Rich table cell formatting
- Style map and color/formatting map appendices
- Error recovery per element
"""

from __future__ import annotations

import argparse
import sys
import zipfile
from collections import defaultdict
from pathlib import Path

import docx
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

from markdownpick.colors import docx_color_name
from markdownpick.docx.inlines import paragraph_to_md_with_hyperlinks
from markdownpick.docx.media import extract_images, images_to_markdown
from markdownpick.docx.styles import (
    _SKIP_STYLES,
    paragraph_to_md,
)
from markdownpick.errors import CorruptFileError
from markdownpick.frontmatter import build_frontmatter
from markdownpick.report import ConversionReport, safe_convert
from markdownpick.tables import markdown_table

_MAX_OCCURRENCES = 15


def validate_docx(path: Path) -> None:
    """Verify a file is a valid OOXML zip package."""
    try:
        with zipfile.ZipFile(path) as zf:
            if "[Content_Types].xml" not in zf.namelist():
                raise CorruptFileError(f"Not a valid OOXML package: {path}")
    except zipfile.BadZipFile as exc:
        raise CorruptFileError(f"File is not a valid zip archive: {path}") from exc


def _iter_body_elements(doc: docx.Document):
    """Yield paragraphs and tables from document body in XML order."""
    para_tag = qn("w:p")
    tbl_tag = qn("w:tbl")
    for child in doc.element.body:
        if child.tag == para_tag:
            yield Paragraph(child, doc)
        elif child.tag == tbl_tag:
            yield Table(child, doc)


def _table_to_md(table: Table, report: ConversionReport | None = None) -> str:
    """Convert a docx table to a GFM markdown table, preserving run formatting in cells."""
    rows_data: list[list[str]] = []
    for row in table.rows:
        cells: list[str] = []
        for cell in row.cells:
            cell_md = ""
            for para in cell.paragraphs:
                line = paragraph_to_md_with_hyperlinks(para)
                if cell_md:
                    cell_md += " " + line.strip()
                else:
                    cell_md = line.strip()
            cells.append(cell_md.replace("|", "\\|").replace("\n", " "))
        rows_data.append(cells)

    if not rows_data or not any(rows_data):
        return ""

    return markdown_table(rows_data[0], rows_data[1:])


def build_style_map(paragraphs: list[tuple[int, Paragraph]]) -> str:
    """Group paragraphs by named style, omitting Normal/default."""
    style_paras: dict[str, list[str]] = defaultdict(list)
    for idx, para in paragraphs:
        name = (para.style.name if para.style else "Normal").strip()
        if name.lower() in _SKIP_STYLES:
            continue
        style_paras[name].append(f"\u00b6{idx}")

    if not style_paras:
        return "_No named styles applied (all paragraphs use Normal/default)._"

    rows = [[name, ", ".join(pnums)] for name, pnums in sorted(style_paras.items())]
    return markdown_table(["Style", "Paragraphs"], rows)


def _run_format_descriptor(run) -> str | None:
    """Build a concise format label for a run."""
    parts: list[str] = []

    if run.font.highlight_color:
        highlight_names = {
            "YELLOW": "yellow highlight",
            "GREEN": "green highlight",
            "CYAN": "cyan highlight",
            "MAGENTA": "magenta highlight",
            "RED": "red highlight",
            "DARK_BLUE": "dark blue highlight",
            "DARK_CYAN": "dark cyan highlight",
            "DARK_GREEN": "dark green highlight",
            "DARK_MAGENTA": "dark magenta highlight",
            "DARK_RED": "dark red highlight",
            "DARK_YELLOW": "dark yellow highlight",
            "DARK_GRAY": "dark gray highlight",
            "LIGHT_GRAY": "light gray highlight",
            "WHITE": "white highlight",
            "BLACK": "black highlight",
        }
        parts.append(highlight_names.get(str(run.font.highlight_color), str(run.font.highlight_color)))
    elif run.font.color and run.font.color.rgb:
        label = docx_color_name(str(run.font.color.rgb))
        if label and label not in ("black", "near-black"):
            parts.append(f"{label} text")

    if run.bold:
        parts.append("bold")
    if run.italic:
        parts.append("italic")
    if run.font.strike:
        parts.append("strikethrough")
    if run.underline:
        parts.append("underline")

    return ", ".join(parts) if parts else None


def _truncate(text: str, max_len: int = 30) -> str:
    text = text.strip()
    return text if len(text) <= max_len else text[:max_len].rstrip() + "\u2026"


def build_formatting_map(paragraphs: list[tuple[int, Paragraph]]) -> str:
    """Group run-level formatting by descriptor; one entry per paragraph per format."""
    fmt_para_snippets: dict[str, dict[int, str]] = defaultdict(dict)

    for idx, para in paragraphs:
        for run in para.runs:
            if not run.text.strip():
                continue
            desc = _run_format_descriptor(run)
            if desc and idx not in fmt_para_snippets[desc]:
                fmt_para_snippets[desc][idx] = _truncate(run.text)

    if not fmt_para_snippets:
        return "_No notable run-level formatting found._"

    rows: list[list[str]] = []
    for desc, para_snippets in sorted(fmt_para_snippets.items()):
        occurrences = [f'\u00b6{idx} ("{snip}")' for idx, snip in sorted(para_snippets.items())]
        displayed = occurrences[:_MAX_OCCURRENCES]
        remainder = len(occurrences) - _MAX_OCCURRENCES
        cell = ", ".join(displayed)
        if remainder > 0:
            cell += f" (and {remainder} more)"
        rows.append([desc, cell])

    return markdown_table(["Format", "Occurrences"], rows)


def _build_frontmatter(
    docx_path: Path,
    doc: docx.Document,
    include_formatting_map: bool,
    images_dir: Path | None = None,
) -> list[str]:
    cp = doc.core_properties
    features = ["style map"]
    if include_formatting_map:
        features.append("color & formatting map")
    if images_dir:
        features.append("extracted images")

    lines: list[str] = []

    if cp.title and cp.title.strip() and cp.title.strip().lower() != "word document":
        lines.append(("title", cp.title.strip()))
    if cp.author and cp.author.strip():
        lines.append(("author", cp.author.strip()))
    if cp.created:
        lines.append(("created", cp.created.strftime("%Y-%m-%d")))
    if cp.modified:
        lines.append(("modified", cp.modified.strftime("%Y-%m-%d")))

    return build_frontmatter(
        source=docx_path.name,
        generator="markdownpick/docx",
        features=features,
        description=(
            "Markdown conversion of a Word document. Headings, lists, and tables "
            "are converted to GFM equivalents. A style map groups paragraphs by "
            "named style; a color and formatting map captures run-level visual "
            "semantics (font color, highlights, bold, italic, underline) with text snippets."
        ),
        **{k: v for k, v in lines},
    )


def _extract_footnotes(doc: docx.Document) -> str:
    """Extract footnotes from the document as a markdown section."""
    notes: list[tuple[str, str]] = []
    try:
        for footnote in doc.footnotes:
            for para in footnote.paragraphs:
                text = para.text.strip()
                if text:
                    notes.append((footnote.id, text))
    except Exception:
        return ""

    if not notes:
        return ""

    lines = ["## Footnotes", ""]
    for note_id, text in notes:
        lines.append(f"[^{note_id}]: {text}")
    lines.append("")
    return "\n".join(lines)


def convert_docx_to_markdown(
    docx_path: Path,
    output_path: Path,
    *,
    images_dir: Path | None = None,
    include_formatting_map: bool = True,
    include_style_map: bool = True,
    report: ConversionReport | None = None,
) -> None:
    """Read a .docx file and write a markdown file.

    Args:
        docx_path: Path to the .docx file.
        output_path: Where to write the markdown output.
        images_dir: If provided, extract images here.
        include_formatting_map: Include the color/formatting map appendix.
        include_style_map: Include the style map appendix.
        report: Optional report to collect warnings/errors.
    """
    validate_docx(docx_path)
    doc = docx.Document(str(docx_path))

    # Extract images first if requested
    image_markdown: dict[int, str] = {}
    if images_dir:
        try:
            extracted = extract_images(doc, images_dir)
            for idx, md_str in images_to_markdown(extracted):
                image_markdown[idx] = md_str
        except Exception as exc:
            if report:
                report.warn(f"Image extraction failed: {exc}")

    sections: list[str] = _build_frontmatter(docx_path, doc, include_formatting_map, images_dir)
    sections.append("")

    para_index = 0
    indexed_paragraphs: list[tuple[int, Paragraph]] = []

    for element in _iter_body_elements(doc):
        if isinstance(element, Paragraph):
            para_index += 1
            indexed_paragraphs.append((para_index, element))

            md_line = safe_convert(
                lambda e=element: paragraph_to_md(e, paragraph_to_md_with_hyperlinks(e)),
                fallback=element.text,
                report=report,
                label=f"paragraph {para_index}",
            )
            sections.append(md_line if md_line.strip() else "")

            # Insert image if this paragraph has one
            if para_index in image_markdown:
                sections.append(image_markdown[para_index])

        elif isinstance(element, Table):
            sections.append("")
            table_md = safe_convert(
                lambda e=element: _table_to_md(e, report),
                fallback="_[Table conversion failed]_",
                report=report,
                label=f"table at position {para_index}",
            )
            sections.append(table_md)
            sections.append("")

    if include_style_map:
        sections.append("")
        sections.append("## Style Map")
        sections.append("")
        sections.append(build_style_map(indexed_paragraphs))

    if include_formatting_map:
        sections.append("")
        sections.append("## Color & Formatting Map")
        sections.append("")
        sections.append(build_formatting_map(indexed_paragraphs))

    # Extract footnotes if present
    footnotes_md = safe_convert(
        lambda: _extract_footnotes(doc),
        fallback="",
        report=report,
        label="footnotes",
    )
    if footnotes_md:
        sections.append("")
        sections.append(footnotes_md)

    output_path.write_text("\n".join(sections).rstrip() + "\n", encoding="utf-8")


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert a Word document to Markdown with style and formatting maps."
    )
    parser.add_argument("docx_file", help="Path to .docx file")
    parser.add_argument(
        "-o", "--output",
        help="Output markdown file path (default: same name with .md)",
        default=None,
    )
    parser.add_argument(
        "--images-dir",
        help="Directory to extract embedded images into",
        default=None,
    )
    parser.add_argument(
        "--no-formatting-map",
        action="store_true",
        help="Do not include a color & formatting map section.",
    )
    parser.add_argument(
        "--no-style-map",
        action="store_true",
        help="Do not include a style map section.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    docx_path = Path(args.docx_file).expanduser().resolve()

    if not docx_path.exists():
        print(f"Error: file not found: {docx_path}", file=sys.stderr)
        return 1

    if docx_path.suffix.lower() != ".docx":
        print("Error: only .docx files are supported.", file=sys.stderr)
        return 1

    output_path = (
        Path(args.output).expanduser().resolve()
        if args.output
        else docx_path.with_suffix(".md")
    )

    images_dir = Path(args.images_dir).expanduser().resolve() if args.images_dir else None

    report = ConversionReport()
    try:
        convert_docx_to_markdown(
            docx_path=docx_path,
            output_path=output_path,
            images_dir=images_dir,
            include_formatting_map=not args.no_formatting_map,
            include_style_map=not args.no_style_map,
            report=report,
        )
        print(f"Wrote: {output_path}")
    except CorruptFileError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    if report.has_issues:
        print(f"\nConversion completed with {len(report.warnings)} warning(s) and {len(report.errors)} error(s).")
        for w in report.warnings:
            print(f"  Warning: {w}")
        for e in report.errors:
            print(f"  Error: {e}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
