#!/usr/bin/env python3
"""Unified multi-file converter: .xlsx and .docx → Markdown.

Dispatches each input file to the appropriate converter based on extension.
All output files are written with the same base name and a .md extension.

Usage:
  mp FILE [FILE ...] [-d OUTPUT_DIR] [flags]
  markdownpick FILE [FILE ...] [-d OUTPUT_DIR] [flags]
"""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from markdownpick.docx.converter import convert_docx_to_markdown
from markdownpick.excel.converter import convert_excel_to_markdown
from markdownpick.report import ConversionReport

XLSX_ONLY_FLAGS = {
    "show_formula_values",
    "no_cell_refs",
    "no_dependency_map",
    "keep_empty_rows",
    "keep_empty_columns",
    "show_comments",
}


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert one or more .xlsx / .docx files to Markdown.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Flags marked (xlsx only) are silently ignored for .docx files.\n\n"
            "Examples:\n"
            "  mp model.xlsx user_stories.docx\n"
            "  mp *.xlsx --output-dir ./docs --show-formula-values\n"
            "  mp spec.docx --no-formatting-map\n"
        ),
    )

    parser.add_argument(
        "files",
        nargs="+",
        help="Input files (.xlsx and/or .docx).",
    )
    parser.add_argument(
        "-d", "--output-dir",
        help="Write all output .md files to this directory (default: same dir as input).",
        default=None,
        metavar="DIR",
    )

    # Shared flags
    parser.add_argument(
        "--no-formatting-map",
        action="store_true",
        help="Do not include a color & formatting map section.",
    )
    parser.add_argument(
        "--no-style-map",
        action="store_true",
        help="(docx) Do not include a style map section.",
    )

    # xlsx-only flags
    xlsx_group = parser.add_argument_group("xlsx-only options")
    xlsx_group.add_argument(
        "--show-formula-values",
        action="store_true",
        help="Include cached formula results as: `=FORMULA` \u21d2 value.",
    )
    xlsx_group.add_argument(
        "--no-cell-refs",
        action="store_true",
        help="Omit the row number reference column.",
    )
    xlsx_group.add_argument(
        "--no-dependency-map",
        action="store_true",
        help="Do not include a formula dependency map section.",
    )
    xlsx_group.add_argument(
        "--keep-empty-rows",
        action="store_true",
        help="Keep trailing completely empty rows instead of trimming them.",
    )
    xlsx_group.add_argument(
        "--keep-empty-columns",
        action="store_true",
        help="Keep trailing completely empty columns instead of trimming them.",
    )
    xlsx_group.add_argument(
        "--show-comments",
        action="store_true",
        help="Include cell comments/notes.",
    )

    # General options
    parser.add_argument(
        "--images-dir",
        help="(docx) Directory to extract embedded images into.",
        default=None,
    )
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Enable verbose logging output.",
    )

    return parser.parse_args(argv)


def _warn_ignored_flags(args: argparse.Namespace, ext: str) -> list[str]:
    """Check if xlsx-only flags were used on a non-xlsx file."""
    if ext != ".xlsx":
        ignored = []
        for flag in XLSX_ONLY_FLAGS:
            if getattr(args, flag, False):
                flag_display = flag.replace("_", "-")
                ignored.append(f"--{flag_display}")
        return ignored
    return []


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])

    if args.verbose:
        logging.basicConfig(level=logging.DEBUG, format="%(levelname)s: %(message)s")

    output_dir: Path | None = None
    if args.output_dir:
        output_dir = Path(args.output_dir).expanduser().resolve()
        try:
            output_dir.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            print(f"Error: cannot create output directory: {exc}", file=sys.stderr)
            return 1

    all_reports: dict[str, ConversionReport] = {}
    errors: list[str] = []
    success_count = 0

    for file_str in args.files:
        input_path = Path(file_str).expanduser().resolve()

        if not input_path.exists():
            errors.append(f"File not found: {input_path}")
            continue

        ext = input_path.suffix.lower()
        if ext not in (".xlsx", ".docx"):
            errors.append(f"Unsupported file type {input_path.suffix!r}: {input_path.name}")
            continue

        output_path = (
            output_dir / input_path.with_suffix(".md").name
            if output_dir
            else input_path.with_suffix(".md")
        )

        report = ConversionReport()

        # Warn about ignored flags
        ignored = _warn_ignored_flags(args, ext)
        for flag_name in ignored:
            msg = f"{input_path.name}: {flag_name} is only applicable to .xlsx files"
            report.warn(msg)
            print(f"Warning: {msg}", file=sys.stderr)

        try:
            if ext == ".xlsx":
                convert_excel_to_markdown(
                    excel_path=input_path,
                    output_path=output_path,
                    include_cell_refs=not args.no_cell_refs,
                    show_formula_values=args.show_formula_values,
                    trim_trailing_empty_rows=not args.keep_empty_rows,
                    trim_trailing_empty_cols=not args.keep_empty_columns,
                    include_dependency_map=not args.no_dependency_map,
                    include_formatting_map=not args.no_formatting_map,
                    show_comments=args.show_comments,
                    report=report,
                )
            else:
                images_dir = (
                    Path(args.images_dir).expanduser().resolve()
                    if args.images_dir
                    else output_path.parent / f"{input_path.stem}_images"
                    if args.no_formatting_map is False
                    else None
                )
                convert_docx_to_markdown(
                    docx_path=input_path,
                    output_path=output_path,
                    images_dir=images_dir,
                    include_formatting_map=not args.no_formatting_map,
                    include_style_map=not args.no_style_map,
                    report=report,
                )

            all_reports[input_path.name] = report
            success_count += 1
            print(f"Wrote: {output_path}")
        except Exception as exc:
            errors.append(f"Failed to convert {input_path.name}: {exc}")

    if errors:
        print(f"\n{len(errors)} file(s) failed:", file=sys.stderr)
        for msg in errors:
            print(f"  Error: {msg}", file=sys.stderr)

    # Summary
    total = success_count + len(errors)
    print(f"\nDone: {success_count}/{total} file(s) converted successfully.")

    # Show accumulated warnings
    all_warnings = []
    for name, rpt in all_reports.items():
        for w in rpt.warnings:
            all_warnings.append(f"  [{name}] {w}")
    if all_warnings:
        print(f"\n{len(all_warnings)} warning(s):")
        for w in all_warnings:
            print(w)

    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
