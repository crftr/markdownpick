<p align="center">
	<img src=".assets/markdownpick-type-logo.avif" alt="markdownpick logo" />
</p>


Distill Word (.docx) and Excel (.xlsx) documents into clean, LLM-friendly Markdown.

## Features

- **Style-aware docx conversion** — Headings, lists, quotes, code mapped to GFM
- **Run-level formatting** — Bold, italic, strikethrough, underline, superscript, subscript
- **Hyperlink preservation** — Both docx and xlsx links become `[text](url)`
- **Image extraction** — Embedded images saved with markdown references
- **Formula preservation** — Excel formulas as `` `=FORMULA` `` with optional cached values
- **Merged cell detection** — Span markers in markdown tables
- **Formula dependency maps** — Track what cells reference what
- **Color & formatting maps** — 150+ named colors with nearest-color fallback
- **Footnote extraction** — docx footnotes preserved as markdown footnotes
- **Number/date formatting** — Respects Excel number formats for currency, dates, percentages
- **Error recovery** — Per-element fallbacks, corrupt file validation, detailed reports

## Usage

```bash
# Unified converter
mp spec.docx model.xlsx --output-dir ./docs

# Single-file converters
mp-docx spec.docx --images-dir ./images
mp-xlsx model.xlsx --show-formula-values

# As a Python module
python -m markdownpick spec.docx
```

## Install

### Quick install (uv, recommended)

```bash
uv tool install .
```

This installs `mp`, `mp-docx`, and `mp-xlsx` globally in an isolated environment managed by uv.

### Install with pip

```bash
pip install .
```

### Install with pipx (isolated, recommended for global CLI tools)

```bash
pipx install .
```

### Developer install

```bash
uv sync --dev
uv run mp --help
```

### Development mode (editable)

```bash
pip install -e ".[dev]"
```

### Requirements

- Python 3.12+
- Dependencies: `openpyxl`, `python-docx`

## Flags

### Shared (all file types)

| Flag | Description |
|---|---|
| `-d / --output-dir DIR` | Write all outputs to this directory |
| `--no-formatting-map` | Omit color & formatting maps |
| `--no-style-map` | (docx) Omit style map |
| `-v / --verbose` | Enable verbose logging |

### docx only

| Flag | Description |
|---|---|
| `--images-dir DIR` | Extract embedded images to this directory |

### xlsx only

| Flag | Description |
|---|---|
| `--show-formula-values` | Show cached formula results: `` `=FORMULA` ⇒ value `` |
| `--no-cell-refs` | Omit row number column |
| `--no-dependency-map` | Omit formula dependency maps |
| `--keep-empty-rows` | Keep trailing empty rows |
| `--keep-empty-columns` | Keep trailing empty columns |
| `--show-comments` | Include cell comments/notes |
