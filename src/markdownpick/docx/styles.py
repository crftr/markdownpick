"""Style-to-Markdown mapping for docx paragraphs."""

from __future__ import annotations

from docx.text.paragraph import Paragraph

_HEADING_STYLES: dict[str, str] = {
    "heading 1": "#",
    "heading 2": "##",
    "heading 3": "###",
    "heading 4": "####",
    "heading 5": "#####",
    "heading 6": "######",
}

_LIST_BULLET_STYLES = {
    "list bullet", "list bullet 2", "list bullet 3",
    "list bullet 4", "list bullet 5",
}

_LIST_NUMBER_STYLES = {
    "list number", "list number 2", "list number 3",
    "list number 4", "list number 5",
}

_QUOTE_STYLES = {"quote", "intense quote", "block text"}
_CODE_STYLES = {"code", "html code", "macro text"}
_SKIP_STYLES = {"normal", "default paragraph font", "no spacing", "body text", ""}


def _style_name(para: Paragraph) -> str:
    return (para.style.name if para.style else "Normal").lower().strip()


def classify_paragraph(para: Paragraph) -> str:
    """Return the semantic classification of a paragraph."""
    name = _style_name(para)
    if name in _HEADING_STYLES:
        return "heading"
    if name in _LIST_BULLET_STYLES:
        return "list_bullet"
    if name in _LIST_NUMBER_STYLES:
        return "list_number"
    if name in _QUOTE_STYLES:
        return "quote"
    if name in _CODE_STYLES:
        return "code"
    return "paragraph"


def paragraph_to_md(para: Paragraph, inline_text: str) -> str:
    """Convert a paragraph to markdown based on its style classification.

    inline_text is the already-processed run-level text (bold/italic/etc).
    """
    name = _style_name(para)
    content = inline_text if inline_text.strip() else para.text
    stripped = content.strip()

    if name in _HEADING_STYLES:
        return f"{_HEADING_STYLES[name]} {stripped}"

    if name in _LIST_BULLET_STYLES:
        level = (int(name[-1]) - 1) if name[-1].isdigit() else 0
        return "  " * level + f"- {stripped}"

    if name in _LIST_NUMBER_STYLES:
        level = (int(name[-1]) - 1) if name[-1].isdigit() else 0
        return "  " * level + f"1. {stripped}"

    if name in _QUOTE_STYLES:
        return f"> {stripped}"

    if name in _CODE_STYLES:
        return f"`{stripped}`"

    return content


def should_skip_style(para: Paragraph) -> bool:
    """Return True if the paragraph's style is too generic to surface in a style map."""
    name = (para.style.name if para.style else "Normal").strip()
    return name.lower() in _SKIP_STYLES
