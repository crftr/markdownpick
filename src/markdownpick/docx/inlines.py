"""Run-level inline formatting for docx paragraphs.

Handles: bold, italic, strikethrough, underline, superscript, subscript, hyperlinks.
"""

from __future__ import annotations

from contextlib import suppress

from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph


def _run_to_inline(run) -> str:
    """Convert a single run to markdown inline text with formatting markers."""
    text = run.text
    if not text:
        return ""

    lpad = len(text) - len(text.lstrip())
    rpad = len(text) - len(text.rstrip())
    inner = text.strip()
    if not inner:
        return text

    # Strikethrough
    if run.font.strike:
        inner = f"~~{inner}~~"

    # Bold + italic combinations
    if run.bold and run.italic:
        inner = f"***{inner}***"
    elif run.bold:
        inner = f"**{inner}**"
    elif run.italic:
        inner = f"*{inner}*"

    # Underline (HTML in markdown)
    if run.underline:
        inner = f"<u>{inner}</u>"

    # Superscript / subscript
    try:
        if run.font.subscript:
            inner = f"<sub>{inner}</sub>"
        elif run.font.superscript:
            inner = f"<sup>{inner}</sup>"
    except Exception:
        pass

    return text[:lpad] + inner + (text[len(text) - rpad:] if rpad else "")


def _extract_hyperlink_runs(para: Paragraph) -> list[tuple[str, str]]:
    """Extract (display_text, url) pairs from <w:hyperlink> elements in a paragraph."""
    links: list[tuple[str, str]] = []
    p_elem = para._element
    for hyperlink_xml in p_elem.findall(qn("w:hyperlink")):
        rid = hyperlink_xml.get(qn("r:id"))
        if not rid:
            continue
        url = ""
        with suppress(Exception):
            url = para.part.rels[rid].target_ref if rid in para.part.rels else ""
        display_texts = []
        for run_xml in hyperlink_xml.findall(qn("w:r")):
            texts = run_xml.findall(qn("w:t"))
            for t_elem in texts:
                if t_elem.text:
                    display_texts.append(t_elem.text)
        if display_texts:
            links.append(("".join(display_texts), url))
    return links


def paragraph_to_md_with_hyperlinks(para: Paragraph) -> str:
    """Convert a paragraph to markdown, integrating hyperlinks.

    Replaces hyperlink display text with [text](url) markdown links.
    """
    hyperlinks = _extract_hyperlink_runs(para)
    if not hyperlinks:
        return "".join(_run_to_inline(r) for r in para.runs)

    # Build inline text, replacing hyperlink display texts with markdown links
    full_text = "".join(_run_to_inline(r) for r in para.runs)
    for display_text, url in hyperlinks:
        escaped_text = display_text.replace("[", "\\[").replace("]", "\\]")
        link_md = f"[{escaped_text}]({url})"
        full_text = full_text.replace(display_text, link_md, 1)

    return full_text
