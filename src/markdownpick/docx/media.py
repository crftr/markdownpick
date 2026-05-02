"""Image extraction from docx documents."""

from __future__ import annotations

import hashlib
from pathlib import Path

import docx
from docx.oxml.ns import qn

from markdownpick.errors import ImageExtractionError


def _blip_rid(blip_elem) -> str | None:
    """Extract the relationship ID from a <a:blip> element."""
    return blip_elem.get(qn("r:embed")) or blip_elem.get(qn("r:link"))


def _image_extension(part) -> str:
    """Determine file extension from an image part's content type."""
    content_type = part.content_type or ""
    mapping = {
        "image/png": "png",
        "image/jpeg": "jpg",
        "image/gif": "gif",
        "image/bmp": "bmp",
        "image/svg+xml": "svg",
        "image/tiff": "tiff",
        "image/webp": "webp",
    }
    return mapping.get(content_type, "png")


def extract_images(
    doc: docx.Document,
    images_dir: Path,
) -> list[tuple[int, str, str]]:
    """Extract all images from a docx document.

    Returns list of (paragraph_index, relative_path, alt_text).
    Images are saved to images_dir/ with deduplicated names.
    """
    images_dir.mkdir(parents=True, exist_ok=True)
    results: list[tuple[int, str, str]] = []
    seen: dict[str, int] = {}
    image_counter = 0

    # Find inline images
    for idx, para in enumerate(doc.paragraphs):
        try:
            p_elem = para._element
        except AttributeError:
            continue
        for inline_elem in p_elem.findall(qn("wp:inline")):
            blip_elems = inline_elem.findall(f".//{qn('a:blip')}")
            for blip in blip_elems:
                rid = _blip_rid(blip)
                if not rid:
                    continue
                try:
                    part = doc.part.related_parts[rid]
                    ext = _image_extension(part)
                    blob = part.blob

                    # Deduplicate by content hash
                    content_hash = hashlib.md5(blob).hexdigest()[:8]
                    if content_hash in seen:
                        rel_path = seen[content_hash]
                    else:
                        image_counter += 1
                        filename = f"img_{image_counter:03d}.{ext}"
                        rel_path = f"images/{filename}"
                        output_path = images_dir / filename
                        output_path.write_bytes(blob)
                        seen[content_hash] = rel_path

                    alt_text = _extract_alt_text(inline_elem, image_counter)
                    results.append((idx, rel_path, alt_text))
                except Exception as exc:
                    raise ImageExtractionError(
                        f"Failed to extract image in paragraph {idx}: {exc}",
                        element_type="image",
                        location=f"paragraph {idx}",
                    ) from exc

        # Also check floating images (anchors)
        for anchor_elem in p_elem.findall(qn("wp:anchor")):
            blip_elems = anchor_elem.findall(f".//{qn('a:blip')}")
            for blip in blip_elems:
                rid = _blip_rid(blip)
                if not rid:
                    continue
                try:
                    part = doc.part.related_parts[rid]
                    ext = _image_extension(part)
                    blob = part.blob

                    content_hash = hashlib.md5(blob).hexdigest()[:8]
                    if content_hash in seen:
                        rel_path = seen[content_hash]
                    else:
                        image_counter += 1
                        filename = f"img_{image_counter:03d}.{ext}"
                        rel_path = f"images/{filename}"
                        output_path = images_dir / filename
                        output_path.write_bytes(blob)
                        seen[content_hash] = rel_path

                    alt_text = _extract_alt_text(anchor_elem, image_counter)
                    results.append((idx, rel_path, alt_text))
                except Exception as exc:
                    raise ImageExtractionError(
                        f"Failed to extract floating image in paragraph {idx}: {exc}",
                        element_type="image",
                        location=f"paragraph {idx}",
                    ) from exc

    return results


def _extract_alt_text(element, image_num: int) -> str:
    """Extract alt text from inline or anchor element."""
    # Check docProperties
    for doc_pr in element.findall(qn("wp:docPr")):
        title = doc_pr.get("title", "")
        descr = doc_pr.get("descr", "")
        if title:
            return title
        if descr:
            return descr
    return f"Image {image_num}"


def images_to_markdown(extracted: list[tuple[int, str, str]]) -> list[tuple[int, str]]:
    """Convert extracted image list to markdown image references.

    Returns list of (paragraph_index, markdown_image_string).
    """
    return [
        (idx, f"![{alt}]({path})")
        for idx, path, alt in extracted
    ]
