"""Custom exception hierarchy for markdownpick."""

from __future__ import annotations


class MarkdownPickError(Exception):
    """Base exception for all markdownpick errors."""


class FileError(MarkdownPickError):
    """File not found, unreadable, or invalid format."""


class CorruptFileError(FileError):
    """File is corrupt, encrypted, or not a valid OOXML document."""


class ConversionError(MarkdownPickError):
    """Error during conversion of a document element."""

    def __init__(self, message: str, element_type: str = "unknown", location: str = "") -> None:
        super().__init__(message)
        self.element_type = element_type
        self.location = location


class ImageExtractionError(ConversionError):
    """Failed to extract or save an embedded image."""


class FormulaParseError(ConversionError):
    """Failed to parse an Excel formula."""
