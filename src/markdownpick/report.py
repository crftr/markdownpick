"""Shared utilities: error-tolerant conversion reporting."""

from __future__ import annotations

import logging
from collections.abc import Callable
from dataclasses import dataclass, field

from markdownpick.errors import ConversionError

logger = logging.getLogger(__name__)


@dataclass
class ConversionReport:
    """Tracks warnings and errors encountered during conversion."""

    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)

    @property
    def has_issues(self) -> bool:
        return bool(self.errors or self.warnings)

    def warn(self, message: str) -> None:
        self.warnings.append(message)
        logger.warning(message)

    def error(self, message: str) -> None:
        self.errors.append(message)
        logger.error(message)


def safe_convert[T](
    func: Callable[[], T],
    fallback: T,
    report: ConversionReport | None = None,
    label: str = "",
) -> T:
    """Call func, returning fallback on any error. Logs to report if provided."""
    try:
        return func()
    except ConversionError as exc:
        msg = f"[{label}] {exc}"
        if report:
            report.warn(msg)
        else:
            logger.warning(msg)
        return fallback
    except Exception as exc:
        msg = f"[{label}] unexpected error: {exc}"
        if report:
            report.error(msg)
        else:
            logger.error(msg)
        return fallback
