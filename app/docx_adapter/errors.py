from __future__ import annotations

from dataclasses import dataclass


class DocumentError(Exception):
    pass


class UnsupportedDocumentError(DocumentError):
    pass


@dataclass(slots=True)
class NavigationResult:
    success: bool
    message: str
