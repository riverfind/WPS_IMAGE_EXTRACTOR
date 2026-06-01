from .adapter import DocxAdapter
from .errors import DocumentError, NavigationResult, UnsupportedDocumentError

__all__ = [
    "DocxAdapter",
    "DocumentError",
    "NavigationResult",
    "UnsupportedDocumentError",
]
