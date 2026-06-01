from __future__ import annotations

import posixpath
from typing import Iterable

PKG_REL_NS = {"pr": "http://schemas.openxmlformats.org/package/2006/relationships"}
NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "v": "urn:schemas-microsoft-com:vml",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
}

DRAWING_TAGS = {
    f"{{{NS['w']}}}drawing",
    f"{{{NS['w']}}}pict",
}

PARAGRAPH_TAG = f"{{{NS['w']}}}p"
ANCHOR_TAG = f"{{{NS['wp']}}}anchor"
INLINE_TAG = f"{{{NS['wp']}}}inline"


def normalize_zip_path(base_part: str, target: str) -> str:
    joined = posixpath.normpath(posixpath.join(posixpath.dirname(base_part), target))
    return joined.lstrip("/")


def xml_parts(names: Iterable[str]) -> list[str]:
    allowed_prefixes = ("word/document.xml", "word/header", "word/footer", "word/footnotes.xml", "word/endnotes.xml")
    return [
        name
        for name in names
        if name.endswith(".xml") and not name.startswith("word/_rels/") and name.startswith("word/") and name.startswith(allowed_prefixes)
    ]
