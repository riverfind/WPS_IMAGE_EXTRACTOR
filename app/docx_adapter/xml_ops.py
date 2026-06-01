from __future__ import annotations

import hashlib
import io
import posixpath
import zipfile
from typing import Iterable
from xml.etree import ElementTree as ET

from PIL import Image

from app.models import ImageLocation

from .constants import ANCHOR_TAG, DRAWING_TAGS, INLINE_TAG, NS, PARAGRAPH_TAG, PKG_REL_NS, normalize_zip_path


class XmlImageOpsMixin:
    def _read_relationships(self, archive: zipfile.ZipFile, part_name: str) -> dict[str, str]:
        rel_name = posixpath.join(posixpath.dirname(part_name), "_rels", f"{posixpath.basename(part_name)}.rels")
        if rel_name not in archive.namelist():
            return {}

        root = ET.fromstring(archive.read(rel_name))
        rels_map: dict[str, str] = {}
        for rel in root.findall("pr:Relationship", PKG_REL_NS):
            rel_id = rel.attrib.get("Id")
            target = rel.attrib.get("Target")
            if not rel_id or not target:
                continue
            rels_map[rel_id] = normalize_zip_path(part_name, target)
        return rels_map

    def _iter_image_refs(self, root: ET.Element) -> Iterable[tuple[ET.Element, str]]:
        blip_tag = f"{{{NS['a']}}}blip"
        imagedata_tag = f"{{{NS['v']}}}imagedata"
        embed_attr = f"{{{NS['r']}}}embed"
        image_id_attr = f"{{{NS['r']}}}id"

        # Preserve the real XML traversal order so thumbnail listing, deletion,
        # and Office collection indexes all point at the same visual object.
        for node in root.iter():
            if node.tag == blip_tag:
                rel_id = node.attrib.get(embed_attr)
                if rel_id:
                    yield node, rel_id
            elif node.tag == imagedata_tag:
                rel_id = node.attrib.get(image_id_attr)
                if rel_id:
                    yield node, rel_id

    def _build_paragraph_indexes(self, root: ET.Element) -> dict[ET.Element, int]:
        return {paragraph: index for index, paragraph in enumerate(root.iter(PARAGRAPH_TAG), start=1)}

    def _find_paragraph_index(
        self,
        node: ET.Element,
        parent_map: dict[ET.Element, ET.Element],
        paragraph_indexes: dict[ET.Element, int],
    ) -> int | None:
        current = node
        while current in parent_map:
            current = parent_map[current]
            if current.tag == PARAGRAPH_TAG:
                return paragraph_indexes.get(current)
        return None

    def _extract_text_context(
        self,
        node: ET.Element,
        parent_map: dict[ET.Element, ET.Element],
        paragraphs: list[ET.Element],
        paragraph_indexes: dict[ET.Element, int],
    ) -> tuple[str, str, str, str]:
        paragraph = self._find_parent_paragraph(node, parent_map)
        if paragraph is None:
            return "", "", "", ""

        paragraph_text = self._normalize_text(self._paragraph_text(paragraph), limit=80)
        paragraph_index = paragraph_indexes.get(paragraph)
        if paragraph_index is None:
            return paragraph_text, "", "", paragraph_text

        context_before = self._collect_neighbor_texts(paragraphs, paragraph_index - 1, step=-1)
        context_after = self._collect_neighbor_texts(paragraphs, paragraph_index - 1, step=1)

        merged = self._normalize_text(
            " ".join(part for part in (context_before, paragraph_text, context_after) if part),
            limit=180,
        )
        return paragraph_text, context_before, context_after, merged

    def _find_parent_paragraph(
        self,
        node: ET.Element,
        parent_map: dict[ET.Element, ET.Element],
    ) -> ET.Element | None:
        current = node
        while current in parent_map:
            current = parent_map[current]
            if current.tag == PARAGRAPH_TAG:
                return current
        return None

    def _paragraph_text(self, paragraph: ET.Element) -> str:
        text_tag = f"{{{NS['w']}}}t"
        parts = []
        for text_node in paragraph.iter(text_tag):
            if text_node.text:
                parts.append(text_node.text.strip())
        return " ".join(part for part in parts if part)

    def _normalize_text(self, text: str, limit: int) -> str:
        normalized = " ".join(text.split())
        return normalized[:limit]

    def _collect_neighbor_texts(
        self,
        paragraphs: list[ET.Element],
        paragraph_offset: int,
        step: int,
        *,
        count: int = 4,
    ) -> str:
        collected: list[str] = []
        index = paragraph_offset + step
        while 0 <= index < len(paragraphs) and len(collected) < count:
            text = self._normalize_text(self._paragraph_text(paragraphs[index]), limit=60)
            if text:
                if step < 0:
                    collected.insert(0, text)
                else:
                    collected.append(text)
            index += step
        return " ".join(collected)

    def _find_removal_container(
        self,
        node: ET.Element,
        parent_map: dict[ET.Element, ET.Element],
    ) -> ET.Element | None:
        current = node
        while current in parent_map:
            current = parent_map[current]
            if current.tag in DRAWING_TAGS:
                return current
        return None

    def _build_stable_image_id(
        self,
        node: ET.Element,
        parent_map: dict[ET.Element, ET.Element],
        part_name: str,
        rel_id: str,
        media_path: str,
        anchor_type: str,
        paragraph_index: int | None,
    ) -> str:
        container = self._find_removal_container(node, parent_map)
        node_xml = ET.tostring(node, encoding="utf-8")
        container_xml = ET.tostring(container, encoding="utf-8") if container is not None else b""
        payload = b"|".join(
            [
                part_name.encode("utf-8", errors="ignore"),
                rel_id.encode("utf-8", errors="ignore"),
                media_path.encode("utf-8", errors="ignore"),
                anchor_type.encode("utf-8", errors="ignore"),
                str(paragraph_index or 0).encode("ascii"),
                node_xml,
                container_xml,
            ]
        )
        digest = hashlib.sha1(payload).hexdigest()[:20]
        return f"{part_name}::{digest}"

    def _detect_anchor_type(self, node: ET.Element, parent_map: dict[ET.Element, ET.Element]) -> str:
        current = node
        while current in parent_map:
            current = parent_map[current]
            if current.tag == ANCHOR_TAG:
                return "floating"
            if current.tag == INLINE_TAG:
                return "inline"
            if current.tag == f"{{{NS['w']}}}pict":
                return "vml"
        return "unknown"

    def _office_collection_name(self, anchor_type: str) -> str:
        if anchor_type == "inline":
            return "InlineShapes"
        if anchor_type in {"floating", "vml"}:
            return "Shapes"
        return ""

    def _read_image_size(self, image_bytes: bytes) -> tuple[int, int]:
        with Image.open(io.BytesIO(image_bytes)) as image:
            return image.size

    def _build_text_candidates(self, location: ImageLocation) -> list[str]:
        raw_candidates = [
            location.text_hint,
            self._normalize_text(f"{location.context_before} {location.paragraph_text}", limit=80),
            self._normalize_text(f"{location.paragraph_text} {location.context_after}", limit=80),
            self._normalize_text(f"{location.context_before} {location.context_after}", limit=80),
            location.paragraph_text,
        ]
        candidates: list[str] = []
        seen: set[str] = set()
        for candidate in raw_candidates:
            normalized = " ".join(candidate.split())
            if len(normalized) < 4 or normalized in seen:
                continue
            seen.add(normalized)
            candidates.append(normalized)
        return candidates
