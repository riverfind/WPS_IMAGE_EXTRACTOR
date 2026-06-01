from __future__ import annotations

import hashlib
import os
import shutil
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from app.models import ImageAsset, ImageLocation

from .constants import PARAGRAPH_TAG, xml_parts
from .errors import DocumentError, UnsupportedDocumentError
from .office import OfficeAutomationMixin
from .xml_ops import XmlImageOpsMixin


class DocxAdapter(XmlImageOpsMixin, OfficeAutomationMixin):
    name = "docx"

    def __init__(self, document_path: Path) -> None:
        self.document_path = document_path.resolve()

    @staticmethod
    def can_open(document_path: Path) -> bool:
        return document_path.suffix.lower() in {".docx", ".docm", ".dotx"}

    def extract_images(self) -> list[ImageAsset]:
        if not self.can_open(self.document_path):
            raise UnsupportedDocumentError("当前版本仅支持可解析的 .docx/.docm/.dotx 文档。")

        images: list[ImageAsset] = []
        with zipfile.ZipFile(self.document_path, "r") as archive:
            names = archive.namelist()
            for part_name in xml_parts(names):
                rels_map = self._read_relationships(archive, part_name)
                if not rels_map:
                    continue
                root = ET.fromstring(archive.read(part_name))
                parent_map = {child: parent for parent in root.iter() for child in parent}
                paragraphs = list(root.iter(PARAGRAPH_TAG))
                paragraph_indexes = self._build_paragraph_indexes(root)
                part_occurrence_index = 0
                office_collection_counters: dict[str, int] = {}

                for image_node, rel_id in self._iter_image_refs(root):
                    occurrence_index = part_occurrence_index
                    part_occurrence_index += 1
                    media_path = rels_map.get(rel_id)
                    if not media_path or media_path not in names:
                        continue

                    image_bytes = archive.read(media_path)
                    width, height = self._read_image_size(image_bytes)
                    anchor_type = self._detect_anchor_type(image_node, parent_map)
                    office_collection = self._office_collection_name(anchor_type)
                    office_collection_index: int | None = None
                    if office_collection:
                        office_collection_index = office_collection_counters.get(office_collection, 0) + 1
                        office_collection_counters[office_collection] = office_collection_index
                    paragraph_index = self._find_paragraph_index(image_node, parent_map, paragraph_indexes)
                    image_id = self._build_stable_image_id(
                        image_node,
                        parent_map,
                        part_name,
                        rel_id,
                        media_path,
                        anchor_type,
                        paragraph_index,
                    )
                    paragraph_text, context_before, context_after, text_hint = self._extract_text_context(
                        image_node,
                        parent_map,
                        paragraphs,
                        paragraph_indexes,
                    )
                    md5 = hashlib.md5(image_bytes).hexdigest()
                    file_name = Path(media_path).name

                    images.append(
                        ImageAsset(
                            id=image_id,
                            name=file_name,
                            extension=Path(file_name).suffix.lower() or ".bin",
                            media_path=media_path,
                            md5=md5,
                            width=width,
                            height=height,
                            size_bytes=len(image_bytes),
                            image_bytes=image_bytes,
                            location=ImageLocation(
                                part_name=part_name,
                                rel_id=rel_id,
                                anchor_type=anchor_type,
                                block_index=paragraph_index,
                                occurrence_index=occurrence_index,
                                text_hint=text_hint,
                                paragraph_text=paragraph_text,
                                context_before=context_before,
                                context_after=context_after,
                                office_collection=office_collection,
                                office_collection_index=office_collection_index,
                            ),
                        )
                    )

        return images

    def delete_images(self, images: list[ImageAsset]) -> tuple[Path, int]:
        if not images:
            raise DocumentError("没有可删除的图片。")

        self._ensure_document_not_open_for_write()
        before_count = len(self.extract_images())

        by_part: dict[str, set[int]] = {}
        for image in images:
            by_part.setdefault(image.location.part_name, set()).add(image.location.occurrence_index)

        backup_path = self.document_path.with_suffix(self.document_path.suffix + ".bak")
        if not backup_path.exists():
            shutil.copy2(self.document_path, backup_path)

        temp_fd, temp_name = tempfile.mkstemp(suffix=self.document_path.suffix)
        os.close(temp_fd)
        temp_file = Path(temp_name)
        try:
            with zipfile.ZipFile(self.document_path, "r") as source:
                updated_parts: dict[str, bytes] = {}
                for part_name, target_indexes in by_part.items():
                    root = ET.fromstring(source.read(part_name))
                    parent_map = {child: parent for parent in root.iter() for child in parent}
                    current_index = 0
                    removed = 0
                    for image_node, _ in self._iter_image_refs(root):
                        if current_index in target_indexes:
                            container = self._find_removal_container(image_node, parent_map)
                            parent = parent_map.get(container) if container is not None else None
                            if container is not None and parent is not None:
                                parent.remove(container)
                                removed += 1
                        current_index += 1

                    if removed != len(target_indexes):
                        missing = len(target_indexes) - removed
                        raise DocumentError(f"删除文档图片时有 {missing} 项未能匹配到 XML 节点。")

                    updated_parts[part_name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

                with zipfile.ZipFile(temp_file, "w") as target:
                    for item in source.infolist():
                        data = updated_parts.get(item.filename, source.read(item.filename))
                        target.writestr(item, data)

            os.replace(temp_file, self.document_path)
        finally:
            if temp_file.exists():
                temp_file.unlink(missing_ok=True)

        after_count = len(self.extract_images())
        expected_after = before_count - len(images)
        if after_count != expected_after:
            raise DocumentError(
                f"删除写回后的图片数量异常：删除前 {before_count} 张，预期剩余 {expected_after} 张，实际剩余 {after_count} 张。"
            )

        return backup_path, len(images)
