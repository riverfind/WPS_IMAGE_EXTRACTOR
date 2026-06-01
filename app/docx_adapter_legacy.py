from __future__ import annotations

import hashlib
import io
import os
import posixpath
import shutil
import tempfile
import time
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET

from PIL import Image

from app.models import ImageAsset, ImageLocation

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


class DocumentError(Exception):
    pass


class UnsupportedDocumentError(DocumentError):
    pass


@dataclass(slots=True)
class NavigationResult:
    success: bool
    message: str


def _normalize_zip_path(base_part: str, target: str) -> str:
    joined = posixpath.normpath(posixpath.join(posixpath.dirname(base_part), target))
    return joined.lstrip("/")


def _xml_parts(names: Iterable[str]) -> list[str]:
    allowed_prefixes = ("word/document.xml", "word/header", "word/footer", "word/footnotes.xml", "word/endnotes.xml")
    return [
        name
        for name in names
        if name.endswith(".xml") and not name.startswith("word/_rels/") and name.startswith("word/") and name.startswith(allowed_prefixes)
    ]


class DocxAdapter:
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
            for part_name in _xml_parts(names):
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

    def locate(self, image: ImageAsset) -> NavigationResult:
        try:
            import pythoncom  # type: ignore
            import win32com.client  # type: ignore
        except Exception:
            pythoncom = None
            win32com_client = None
        else:
            win32com_client = win32com.client

        if win32com_client is not None and image.location.part_name == "word/document.xml" and image.location.block_index:
            try:
                pythoncom.CoInitialize()
                app, app_name = self._open_writer_app(win32com_client)
                document = self._open_or_get_document(app, self.document_path)
                paragraph_index = min(max(1, image.location.block_index), document.Paragraphs.Count)
                paragraph_range = document.Paragraphs(paragraph_index).Range
                bookmark_name = "__wps_image_locator__"
                try:
                    document.Activate()
                except Exception:
                    pass
                try:
                    if document.Bookmarks.Exists(bookmark_name):
                        document.Bookmarks(bookmark_name).Delete()
                except Exception:
                    pass
                exact_located = self._locate_by_office_collection(document, app, image.location)
                text_located = False
                if not exact_located and image.location.text_hint:
                    text_located = self._locate_by_text_hint(document, app, image.location)
                try:
                    if not exact_located and not text_located:
                        document.Bookmarks.Add(bookmark_name, paragraph_range)
                except Exception:
                    pass
                try:
                    if not exact_located and not text_located:
                        app.Selection.GoTo(What=-1, Name=bookmark_name)
                except Exception:
                    pass
                try:
                    if not exact_located and not text_located:
                        paragraph_range.Select()
                except Exception:
                    pass
                try:
                    if not exact_located and not text_located:
                        app.Selection.SetRange(paragraph_range.Start, paragraph_range.End)
                except Exception:
                    pass
                try:
                    if not exact_located and not text_located:
                        app.Selection.Collapse(1)
                except Exception:
                    pass
                try:
                    if not exact_located and not text_located:
                        document.ActiveWindow.ScrollIntoView(app.Selection.Range, True)
                except Exception:
                    pass
                try:
                    if not exact_located and not text_located:
                        app.ActiveWindow.ScrollIntoView(app.Selection.Range, True)
                except Exception:
                    pass
                try:
                    if not exact_located and not text_located:
                        app.Selection.Range.Select()
                except Exception:
                    pass
                app.Visible = True
                try:
                    app.Activate()
                except Exception:
                    pass
                try:
                    if document.Bookmarks.Exists(bookmark_name):
                        document.Bookmarks(bookmark_name).Delete()
                except Exception:
                    pass
                if exact_located:
                    collection_label = image.location.office_collection_index or image.location.occurrence_index + 1
                    return NavigationResult(True, f"已通过 {app_name} 按图片对象序号定位到第 {collection_label} 项。")
                if text_located:
                    return NavigationResult(True, f"已通过 {app_name} 按文本锚点尝试定位。")
                return NavigationResult(True, f"已通过 {app_name} 尝试定位到第 {paragraph_index} 段附近。")
            except Exception as exc:  # pragma: no cover - depends on local Office/WPS environment.
                return NavigationResult(False, f"自动定位失败：{exc}")
            finally:
                if pythoncom is not None:
                    try:
                        pythoncom.CoUninitialize()
                    except Exception:
                        pass

        try:
            os.startfile(self.document_path)  # type: ignore[attr-defined]
        except OSError as exc:
            return NavigationResult(False, f"无法打开文档：{exc}")

        return NavigationResult(False, f"已打开文档，请按位置提示手动定位：{image.location.display_text}")

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
            rels_map[rel_id] = _normalize_zip_path(part_name, target)
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

    def _open_writer_app(self, win32com_client) -> tuple[object, str]:
        errors: list[str] = []
        for prog_id, display_name in (
            ("kwps.Application", "WPS Writer"),
            ("Word.Application", "Microsoft Word"),
        ):
            try:
                app = win32com_client.GetActiveObject(prog_id)
            except Exception as active_exc:
                try:
                    app = win32com_client.Dispatch(prog_id)
                except Exception as dispatch_exc:
                    errors.append(f"{display_name}: {active_exc}; {dispatch_exc}")
                    continue
            try:
                return app, display_name
            except Exception as exc:
                errors.append(f"{display_name}: {exc}")
        detail = "；".join(errors) if errors else "未检测到可用的 COM 应用。"
        raise DocumentError(f"未找到可用于定位的 WPS/Word 应用。{detail}")

    def _ensure_document_not_open_for_write(self) -> None:
        try:
            import pythoncom  # type: ignore
            import win32com.client  # type: ignore
        except Exception:
            return

        try:
            pythoncom.CoInitialize()
            for prog_id, display_name in (
                ("kwps.Application", "WPS Writer"),
                ("Word.Application", "Microsoft Word"),
            ):
                try:
                    app = win32com.client.GetActiveObject(prog_id)
                except Exception:
                    continue
                open_document = self._get_open_document(app, self.document_path)
                if open_document is not None:
                    raise DocumentError(f"检测到文档当前仍在 {display_name} 中打开，请先关闭该文档后再执行删除。")
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def _get_open_document(self, app, document_path: Path):
        resolved_path = document_path.resolve()
        normalized_target = os.path.normcase(os.path.normpath(str(resolved_path)))
        try:
            documents = app.Documents
        except Exception:
            return None

        for index in range(1, documents.Count + 1):
            current = documents.Item(index)
            full_name = os.path.normcase(os.path.normpath(str(getattr(current, "FullName", ""))))
            if full_name == normalized_target:
                return current
        return None

    def _open_or_get_document(self, app, document_path: Path):
        try:
            documents = app.Documents
        except Exception as exc:
            raise DocumentError(f"当前应用不支持文档自动化接口：{exc}") from exc

        current = self._get_open_document(app, document_path)
        if current is not None:
            try:
                current.Activate()
            except Exception:
                pass
            return current

        resolved_path = document_path.resolve()
        document = documents.Open(str(resolved_path))
        for _ in range(20):
            try:
                _ = document.Paragraphs.Count
                try:
                    document.Activate()
                except Exception:
                    pass
                return document
            except Exception:
                time.sleep(0.1)
        return document

    def _locate_by_text_hint(self, document, app, location: ImageLocation) -> bool:
        candidates = self._build_text_candidates(location)
        if not candidates:
            return False
        try:
            paragraph_index = location.block_index
            if paragraph_index is not None:
                for radius in (0, 2, 6, 12):
                    paragraph_range = self._paragraph_window_range(document, paragraph_index, radius)
                    if paragraph_range is not None and self._find_first_candidate(document, paragraph_range, candidates, app):
                        return True
            return self._find_first_candidate(document, document.Content, candidates[:3], app)
        except Exception:
            return False

    def _locate_by_office_collection(self, document, app, location: ImageLocation) -> bool:
        collection_name = location.office_collection
        collection_index = location.office_collection_index
        if not collection_name or collection_index is None:
            return False
        try:
            collection = getattr(document, collection_name)
            item = collection(collection_index)
        except Exception:
            try:
                collection = getattr(document, collection_name)
                item = collection.Item(collection_index)
            except Exception:
                return False

        target_range = None
        if collection_name == "InlineShapes":
            try:
                item.Select()
            except Exception:
                pass
            try:
                target_range = item.Range
            except Exception:
                target_range = None
        else:
            try:
                item.Select()
            except Exception:
                pass
            try:
                target_range = item.Anchor
            except Exception:
                target_range = None

        if target_range is None:
            try:
                target_range = app.Selection.Range
            except Exception:
                target_range = None
        if target_range is None:
            return False

        try:
            app.Selection.SetRange(target_range.Start, target_range.End)
        except Exception:
            pass
        try:
            document.ActiveWindow.ScrollIntoView(target_range, True)
        except Exception:
            pass
        try:
            app.ActiveWindow.ScrollIntoView(target_range, True)
        except Exception:
            pass
        try:
            target_range.Select()
        except Exception:
            pass
        return True

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

    def _paragraph_window_range(self, document, paragraph_index: int, radius: int):
        try:
            paragraph_count = document.Paragraphs.Count
            start_index = max(1, paragraph_index - radius)
            end_index = min(paragraph_count, paragraph_index + radius)
            start_range = document.Paragraphs(start_index).Range
            end_range = document.Paragraphs(end_index).Range
            return document.Range(start_range.Start, end_range.End)
        except Exception:
            return None

    def _find_first_candidate(self, document, search_scope, candidates: list[str], app) -> bool:
        try:
            base_start = search_scope.Start
            base_end = search_scope.End
        except Exception:
            return False

        for candidate in candidates:
            try:
                search_range = document.Range(base_start, base_end)
                finder = search_range.Find
                found = finder.Execute(
                    FindText=candidate,
                    Forward=True,
                    Wrap=0,
                    Format=False,
                    MatchCase=False,
                    MatchWholeWord=False,
                    MatchWildcards=False,
                    MatchSoundsLike=False,
                    MatchAllWordForms=False,
                )
            except Exception:
                continue
            if not found:
                continue
            try:
                search_range.Select()
            except Exception:
                pass
            try:
                app.Selection.SetRange(search_range.Start, search_range.End)
            except Exception:
                pass
            try:
                app.ActiveWindow.ScrollIntoView(search_range, True)
            except Exception:
                pass
            return True
        return False
