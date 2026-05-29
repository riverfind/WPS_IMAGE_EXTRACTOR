from __future__ import annotations

import io
from collections import defaultdict
from pathlib import Path

from PIL import Image, ImageOps

from app.docx_adapter import DocxAdapter, DocumentError, NavigationResult, UnsupportedDocumentError
from app.models import DocumentSession, ImageAsset


class SessionService:
    def __init__(self) -> None:
        self.session: DocumentSession | None = None
        self.adapter: DocxAdapter | None = None
        self.thumbnail_cache: dict[tuple[str, tuple[int, int]], Image.Image] = {}

    def open_document(self, document_path: Path) -> DocumentSession:
        document_path = document_path.resolve()
        if not DocxAdapter.can_open(document_path):
            raise UnsupportedDocumentError("当前版本先支持 .docx/.docm/.dotx，后续再补强原生 .wps。")

        self.adapter = DocxAdapter(document_path)
        images = self.adapter.extract_images()
        duplicate_groups = self._build_duplicate_groups(images)
        preview_id = images[0].id if images else None
        self.session = DocumentSession(
            document_path=document_path,
            images=images,
            adapter_name=self.adapter.name,
            duplicate_groups=duplicate_groups,
            preview_image_id=preview_id,
        )
        self.thumbnail_cache.clear()
        return self.session

    def require_session(self) -> DocumentSession:
        if self.session is None:
            raise DocumentError("请先打开一个文档。")
        return self.session

    def visible_images(self) -> list[ImageAsset]:
        session = self.require_session()
        normalized_query = session.query_text.strip().lower()
        return [
            image
            for image in session.images
            if not image.hidden
            and self._matches_query(image, normalized_query)
            and self._matches_size_filter(
                image,
                session.size_filter_mode,
                session.size_filter_width,
                session.size_filter_height,
            )
        ]

    def set_text_query(self, query: str) -> None:
        session = self.require_session()
        session.query_text = query

    def set_size_filter(self, mode: str, width: str, height: str) -> None:
        session = self.require_session()
        if mode not in {"exact", "min", "max"}:
            raise DocumentError("尺寸过滤模式无效。")
        session.size_filter_mode = mode
        session.size_filter_width = self._normalize_size_token(width)
        session.size_filter_height = self._normalize_size_token(height)

    def clear_size_filter(self) -> None:
        session = self.require_session()
        session.size_filter_mode = "exact"
        session.size_filter_width = "*"
        session.size_filter_height = "*"

    def set_preview(self, image_id: str) -> ImageAsset:
        session = self.require_session()
        image = session.get_image(image_id)
        if image is None:
            raise DocumentError("未找到要预览的图片。")
        session.preview_image_id = image_id
        return image

    def current_preview(self) -> ImageAsset | None:
        session = self.require_session()
        if not session.preview_image_id:
            return None
        return session.get_image(session.preview_image_id)

    def toggle_selection(self, image_id: str, selected: bool) -> None:
        session = self.require_session()
        image = session.get_image(image_id)
        if image is None:
            raise DocumentError("未找到要勾选的图片。")
        image.selected = selected

    def select_all_visible(self) -> None:
        for image in self.visible_images():
            image.selected = True

    def invert_visible(self) -> None:
        for image in self.visible_images():
            image.selected = not image.selected

    def clear_selection(self) -> None:
        session = self.require_session()
        for image in session.images:
            image.selected = False

    def select_same_md5_as_preview(self) -> int:
        session = self.require_session()
        preview = self.current_preview()
        if preview is None:
            raise DocumentError("请先预览一张图片，再执行按 MD5 选中。")
        return self.select_same_md5(preview.md5)

    def select_same_md5(self, md5: str) -> int:
        session = self.require_session()
        current = session.duplicate_groups.get(md5, [])
        ids = current or [image.id for image in session.images if image.md5 == md5]
        for image_id in ids:
            image = session.get_image(image_id)
            if image is not None:
                image.selected = True
        return len(ids)

    def filter_selected(self) -> int:
        session = self.require_session()
        hidden_count = 0
        for image in session.images:
            if image.selected and not image.hidden:
                image.hidden = True
                image.selected = False
                hidden_count += 1
        return hidden_count

    def clear_filters(self) -> None:
        session = self.require_session()
        for image in session.images:
            image.hidden = False

    def filtered_images(self) -> list[ImageAsset]:
        session = self.require_session()
        return [image for image in session.images if image.hidden]

    def selected_images(self) -> list[ImageAsset]:
        session = self.require_session()
        return [image for image in session.images if image.selected and not image.hidden]

    def export_selected(self, output_dir: Path) -> tuple[int, int]:
        output_dir.mkdir(parents=True, exist_ok=True)
        success = 0
        failed = 0

        for image in self.selected_images():
            target = self._unique_output_path(output_dir, image.name)
            try:
                target.write_bytes(image.image_bytes)
            except OSError:
                failed += 1
            else:
                success += 1

        return success, failed

    def export_filtered_md5s(self, output_path: Path) -> int:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        md5s = list(dict.fromkeys(image.md5 for image in self.filtered_images()))
        output_path.write_text("\n".join(md5s), encoding="utf-8")
        return len(md5s)

    def import_filtered_md5s(self, input_path: Path) -> tuple[int, int]:
        try:
            content = input_path.read_text(encoding="utf-8")
        except OSError as exc:
            raise DocumentError(f"读取 MD5 文件失败：{exc}") from exc

        md5s = {line.strip().lower() for line in content.splitlines() if line.strip()}
        if not md5s:
            raise DocumentError("导入文件中没有有效的 MD5。")

        session = self.require_session()
        hidden_count = 0
        matched_md5s: set[str] = set()
        for image in session.images:
            image_md5 = image.md5.lower()
            if image_md5 not in md5s:
                continue
            matched_md5s.add(image_md5)
            if not image.hidden:
                image.hidden = True
                image.selected = False
                hidden_count += 1
        return hidden_count, len(matched_md5s)

    def delete_selected(self) -> tuple[Path, int]:
        if self.adapter is None:
            raise DocumentError("当前没有活动文档。")

        selected = self.selected_images()
        if not selected:
            raise DocumentError("请先勾选要删除的图片。")

        backup_path, deleted_count = self.adapter.delete_images(selected)
        current_path = self.require_session().document_path
        self.open_document(current_path)
        return backup_path, deleted_count

    def locate_image(self, image_id: str) -> NavigationResult:
        if self.adapter is None:
            raise DocumentError("当前没有活动文档。")
        image = self.set_preview(image_id)
        return self.adapter.locate(image)

    def build_thumbnail(self, image: ImageAsset, max_size: tuple[int, int]) -> Image.Image:
        cache_key = (image.id, max_size)
        cached = self.thumbnail_cache.get(cache_key)
        if cached is not None:
            return cached.copy()
        canvas = Image.new("RGBA", max_size, (255, 255, 255, 0))
        try:
            with Image.open(io.BytesIO(image.image_bytes)) as source:
                preview = source.convert("RGBA")
            if preview.width <= 0 or preview.height <= 0:
                self.thumbnail_cache[cache_key] = canvas
                return canvas.copy()
            contained = ImageOps.contain(preview, max_size)
            offset = ((max_size[0] - contained.width) // 2, (max_size[1] - contained.height) // 2)
            canvas.paste(contained, offset, contained)
        except Exception:
            self.thumbnail_cache[cache_key] = canvas
            return canvas.copy()
        self.thumbnail_cache[cache_key] = canvas
        return canvas.copy()

    def _build_duplicate_groups(self, images: list[ImageAsset]) -> dict[str, list[str]]:
        groups: dict[str, list[str]] = defaultdict(list)
        for image in images:
            groups[image.md5].append(image.id)
        return {md5: image_ids for md5, image_ids in groups.items() if len(image_ids) > 1}

    def _matches_query(self, image: ImageAsset, normalized_query: str) -> bool:
        if not normalized_query:
            return True
        return (
            normalized_query in image.name.lower()
            or normalized_query in image.md5.lower()
            or normalized_query in image.location.part_name.lower()
        )

    def _normalize_size_token(self, value: str) -> str:
        normalized = value.strip()
        if not normalized or normalized == "*":
            return "*"
        if not normalized.isdigit():
            raise DocumentError("宽和高只能填写正整数或 *。")
        parsed = int(normalized)
        if parsed <= 0:
            raise DocumentError("宽和高只能填写正整数或 *。")
        return str(parsed)

    def _matches_size_filter(self, image: ImageAsset, mode: str, width: str, height: str) -> bool:
        if width != "*" and not self._compare_dimension(image.width, int(width), mode):
            return False
        if height != "*" and not self._compare_dimension(image.height, int(height), mode):
            return False
        return True

    def _compare_dimension(self, actual: int, expected: int, mode: str) -> bool:
        if mode == "exact":
            return actual == expected
        if mode == "min":
            return actual >= expected
        if mode == "max":
            return actual <= expected
        raise DocumentError("尺寸过滤模式无效。")

    def _unique_output_path(self, output_dir: Path, file_name: str) -> Path:
        base = Path(file_name).stem
        suffix = Path(file_name).suffix
        candidate = output_dir / file_name
        index = 1
        while candidate.exists():
            candidate = output_dir / f"{base}_{index}{suffix}"
            index += 1
        return candidate
