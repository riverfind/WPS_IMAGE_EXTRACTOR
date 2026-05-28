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

    def visible_images(self, query: str = "") -> list[ImageAsset]:
        session = self.require_session()
        normalized = query.strip().lower()
        images = [image for image in session.images if not image.hidden]
        if not normalized:
            return images
        return [
            image
            for image in images
            if normalized in image.name.lower()
            or normalized in image.md5.lower()
            or normalized in image.location.part_name.lower()
        ]

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

    def select_all_visible(self, query: str = "") -> None:
        for image in self.visible_images(query):
            image.selected = True

    def invert_visible(self, query: str = "") -> None:
        for image in self.visible_images(query):
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
                hidden_count += 1
        return hidden_count

    def clear_filters(self) -> None:
        session = self.require_session()
        for image in session.images:
            image.hidden = False

    def selected_images(self) -> list[ImageAsset]:
        session = self.require_session()
        return [image for image in session.images if image.selected]

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

    def _unique_output_path(self, output_dir: Path, file_name: str) -> Path:
        base = Path(file_name).stem
        suffix = Path(file_name).suffix
        candidate = output_dir / file_name
        index = 1
        while candidate.exists():
            candidate = output_dir / f"{base}_{index}{suffix}"
            index += 1
        return candidate
