from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


@dataclass(slots=True)
class ImageLocation:
    part_name: str
    rel_id: str | None
    anchor_type: str
    block_index: int | None
    occurrence_index: int
    text_hint: str = ""
    paragraph_text: str = ""
    context_before: str = ""
    context_after: str = ""
    office_collection: str = ""
    office_collection_index: int | None = None

    @property
    def display_text(self) -> str:
        block = f"段落#{self.block_index}" if self.block_index is not None else "未知段落"
        return f"{self.part_name} / {block} / {self.anchor_type}"


@dataclass(slots=True)
class ImageAsset:
    id: str
    name: str
    extension: str
    media_path: str
    md5: str
    width: int
    height: int
    size_bytes: int
    image_bytes: bytes
    location: ImageLocation
    selected: bool = False
    hidden: bool = False

    @property
    def duplicate_key(self) -> str:
        return self.md5

    @property
    def size_kb(self) -> int:
        return max(1, round(self.size_bytes / 1024))

    @property
    def resolution_text(self) -> str:
        return f"{self.width}x{self.height}"


@dataclass(slots=True)
class DocumentSession:
    document_path: Path
    images: list[ImageAsset]
    adapter_name: str
    duplicate_groups: dict[str, list[str]] = field(default_factory=dict)
    preview_image_id: str | None = None

    @property
    def total_images(self) -> int:
        return len(self.images)

    @property
    def duplicate_group_count(self) -> int:
        return len(self.duplicate_groups)

    @property
    def selected_count(self) -> int:
        return sum(1 for image in self.images if image.selected)

    @property
    def hidden_count(self) -> int:
        return sum(1 for image in self.images if image.hidden)

    def get_image(self, image_id: str) -> ImageAsset | None:
        for image in self.images:
            if image.id == image_id:
                return image
        return None
