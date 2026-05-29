from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk
from typing import TYPE_CHECKING

from PIL import ImageTk

from app.models import ImageAsset

if TYPE_CHECKING:
    from .app import ImageExtractorApp


class DeleteConfirmDialog(tk.Toplevel):
    def __init__(self, master: "ImageExtractorApp", images: list[ImageAsset]) -> None:
        super().__init__(master)
        self.app = master
        self.images = images
        self.result = False
        self.preview_id = images[0].id if images else None
        self.thumbnail_refs: dict[str, ImageTk.PhotoImage] = {}
        self.preview_ref: ImageTk.PhotoImage | None = None

        self.title("确认删除图片")
        self.geometry("920x620")
        self.minsize(820, 560)
        self.transient(master)
        self.grab_set()

        self._build()
        self._refresh()

    def _build(self) -> None:
        outer = ttk.Frame(self, padding=12)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text=f"将从文档中删除以下 {len(self.images)} 张图片（不可撤销）", font=("Microsoft YaHei UI", 12, "bold")).pack(
            anchor="w", pady=(0, 10)
        )

        body = ttk.Frame(outer)
        body.pack(fill="both", expand=True)

        left = ttk.Frame(body)
        left.pack(side="left", fill="both", expand=True)

        canvas = tk.Canvas(left, highlightthickness=0, bg="#f6f6f6")
        scrollbar = ttk.Scrollbar(left, orient="vertical", command=canvas.yview)
        self.cards_frame = ttk.Frame(canvas)
        self.cards_frame.bind(
            "<Configure>",
            lambda event: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas.create_window((0, 0), window=self.cards_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        right = ttk.LabelFrame(body, text="预览", padding=10)
        right.pack(side="left", fill="y", padx=(12, 0))

        self.preview_label = ttk.Label(right, text="单击缩略图预览\n双击缩略图定位", anchor="center")
        self.preview_label.pack(fill="both", expand=True)

        self.preview_info_var = tk.StringVar(value="")
        ttk.Label(right, textvariable=self.preview_info_var, justify="left", wraplength=220).pack(fill="x", pady=(10, 0))

        actions = ttk.Frame(outer)
        actions.pack(fill="x", pady=(12, 0))

        ttk.Label(actions, text="删除范围仅按当前勾选图片执行；若要扩展到相同图片，请先在主界面点击卡片上的“重复 N”。").pack(side="left")
        ttk.Button(actions, text="取消", command=self.destroy).pack(side="right")
        ttk.Button(actions, text="删除", command=self._confirm).pack(side="right", padx=(0, 8))

    def _refresh(self) -> None:
        for child in self.cards_frame.winfo_children():
            child.destroy()

        for index, image in enumerate(self.images):
            card = tk.Frame(
                self.cards_frame,
                bg="#eaf4ff" if image.id == self.preview_id else "#ffffff",
                bd=1,
                relief="solid",
                padx=6,
                pady=6,
            )
            row = index // 3
            column = index % 3
            card.grid(row=row, column=column, padx=6, pady=6, sticky="nsew")

            thumbnail = self.app.build_photo(image, (150, 110), self.thumbnail_refs)
            image_label = tk.Label(card, image=thumbnail, bg=card["bg"], cursor="hand2")
            image_label.pack()
            image_label.bind("<Button-1>", lambda _event, image_id=image.id: self._set_preview(image_id))
            image_label.bind("<Double-Button-1>", lambda _event, image_id=image.id: self._locate(image_id))

            info = tk.Label(
                card,
                text=f"{image.name}\n{image.resolution_text}\n{image.size_kb} KB",
                justify="left",
                bg=card["bg"],
                anchor="w",
            )
            info.pack(fill="x", pady=(4, 0))
            info.bind("<Button-1>", lambda _event, image_id=image.id: self._set_preview(image_id))
            info.bind("<Double-Button-1>", lambda _event, image_id=image.id: self._locate(image_id))

        self._refresh_preview()

    def _refresh_preview(self) -> None:
        current = next((image for image in self.images if image.id == self.preview_id), None)
        if current is None:
            self.preview_label.configure(text="没有可预览图片", image="")
            self.preview_info_var.set("")
            return

        preview = self.app.build_photo(current, (260, 260), {"dialog-preview": None})
        self.preview_ref = preview
        self.preview_label.configure(image=preview, text="")
        self.preview_info_var.set(f"{current.name}\n{current.resolution_text}\nMD5: {current.md5}\n{current.location.display_text}")

    def _set_preview(self, image_id: str) -> None:
        self.preview_id = image_id
        self._refresh()

    def _locate(self, image_id: str) -> None:
        result = self.app.locate_image(image_id)
        if result.success:
            messagebox.showinfo("定位结果", result.message, parent=self)
        else:
            messagebox.showwarning("定位提示", result.message, parent=self)

    def _confirm(self) -> None:
        self.result = True
        self.destroy()


class FilteredImagesDialog(tk.Toplevel):
    def __init__(self, master: "ImageExtractorApp", images: list[ImageAsset]) -> None:
        super().__init__(master)
        self.app = master
        self.images = images
        self.preview_id = images[0].id if images else None
        self.thumbnail_refs: dict[str, ImageTk.PhotoImage] = {}
        self.preview_ref: ImageTk.PhotoImage | None = None

        self.title("已过滤图片")
        self.geometry("980x660")
        self.minsize(860, 580)
        self.transient(master)

        self._build()
        self._refresh()

    def _build(self) -> None:
        outer = ttk.Frame(self, padding=12)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text=f"当前共有 {len(self.images)} 张已过滤图片", font=("Microsoft YaHei UI", 12, "bold")).pack(
            anchor="w", pady=(0, 10)
        )

        body = ttk.Frame(outer)
        body.pack(fill="both", expand=True)

        left = ttk.Frame(body)
        left.pack(side="left", fill="both", expand=True)

        canvas = tk.Canvas(left, highlightthickness=0, bg="#f6f6f6")
        scrollbar = ttk.Scrollbar(left, orient="vertical", command=canvas.yview)
        self.cards_frame = ttk.Frame(canvas)
        self.cards_frame.bind(
            "<Configure>",
            lambda event: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas.create_window((0, 0), window=self.cards_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        right = ttk.LabelFrame(body, text="预览", padding=10)
        right.pack(side="left", fill="y", padx=(12, 0))

        self.preview_label = ttk.Label(right, text="单击缩略图预览\n双击缩略图定位", anchor="center")
        self.preview_label.pack(fill="both", expand=True)

        self.preview_info_var = tk.StringVar(value="")
        ttk.Label(right, textvariable=self.preview_info_var, justify="left", wraplength=240).pack(fill="x", pady=(10, 0))

        actions = ttk.Frame(outer)
        actions.pack(fill="x", pady=(12, 0))
        ttk.Label(actions, text="此窗口仅查看已过滤图片；若需恢复显示，请回到主窗口点击“恢复过滤”。").pack(side="left")
        ttk.Button(actions, text="关闭", command=self.destroy).pack(side="right")

    def _refresh(self) -> None:
        for child in self.cards_frame.winfo_children():
            child.destroy()

        for index, image in enumerate(self.images):
            card = tk.Frame(
                self.cards_frame,
                bg="#eaf4ff" if image.id == self.preview_id else "#ffffff",
                bd=1,
                relief="solid",
                padx=6,
                pady=6,
            )
            row = index // 3
            column = index % 3
            card.grid(row=row, column=column, padx=6, pady=6, sticky="nsew")

            thumbnail = self.app.build_photo(image, (150, 110), self.thumbnail_refs)
            image_label = tk.Label(card, image=thumbnail, bg=card["bg"], cursor="hand2")
            image_label.pack()
            image_label.bind("<Button-1>", lambda _event, image_id=image.id: self._set_preview(image_id))
            image_label.bind("<Double-Button-1>", lambda _event, image_id=image.id: self._locate(image_id))

            info = tk.Label(
                card,
                text=f"{image.name}\n{image.resolution_text}\n{image.size_kb} KB\nMD5: {image.md5[:12]}...",
                justify="left",
                bg=card["bg"],
                anchor="w",
            )
            info.pack(fill="x", pady=(4, 0))
            info.bind("<Button-1>", lambda _event, image_id=image.id: self._set_preview(image_id))
            info.bind("<Double-Button-1>", lambda _event, image_id=image.id: self._locate(image_id))

        self._refresh_preview()

    def _refresh_preview(self) -> None:
        current = next((image for image in self.images if image.id == self.preview_id), None)
        if current is None:
            self.preview_label.configure(text="没有可预览图片", image="")
            self.preview_info_var.set("")
            return

        preview = self.app.build_photo(current, (300, 300), {"filtered-dialog-preview": None})
        self.preview_ref = preview
        self.preview_label.configure(image=preview, text="")
        self.preview_info_var.set(
            "\n".join(
                [
                    current.name,
                    f"分辨率：{current.resolution_text}",
                    f"大小：{current.size_kb} KB",
                    f"MD5：{current.md5}",
                    f"位置：{current.location.display_text}",
                ]
            )
        )

    def _set_preview(self, image_id: str) -> None:
        self.preview_id = image_id
        self._refresh()

    def _locate(self, image_id: str) -> None:
        result = self.app.locate_image(image_id)
        if result.success:
            messagebox.showinfo("定位结果", result.message, parent=self)
        else:
            messagebox.showwarning("定位提示", result.message, parent=self)
