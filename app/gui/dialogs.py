from __future__ import annotations

import io
import tkinter as tk
from tkinter import messagebox, ttk
from typing import TYPE_CHECKING

from PIL import Image, ImageTk

from app.models import ImageAsset

if TYPE_CHECKING:
    from .app import ImageExtractorApp


class ImageViewerDialog(tk.Toplevel):
    MIN_SCALE = 0.1
    MAX_SCALE = 8.0
    SCALE_STEP = 1.15
    INITIAL_FIT_RETRY_MS = 16
    INITIAL_FIT_MAX_RETRIES = 6
    DEFAULT_WIDTH = 980
    DEFAULT_HEIGHT = 760
    MIN_WINDOW_WIDTH = 420
    MIN_WINDOW_HEIGHT = 320
    SCREEN_MARGIN_X = 120
    SCREEN_MARGIN_Y = 120
    WINDOW_CHROME_WIDTH = 40
    WINDOW_CHROME_HEIGHT = 140

    def __init__(self, master: "ImageExtractorApp", image: ImageAsset) -> None:
        super().__init__(master)
        self.app = master
        self.current_image = image
        self.original_image: Image.Image | None = None
        self.photo_ref: ImageTk.PhotoImage | None = None
        self.scale = 1.0
        self.image_item_id: int | None = None
        self._pending_initial_fit = False
        self._pending_view_reset = False
        self._initial_fit_retry_count = 0
        self._initial_fit_canvas_size: tuple[int, int] | None = None
        self._pending_window_size: tuple[int, int] | None = None

        self.title("图片查看器")
        self.geometry(f"{self.DEFAULT_WIDTH}x{self.DEFAULT_HEIGHT}")
        self.minsize(self.MIN_WINDOW_WIDTH, self.MIN_WINDOW_HEIGHT)
        self.transient(master)

        self._build()
        self._bind_navigation_keys()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.bind("<Escape>", self._on_escape_key)
        self.set_image(image)
        self.after_idle(self._focus_viewer)

    def _build(self) -> None:
        outer = ttk.Frame(self, padding=10)
        outer.pack(fill="both", expand=True)

        self.info_var = tk.StringVar(value="")
        ttk.Label(outer, textvariable=self.info_var, justify="left").pack(anchor="w", pady=(0, 8))

        canvas_frame = ttk.Frame(outer)
        canvas_frame.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(canvas_frame, bg="#111827", highlightthickness=0)
        self.v_scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        self.h_scrollbar = ttk.Scrollbar(canvas_frame, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=self.v_scrollbar.set, xscrollcommand=self.h_scrollbar.set)

        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.v_scrollbar.grid(row=0, column=1, sticky="ns")
        self.h_scrollbar.grid(row=1, column=0, sticky="ew")
        canvas_frame.grid_rowconfigure(0, weight=1)
        canvas_frame.grid_columnconfigure(0, weight=1)

        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

    def _bind_navigation_keys(self) -> None:
        self.bind("<Left>", lambda _event: self._move_visible_image(-1))
        self.bind("<Right>", lambda _event: self._move_visible_image(1))
        self.bind("<h>", lambda _event: self._move_visible_image(-1))
        self.bind("<l>", lambda _event: self._move_visible_image(1))
        self.bind("<Return>", self._on_enter_key)
        self.canvas.bind("<Left>", lambda _event: self._move_visible_image(-1))
        self.canvas.bind("<Right>", lambda _event: self._move_visible_image(1))
        self.canvas.bind("<h>", lambda _event: self._move_visible_image(-1))
        self.canvas.bind("<l>", lambda _event: self._move_visible_image(1))
        self.canvas.bind("<Return>", self._on_enter_key)

    def _focus_viewer(self) -> None:
        self.focus_force()
        self.canvas.focus_set()

    def set_image(self, image: ImageAsset) -> None:
        self.current_image = image
        try:
            with Image.open(io.BytesIO(image.image_bytes)) as source:
                self.original_image = source.convert("RGBA")
        except Exception:
            self.original_image = None
            self.photo_ref = None
            self._pending_initial_fit = False
            self._pending_view_reset = False
            self._initial_fit_retry_count = 0
            self._initial_fit_canvas_size = None
            self._pending_window_size = None
            self.canvas.delete("all")
            self.canvas.configure(scrollregion=(0, 0, 0, 0))
            self.title("图片查看器 - 加载失败")
            self.info_var.set(f"{image.name}\n分辨率：{image.resolution_text}\n图片加载失败")
            self.after_idle(self._focus_viewer)
            return
        self._apply_initial_view()
        self.after_idle(self._focus_viewer)

    def _compute_target_window_size(self, image: Image.Image) -> tuple[int, int]:
        screen_width = max(self.MIN_WINDOW_WIDTH, self.winfo_screenwidth() - self.SCREEN_MARGIN_X)
        screen_height = max(self.MIN_WINDOW_HEIGHT, self.winfo_screenheight() - self.SCREEN_MARGIN_Y)
        target_width = min(screen_width, max(self.MIN_WINDOW_WIDTH, image.width + self.WINDOW_CHROME_WIDTH))
        target_height = min(screen_height, max(self.MIN_WINDOW_HEIGHT, image.height + self.WINDOW_CHROME_HEIGHT))
        return (target_width, target_height)

    def _compute_fit_scale(self, canvas_width: int, canvas_height: int) -> float:
        if self.original_image is None or self.original_image.width <= 0 or self.original_image.height <= 0:
            return 1.0
        width_scale = canvas_width / self.original_image.width
        height_scale = canvas_height / self.original_image.height
        fit_scale = min(width_scale, height_scale)
        return max(self.MIN_SCALE, min(self.MAX_SCALE, fit_scale))

    def _apply_initial_view(self) -> None:
        if self.original_image is None:
            return
        target_width, target_height = self._compute_target_window_size(self.original_image)
        self._pending_initial_fit = True
        self._pending_view_reset = True
        self._initial_fit_retry_count = 0
        self._initial_fit_canvas_size = None
        self._pending_window_size = (target_width, target_height)
        self.geometry(f"{target_width}x{target_height}")
        self.after_idle(self._finalize_initial_view)

    def _finalize_initial_view(self) -> None:
        if self.original_image is None:
            self._pending_initial_fit = False
            self._pending_view_reset = False
            self._initial_fit_retry_count = 0
            self._initial_fit_canvas_size = None
            self._pending_window_size = None
            return
        self.update_idletasks()
        if self._should_retry_initial_fit():
            self._initial_fit_retry_count += 1
            self.after(self.INITIAL_FIT_RETRY_MS, self._finalize_initial_view)
            return
        canvas_width = max(1, self.canvas.winfo_width())
        canvas_height = max(1, self.canvas.winfo_height())
        self.scale = self._compute_fit_scale(canvas_width, canvas_height)
        self._initial_fit_canvas_size = (canvas_width, canvas_height)
        self._redraw()
        self.after_idle(self._complete_initial_view)

    def _should_retry_initial_fit(self) -> bool:
        if self._initial_fit_retry_count >= self.INITIAL_FIT_MAX_RETRIES:
            return False
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        if canvas_width <= 1 or canvas_height <= 1:
            return True
        if self._pending_window_size is None:
            return False
        window_width = self.winfo_width()
        window_height = self.winfo_height()
        target_width, target_height = self._pending_window_size
        return abs(window_width - target_width) > 4 or abs(window_height - target_height) > 4

    def _complete_initial_view(self) -> None:
        if self.original_image is None:
            self._pending_initial_fit = False
            self._pending_view_reset = False
            self._initial_fit_retry_count = 0
            self._initial_fit_canvas_size = None
            self._pending_window_size = None
            return
        self.update_idletasks()
        if self._should_refit_after_redraw():
            self._initial_fit_retry_count += 1
            self.after(self.INITIAL_FIT_RETRY_MS, self._finalize_initial_view)
            return
        self._pending_initial_fit = False
        self._initial_fit_retry_count = 0
        self._initial_fit_canvas_size = None
        self._pending_window_size = None
        self._reset_canvas_view()

    def _should_refit_after_redraw(self) -> bool:
        if self._initial_fit_retry_count >= self.INITIAL_FIT_MAX_RETRIES:
            return False
        if self._initial_fit_canvas_size is None:
            return False
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        fit_canvas_width, fit_canvas_height = self._initial_fit_canvas_size
        return abs(canvas_width - fit_canvas_width) > 4 or abs(canvas_height - fit_canvas_height) > 4

    def _reset_canvas_view(self) -> None:
        if not self._pending_view_reset:
            return
        self.canvas.xview_moveto(0.0)
        self.canvas.yview_moveto(0.0)
        self._pending_view_reset = False

    def _on_mousewheel(self, event: tk.Event) -> None:
        if self.original_image is None:
            return
        delta = getattr(event, "delta", 0)
        if delta == 0:
            return
        next_scale = self.scale * self.SCALE_STEP if delta > 0 else self.scale / self.SCALE_STEP
        next_scale = max(self.MIN_SCALE, min(self.MAX_SCALE, next_scale))
        if abs(next_scale - self.scale) < 1e-6:
            return
        self.scale = next_scale
        self._redraw()

    def _on_canvas_configure(self, _event: tk.Event) -> None:
        if self._pending_initial_fit:
            return
        self._redraw()

    def _move_visible_image(self, step: int) -> str:
        visible_images = getattr(self.app, "_visible_images", [])
        if not visible_images:
            return "break"

        current = self.app.service.current_preview()
        if current is None:
            return "break"

        current_index = None
        for index, image in enumerate(visible_images):
            if image.id == current.id:
                current_index = index
                break
        if current_index is None:
            return "break"

        target_index = current_index + step
        if target_index < 0 or target_index >= len(visible_images):
            return "break"

        target = visible_images[target_index]
        if not self.app.set_preview_from_viewer(target.id):
            return "break"
        self.set_image(target)
        return "break"

    def _on_enter_key(self, _event: tk.Event) -> str:
        self._locate_current_image()
        return "break"

    def _locate_current_image(self) -> None:
        result = self.app.locate_image(self.current_image.id)
        if result.success:
            messagebox.showinfo("定位结果", result.message, parent=self)
        else:
            messagebox.showwarning("定位提示", result.message, parent=self)
        self.app.refresh_preview()
        self._focus_viewer()

    def _redraw(self) -> None:
        if self.original_image is None:
            return

        width = max(1, round(self.original_image.width * self.scale))
        height = max(1, round(self.original_image.height * self.scale))
        resized = self.original_image.resize((width, height), Image.Resampling.LANCZOS)
        self.photo_ref = ImageTk.PhotoImage(resized)

        self.canvas.delete("all")
        self.image_item_id = self.canvas.create_image(0, 0, anchor="nw", image=self.photo_ref)
        self.canvas.configure(scrollregion=(0, 0, width, height))

        canvas_width = max(1, self.canvas.winfo_width())
        canvas_height = max(1, self.canvas.winfo_height())
        x = max(0, (canvas_width - width) // 2)
        y = max(0, (canvas_height - height) // 2)
        self.canvas.coords(self.image_item_id, x, y)
        self.canvas.configure(scrollregion=(0, 0, max(width, canvas_width), max(height, canvas_height)))

        self.title(f"图片查看器 - {self.current_image.name}")
        self.info_var.set(
            "\n".join(
                [
                    self.current_image.name,
                    f"分辨率：{self.current_image.resolution_text}",
                    f"缩放：{round(self.scale * 100)}%",
                ]
            )
        )

    def _on_close(self) -> None:
        self.app.image_viewer = None
        self.destroy()

    def _on_escape_key(self, _event: tk.Event) -> str:
        self._on_close()
        return "break"


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
