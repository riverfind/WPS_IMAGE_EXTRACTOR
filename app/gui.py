from __future__ import annotations

import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from PIL import Image, ImageTk

from app.docx_adapter import DocumentError, NavigationResult, UnsupportedDocumentError
from app.models import ImageAsset
from app.services import SessionService


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


class ImageExtractorApp(tk.Tk):
    GRID_COLUMNS = 4
    MIN_GRID_COLUMNS = 1
    CARD_WIDTH = 210
    CARD_HEIGHT = 236
    CARD_PAD_X = 10
    CARD_PAD_Y = 10
    OVERSCAN_ROWS = 2

    def __init__(self) -> None:
        super().__init__()
        self.title("WPS Image Extractor (Tkinter)")
        self.geometry("1380x860")
        self.minsize(420, 720)

        self.service = SessionService()
        self.query_var = tk.StringVar()
        self.doc_name_var = tk.StringVar(value="未打开文档")
        self.total_var = tk.StringVar(value="总图片数：0")
        self.duplicate_var = tk.StringVar(value="重复组：0")
        self.selection_var = tk.StringVar(value="已选：0")
        self.filtered_var = tk.StringVar(value="已过滤：0")
        self.status_var = tk.StringVar(value="请选择一个 .docx 文档开始。")
        self.preview_info_var = tk.StringVar(value="单击缩略图后在此查看大图预览。")

        self.thumbnail_refs: dict[str, ImageTk.PhotoImage] = {}
        self.preview_refs: dict[str, ImageTk.PhotoImage] = {}
        self.preview_ref: ImageTk.PhotoImage | None = None
        self._visible_images: list[ImageAsset] = []
        self._slot_pool: list[dict[str, object]] = []
        self._viewport_after_id: str | None = None
        self._empty_label_id: int | None = None
        self._last_render_range: tuple[int, int] | None = None
        self._thumbnail_result_queue: queue.Queue[tuple[int, str, Image.Image]] = queue.Queue()
        self._thumbnail_pending: set[str] = set()
        self._thumbnail_generation = 0
        self._grid_columns = self.GRID_COLUMNS
        self._placeholder_thumbnail = ImageTk.PhotoImage(Image.new("RGBA", (170, 120), "#e5e7eb"))

        self._build_layout()
        self._bind_keyboard_navigation()
        self.query_var.trace_add("write", lambda *_args: self.refresh_grid())
        self.after(50, self._poll_thumbnail_results)

    def _build_layout(self) -> None:
        root = ttk.Frame(self, padding=10)
        root.pack(fill="both", expand=True)

        toolbar = ttk.Frame(root)
        toolbar.pack(fill="x", pady=(0, 8))

        ttk.Button(toolbar, text="打开 WPS 文档...", command=self.open_document).pack(side="left")
        ttk.Button(toolbar, text="重新加载", command=self.reload_document).pack(side="left", padx=(8, 0))
        ttk.Button(toolbar, text="全选", command=self.select_all).pack(side="left", padx=(8, 0))
        ttk.Button(toolbar, text="反选", command=self.invert_selection).pack(side="left", padx=(8, 0))
        ttk.Button(toolbar, text="删除所选...", command=self.delete_selected).pack(side="left", padx=(8, 0))
        ttk.Button(toolbar, text="过滤所选", command=self.filter_selected).pack(side="left", padx=(8, 0))
        ttk.Button(toolbar, text="恢复过滤", command=self.clear_filters).pack(side="left", padx=(8, 0))
        ttk.Button(toolbar, text="导出所选...", command=self.export_selected).pack(side="left", padx=(8, 0))

        body = ttk.Frame(root)
        body.pack(fill="both", expand=True)

        left = ttk.LabelFrame(body, text="文档信息", padding=10)
        left.pack(side="left", fill="y")

        ttk.Label(left, textvariable=self.doc_name_var, wraplength=220).pack(anchor="w")
        ttk.Label(left, textvariable=self.total_var).pack(anchor="w", pady=(8, 0))
        ttk.Label(left, textvariable=self.duplicate_var).pack(anchor="w", pady=(4, 0))
        ttk.Label(left, textvariable=self.selection_var).pack(anchor="w", pady=(4, 0))
        ttk.Label(left, textvariable=self.filtered_var).pack(anchor="w", pady=(4, 0))

        ttk.Separator(left).pack(fill="x", pady=10)
        ttk.Label(left, text="搜索文件名 / MD5 / 部件").pack(anchor="w")
        self.search_entry = ttk.Entry(left, textvariable=self.query_var, width=26)
        self.search_entry.pack(fill="x", pady=(6, 0))

        ttk.Label(
            left,
            text="交互说明：\n- 单击缩略图：仅预览\n- 复选框：勾选图片\n- 双击缩略图：定位文档",
            justify="left",
        ).pack(anchor="w", pady=(12, 0))

        center = ttk.LabelFrame(body, text="缩略图区", padding=0)
        center.pack(side="left", fill="both", expand=True, padx=(10, 10))

        canvas_area = ttk.Frame(center)
        canvas_area.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(canvas_area, bg="#f3f4f6", highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(canvas_area, orient="vertical", command=self._on_scrollbar)
        self.h_scrollbar = ttk.Scrollbar(canvas_area, orient="horizontal", command=self._on_horizontal_scroll)
        self.canvas.configure(yscrollcommand=self.scrollbar.set, xscrollcommand=self.h_scrollbar.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar.grid(row=0, column=1, sticky="ns")
        self.h_scrollbar.grid(row=1, column=0, sticky="ew")
        canvas_area.grid_rowconfigure(0, weight=1)
        canvas_area.grid_columnconfigure(0, weight=1)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<Shift-MouseWheel>", self._on_shift_mousewheel)

        right = ttk.LabelFrame(body, text="预览", padding=10)
        right.pack(side="left", fill="y")

        self.preview_label = ttk.Label(right, text="单击缩略图后显示大图预览", anchor="center")
        self.preview_label.pack(fill="both", expand=True)
        ttk.Label(right, textvariable=self.preview_info_var, justify="left", wraplength=260).pack(fill="x", pady=(10, 0))

        status = ttk.Frame(root)
        status.pack(fill="x", pady=(8, 0))
        ttk.Label(status, textvariable=self.status_var).pack(side="left")

    def build_photo(
        self,
        image: ImageAsset,
        size: tuple[int, int],
        target_cache: dict[str, ImageTk.PhotoImage | None],
    ) -> ImageTk.PhotoImage:
        cache_key = f"{image.id}:{size[0]}x{size[1]}"
        cached = target_cache.get(cache_key)
        if cached is not None:
            return cached
        pil_image = self.service.build_thumbnail(image, size)
        photo = ImageTk.PhotoImage(pil_image)
        target_cache[cache_key] = photo
        return photo

    def _bind_keyboard_navigation(self) -> None:
        self.bind_all("<Up>", lambda event: self._on_grid_key(event, "up"))
        self.bind_all("<Down>", lambda event: self._on_grid_key(event, "down"))
        self.bind_all("<Left>", lambda event: self._on_grid_key(event, "left"))
        self.bind_all("<Right>", lambda event: self._on_grid_key(event, "right"))
        self.bind_all("<Prior>", lambda event: self._on_grid_key(event, "page_up"))
        self.bind_all("<Next>", lambda event: self._on_grid_key(event, "page_down"))

    def open_document(self) -> None:
        path = filedialog.askopenfilename(
            title="选择 WPS 文档",
            filetypes=[
                ("Word 文档", "*.docx *.docm *.dotx *.wps"),
                ("所有文件", "*.*"),
            ],
        )
        if not path:
            return

        try:
            session = self.service.open_document(Path(path))
        except UnsupportedDocumentError as exc:
            messagebox.showwarning("格式限制", str(exc), parent=self)
            self.status_var.set(str(exc))
            return
        except (DocumentError, OSError) as exc:
            messagebox.showerror("打开失败", str(exc), parent=self)
            self.status_var.set(f"打开失败：{exc}")
            return

        self._reset_document_view_state()
        self.status_var.set(f"已打开 {session.document_path.name}，共提取 {session.total_images} 张图片。")
        self.refresh_all()

    def reload_document(self) -> None:
        try:
            current_path = self.service.require_session().document_path
        except DocumentError:
            messagebox.showwarning("提示", "请先打开一个文档。", parent=self)
            return

        try:
            session = self.service.open_document(current_path)
        except (DocumentError, OSError) as exc:
            messagebox.showerror("重新加载失败", str(exc), parent=self)
            self.status_var.set(f"重新加载失败：{exc}")
            return

        self._reset_document_view_state()
        self.status_var.set(f"已重新加载 {session.document_path.name}，共提取 {session.total_images} 张图片。")
        self.refresh_all()

    def refresh_all(self) -> None:
        try:
            session = self.service.require_session()
        except DocumentError:
            return

        self.doc_name_var.set(session.document_path.name)
        self.total_var.set(f"总图片数：{session.total_images}")
        self.duplicate_var.set(f"重复组：{session.duplicate_group_count}")
        self.selection_var.set(f"已选：{session.selected_count}")
        self.filtered_var.set(f"已过滤：{session.hidden_count}")

        preview = self.service.current_preview()
        visible = self.service.visible_images(self.query_var.get())
        visible_ids = {image.id for image in visible}
        if visible and (preview is None or preview.hidden or preview.id not in visible_ids):
            self.service.set_preview(visible[0].id)
        self.refresh_grid()
        self.refresh_preview()

    def refresh_grid(self) -> None:
        try:
            self._visible_images = self.service.visible_images(self.query_var.get())
            session = self.service.require_session()
        except DocumentError:
            return

        self._thumbnail_generation += 1
        self._thumbnail_pending.clear()
        self._last_render_range = None
        self.selection_var.set(f"已选：{session.selected_count}")
        self.filtered_var.set(f"已过滤：{session.hidden_count}")

        if self._empty_label_id is not None:
            self.canvas.delete(self._empty_label_id)
            self._empty_label_id = None

        if not self._visible_images:
            self._hide_all_slots()
            self.canvas.configure(scrollregion=(0, 0, 0, 0))
            self._empty_label_id = self.canvas.create_text(40, 40, anchor="nw", text="当前没有可显示的图片。")
            return

        self._update_grid_layout_metrics(reset_x=True)
        self._ensure_slot_pool()
        self._schedule_viewport_render(force=True)

    def refresh_preview(self) -> None:
        try:
            current = self.service.current_preview()
        except DocumentError:
            return

        if current is None:
            self.preview_label.configure(text="没有可预览图片", image="")
            self.preview_info_var.set("请先打开文档。")
            return

        preview = self.build_photo(current, (320, 320), self.preview_refs)
        self.preview_ref = preview
        self.preview_label.configure(image=preview, text="")

        duplicate_ids = self.service.require_session().duplicate_groups.get(current.md5, [])
        duplicate_text = f"{len(duplicate_ids)} 张" if duplicate_ids else "无重复"
        self.preview_info_var.set(
            "\n".join(
                [
                    current.name,
                    f"对象序号：#{current.location.occurrence_index + 1}",
                    f"分辨率：{current.resolution_text}",
                    f"大小：{current.size_kb} KB",
                    f"MD5：{current.md5}",
                    f"重复组：{duplicate_text}",
                    f"位置：{current.location.display_text}",
                ]
            )
        )

    def set_preview(self, image_id: str) -> None:
        try:
            self.service.set_preview(image_id)
        except DocumentError as exc:
            messagebox.showerror("预览失败", str(exc), parent=self)
            return
        self._rerender_visible_slots(force=True)
        self.refresh_preview()

    def on_check(self, image_id: str, selected: bool) -> None:
        try:
            self.service.toggle_selection(image_id, selected)
            session = self.service.require_session()
        except DocumentError as exc:
            messagebox.showerror("勾选失败", str(exc), parent=self)
            return
        self.selection_var.set(f"已选：{session.selected_count}")

    def select_all(self) -> None:
        try:
            self.service.select_all_visible(self.query_var.get())
        except DocumentError as exc:
            messagebox.showwarning("提示", str(exc), parent=self)
            return
        self.status_var.set("已选中当前可见图片。")
        self._rerender_visible_slots(force=True)
        self.refresh_all_stats()

    def invert_selection(self) -> None:
        try:
            self.service.invert_visible(self.query_var.get())
        except DocumentError as exc:
            messagebox.showwarning("提示", str(exc), parent=self)
            return
        self.status_var.set("已对当前可见图片执行反选。")
        self._rerender_visible_slots(force=True)
        self.refresh_all_stats()

    def on_select_duplicate_group(self, md5: str) -> None:
        try:
            count = self.service.select_same_md5(md5)
            session = self.service.require_session()
        except DocumentError as exc:
            messagebox.showwarning("提示", str(exc), parent=self)
            return
        self.selection_var.set(f"已选：{session.selected_count}")
        self.status_var.set(f"已通过重复组入口选中 {count} 张相同图片。")
        self._rerender_visible_slots(force=True)

    def filter_selected(self) -> None:
        try:
            count = self.service.filter_selected()
        except DocumentError as exc:
            messagebox.showwarning("提示", str(exc), parent=self)
            return
        self.status_var.set(f"已隐藏 {count} 张选中图片。")
        self.refresh_all()

    def clear_filters(self) -> None:
        try:
            self.service.clear_filters()
        except DocumentError as exc:
            messagebox.showwarning("提示", str(exc), parent=self)
            return
        self.status_var.set("已恢复全部过滤项。")
        self.refresh_all()

    def export_selected(self) -> None:
        try:
            selected = self.service.selected_images()
        except DocumentError as exc:
            messagebox.showwarning("提示", str(exc), parent=self)
            return
        if not selected:
            messagebox.showwarning("提示", "请先勾选要导出的图片。", parent=self)
            return

        output_dir = filedialog.askdirectory(title="选择导出目录")
        if not output_dir:
            return

        success, failed = self.service.export_selected(Path(output_dir))
        self.status_var.set(f"导出完成：成功 {success} 张，失败 {failed} 张。")
        messagebox.showinfo("导出完成", f"成功：{success} 张\n失败：{failed} 张", parent=self)

    def delete_selected(self) -> None:
        try:
            selected = self.service.selected_images()
        except DocumentError as exc:
            messagebox.showwarning("提示", str(exc), parent=self)
            return
        if not selected:
            messagebox.showwarning("提示", "请先勾选要删除的图片。", parent=self)
            return

        dialog = DeleteConfirmDialog(self, selected)
        self.wait_window(dialog)
        if not dialog.result:
            return

        try:
            backup_path, deleted_count = self.service.delete_selected()
        except DocumentError as exc:
            messagebox.showerror("删除失败", str(exc), parent=self)
            self.status_var.set(f"删除失败：{exc}")
            return

        self._reset_document_view_state()
        self.status_var.set(f"已删除 {deleted_count} 张图片，原文件备份为：{backup_path.name}")
        messagebox.showinfo(
            "删除完成",
            f"已删除 {deleted_count} 张图片。\n已生成备份：{backup_path}",
            parent=self,
        )
        self.refresh_all()

    def locate_image(self, image_id: str) -> NavigationResult:
        try:
            return self.service.locate_image(image_id)
        except DocumentError as exc:
            return NavigationResult(False, str(exc))

    def on_locate(self, image_id: str) -> None:
        result = self.locate_image(image_id)
        if result.success:
            messagebox.showinfo("定位结果", result.message, parent=self)
        else:
            messagebox.showwarning("定位提示", result.message, parent=self)
        self.refresh_preview()

    def refresh_all_stats(self) -> None:
        try:
            session = self.service.require_session()
        except DocumentError:
            return
        self.selection_var.set(f"已选：{session.selected_count}")
        self.filtered_var.set(f"已过滤：{session.hidden_count}")

    def _reset_document_view_state(self) -> None:
        self.thumbnail_refs.clear()
        self.preview_refs.clear()
        self.preview_ref = None
        self._thumbnail_generation += 1
        self._thumbnail_pending.clear()
        self._last_render_range = None
        self._destroy_slot_pool()
        self.preview_label.configure(image="", text="单击缩略图后显示大图预览")

    def _destroy_slot_pool(self) -> None:
        for slot in self._slot_pool:
            try:
                self.canvas.delete(slot["window_id"])
            except Exception:
                pass
            try:
                slot["frame"].destroy()
            except Exception:
                pass
        self._slot_pool.clear()

    def _on_grid_key(self, event: tk.Event, direction: str) -> str | None:
        if self._should_ignore_grid_key_event():
            return None
        if not self._visible_images:
            return "break"

        current_index = self._current_visible_index()
        if current_index is None:
            current_index = 0

        if direction == "left":
            target_index = max(0, current_index - 1)
        elif direction == "right":
            target_index = min(len(self._visible_images) - 1, current_index + 1)
        elif direction == "up":
            target_index = max(0, current_index - self._grid_columns)
        elif direction == "down":
            target_index = min(len(self._visible_images) - 1, current_index + self._grid_columns)
        elif direction == "page_up":
            target_index = max(0, current_index - self._page_step())
        elif direction == "page_down":
            target_index = min(len(self._visible_images) - 1, current_index + self._page_step())
        else:
            return None

        if target_index != current_index:
            self._set_preview_by_index(target_index)
        return "break"

    def _should_ignore_grid_key_event(self) -> bool:
        focused = self.focus_get()
        if focused is None:
            return False
        if focused == self.search_entry:
            return True
        widget_class = focused.winfo_class()
        return widget_class in {"Entry", "TEntry", "Text"}

    def _current_visible_index(self) -> int | None:
        try:
            current = self.service.current_preview()
        except DocumentError:
            return None
        if current is None:
            return None
        for index, image in enumerate(self._visible_images):
            if image.id == current.id:
                return index
        return None

    def _page_step(self) -> int:
        row_span = self.CARD_HEIGHT + self.CARD_PAD_Y * 2
        visible_rows = max(1, max(self.canvas.winfo_height(), 1) // row_span)
        return max(1, visible_rows * self._grid_columns)

    def _set_preview_by_index(self, index: int) -> None:
        target = self._visible_images[index]
        self.set_preview(target.id)
        self._scroll_index_into_view(index)

    def _scroll_index_into_view(self, index: int) -> None:
        row_span = self.CARD_HEIGHT + self.CARD_PAD_Y * 2
        row = index // self._grid_columns
        target_top = row * row_span
        target_bottom = target_top + row_span
        viewport_height = max(self.canvas.winfo_height(), 1)
        top_y = self.canvas.canvasy(0)
        bottom_y = top_y + viewport_height
        scrollregion = self.canvas.cget("scrollregion")
        if not scrollregion:
            return
        try:
            _, _, _, total_height = (float(value) for value in scrollregion.split())
        except Exception:
            return
        if total_height <= viewport_height:
            return
        if target_top < top_y:
            self.canvas.yview_moveto(target_top / total_height)
        elif target_bottom > bottom_y:
            desired_top = max(0.0, target_bottom - viewport_height)
            self.canvas.yview_moveto(desired_top / total_height)
        self._schedule_viewport_render(force=True)

    def _column_span(self) -> int:
        return self.CARD_WIDTH + self.CARD_PAD_X * 2

    def _compute_grid_columns(self) -> int:
        canvas_width = self.canvas.winfo_width()
        if canvas_width <= 1:
            return self._grid_columns
        max_columns = max(self.MIN_GRID_COLUMNS, min(self.GRID_COLUMNS, canvas_width // self._column_span()))
        return max_columns

    def _update_grid_layout_metrics(self, *, reset_x: bool = False) -> None:
        new_columns = self._compute_grid_columns()
        if new_columns != self._grid_columns:
            self._grid_columns = new_columns
            self._last_render_range = None
        total_rows = (len(self._visible_images) + self._grid_columns - 1) // self._grid_columns
        total_width = max(self._column_span(), self._grid_columns * self._column_span())
        total_height = total_rows * (self.CARD_HEIGHT + self.CARD_PAD_Y * 2)
        self.canvas.configure(scrollregion=(0, 0, total_width, total_height))
        if reset_x:
            self.canvas.xview_moveto(0)

    def _ensure_slot_pool(self) -> None:
        canvas_height = max(self.canvas.winfo_height(), 1)
        visible_rows = max(1, (canvas_height + self.CARD_HEIGHT + self.CARD_PAD_Y * 2 - 1) // (self.CARD_HEIGHT + self.CARD_PAD_Y * 2))
        needed_slots = (visible_rows + self.OVERSCAN_ROWS * 2) * self._grid_columns
        while len(self._slot_pool) < needed_slots:
            self._slot_pool.append(self._create_slot())

    def _create_slot(self) -> dict[str, object]:
        frame = tk.Frame(self.canvas, bg="#ffffff", bd=1, relief="solid", width=self.CARD_WIDTH, height=self.CARD_HEIGHT)
        frame.pack_propagate(False)

        content = tk.Frame(frame, bg="#ffffff")
        content.pack(fill="both", expand=True, padx=6, pady=6)

        top = tk.Frame(content, bg="#ffffff")
        top.pack(fill="x")

        var = tk.BooleanVar(value=False)
        check = ttk.Checkbutton(top, variable=var)
        check.pack(side="left")

        duplicate_label = tk.Label(top, text="唯一", bg="#ffffff", fg="#6b7280")
        duplicate_label.pack(side="right")

        image_box = tk.Frame(content, bg="#ffffff", width=170, height=120)
        image_box.pack(fill="x", pady=(6, 6))
        image_box.pack_propagate(False)

        image_label = tk.Label(image_box, bg="#ffffff", cursor="hand2")
        image_label.pack(fill="both", expand=True)

        name_label = tk.Label(content, text="", bg="#ffffff", anchor="w", justify="left")
        name_label.pack(fill="x")
        meta_label = tk.Label(content, text="", bg="#ffffff", anchor="w", justify="left", fg="#6b7280")
        meta_label.pack(fill="x")
        resolution_label = tk.Label(content, text="", bg="#ffffff", anchor="w", justify="left")
        resolution_label.pack(fill="x")
        size_label = tk.Label(content, text="", bg="#ffffff", anchor="w", justify="left")
        size_label.pack(fill="x")

        window_id = self.canvas.create_window(
            0,
            0,
            window=frame,
            anchor="nw",
            width=self.CARD_WIDTH,
            height=self.CARD_HEIGHT,
            state="hidden",
        )

        return {
            "window_id": window_id,
            "frame": frame,
            "content": content,
            "top": top,
            "check_var": var,
            "check": check,
            "duplicate_label": duplicate_label,
            "image_box": image_box,
            "image_label": image_label,
            "name_label": name_label,
            "meta_label": meta_label,
            "resolution_label": resolution_label,
            "size_label": size_label,
            "image_id": None,
        }

    def _schedule_viewport_render(self, force: bool = False) -> None:
        if self._viewport_after_id is not None:
            self.after_cancel(self._viewport_after_id)
        self._viewport_after_id = self.after(16, lambda: self._rerender_visible_slots(force=force))

    def _rerender_visible_slots(self, force: bool = False) -> None:
        self._viewport_after_id = None
        if not self._visible_images:
            self._hide_all_slots()
            return

        self._ensure_slot_pool()

        top_y = self.canvas.canvasy(0)
        bottom_y = top_y + max(self.canvas.winfo_height(), 1)
        row_span = self.CARD_HEIGHT + self.CARD_PAD_Y * 2
        start_row = max(0, int(top_y // row_span) - self.OVERSCAN_ROWS)
        end_row = min(
            (len(self._visible_images) + self._grid_columns - 1) // self._grid_columns,
            int(bottom_y // row_span) + self.OVERSCAN_ROWS + 1,
        )

        start_index = start_row * self._grid_columns
        end_index = min(len(self._visible_images), end_row * self._grid_columns)
        current_range = (start_index, end_index)
        if not force and self._last_render_range == current_range:
            return
        self._last_render_range = current_range

        visible_slice = self._visible_images[start_index:end_index]
        current_preview = self.service.current_preview()
        current_preview_id = current_preview.id if current_preview is not None else None
        for offset, (slot, image) in enumerate(zip(self._slot_pool, visible_slice)):
            index = start_index + offset
            row = index // self._grid_columns
            column = index % self._grid_columns
            x = column * self._column_span() + self.CARD_PAD_X
            y = row * (self.CARD_HEIGHT + self.CARD_PAD_Y * 2) + self.CARD_PAD_Y
            window_id = slot["window_id"]
            if slot.get("image_id") != image.id:
                self.canvas.itemconfigure(window_id, state="hidden")
            self.canvas.coords(window_id, x, y)
            self._bind_slot(slot, image, current_preview_id)
            self.canvas.itemconfigure(window_id, state="normal")

        for slot in self._slot_pool[len(visible_slice):]:
            self.canvas.itemconfigure(slot["window_id"], state="hidden")
            slot["image_id"] = None

        self._prefetch_grid_thumbnails(start_index, end_index)

    def _bind_slot(self, slot: dict[str, object], image: ImageAsset, current_preview_id: str | None) -> None:
        image_id = slot.get("image_id")
        background = "#eaf4ff" if current_preview_id == image.id else "#ffffff"

        frame = slot["frame"]
        content = slot["content"]
        top = slot["top"]
        duplicate_label = slot["duplicate_label"]
        image_box = slot["image_box"]
        image_label = slot["image_label"]
        name_label = slot["name_label"]
        meta_label = slot["meta_label"]
        resolution_label = slot["resolution_label"]
        size_label = slot["size_label"]
        check = slot["check"]
        check_var = slot["check_var"]

        for widget in (frame, content, top, duplicate_label, image_box, image_label, name_label, meta_label, resolution_label, size_label):
            widget.configure(bg=background)

        check_var.set(image.selected)
        check.configure(command=lambda image_id=image.id, var=check_var: self.on_check(image_id, var.get()))

        duplicate_ids = self.service.require_session().duplicate_groups.get(image.md5, [])
        if duplicate_ids:
            duplicate_label.configure(text=f"重复 {len(duplicate_ids)}", fg="#1d4ed8", cursor="hand2")
            duplicate_label.bind("<Button-1>", lambda _event, md5=image.md5: self.on_select_duplicate_group(md5))
        else:
            duplicate_label.configure(text="唯一", fg="#6b7280", cursor="")
            duplicate_label.unbind("<Button-1>")

        if image_id != image.id:
            self._assign_grid_thumbnail(image, image_label)

        name_label.configure(text=self._truncate_text(image.name, 24))
        meta_label.configure(
            text=self._truncate_text(
                f"#{image.location.occurrence_index + 1} | {image.location.display_text}",
                32,
            )
        )
        resolution_label.configure(text=image.resolution_text)
        size_label.configure(text=f"{image.size_kb} KB")

        for widget in (image_label, name_label, meta_label, resolution_label, size_label):
            widget.bind("<Button-1>", lambda _event, image_id=image.id: self.set_preview(image_id))
            widget.bind("<Double-Button-1>", lambda _event, image_id=image.id: self.on_locate(image_id))

        slot["image_id"] = image.id

    def _hide_all_slots(self) -> None:
        for slot in self._slot_pool:
            self.canvas.itemconfigure(slot["window_id"], state="hidden")
            slot["image_id"] = None

    def _on_scrollbar(self, *args) -> None:
        self.canvas.yview(*args)
        self._schedule_viewport_render()

    def _on_horizontal_scroll(self, *args) -> None:
        self.canvas.xview(*args)

    def _on_canvas_configure(self, _event: tk.Event) -> None:
        previous_columns = self._grid_columns
        self._update_grid_layout_metrics(reset_x=previous_columns != self._compute_grid_columns())
        self._ensure_slot_pool()
        self._schedule_viewport_render()

    def _on_mousewheel(self, event: tk.Event) -> None:
        if self.canvas.winfo_exists():
            step = int(-1 * (event.delta / 120))
            if step == 0:
                step = -1 if event.delta > 0 else 1
            self.canvas.yview_scroll(step, "units")
            self._schedule_viewport_render()

    def _on_shift_mousewheel(self, event: tk.Event) -> None:
        if self.canvas.winfo_exists():
            step = int(-1 * (event.delta / 120))
            if step == 0:
                step = -1 if event.delta > 0 else 1
            self.canvas.xview_scroll(step, "units")

    def _assign_grid_thumbnail(self, image: ImageAsset, image_label: tk.Label) -> None:
        cache_key = f"{image.id}:170x120"
        photo = self.thumbnail_refs.get(cache_key)
        if photo is not None:
            image_label.configure(image=photo)
            image_label.image = photo
            return

        image_label.configure(image=self._placeholder_thumbnail)
        image_label.image = self._placeholder_thumbnail

        if cache_key in self._thumbnail_pending:
            return

        self._thumbnail_pending.add(cache_key)
        generation = self._thumbnail_generation
        threading.Thread(
            target=self._build_grid_thumbnail_worker,
            args=(generation, cache_key, image),
            daemon=True,
        ).start()

    def _prefetch_grid_thumbnails(self, start_index: int, end_index: int) -> None:
        prefetch_start = max(0, start_index - self._grid_columns * 2)
        prefetch_end = min(len(self._visible_images), end_index + self._grid_columns * 3)
        generation = self._thumbnail_generation
        for image in self._visible_images[prefetch_start:prefetch_end]:
            cache_key = f"{image.id}:170x120"
            if cache_key in self.thumbnail_refs or cache_key in self._thumbnail_pending:
                continue
            self._thumbnail_pending.add(cache_key)
            threading.Thread(
                target=self._build_grid_thumbnail_worker,
                args=(generation, cache_key, image),
                daemon=True,
            ).start()

    def _build_grid_thumbnail_worker(self, generation: int, cache_key: str, image: ImageAsset) -> None:
        pil_image = self.service.build_thumbnail(image, (170, 120))
        self._thumbnail_result_queue.put((generation, cache_key, pil_image))

    def _poll_thumbnail_results(self) -> None:
        processed = 0
        while processed < 12:
            try:
                generation, cache_key, pil_image = self._thumbnail_result_queue.get_nowait()
            except queue.Empty:
                break

            self._thumbnail_pending.discard(cache_key)
            if generation != self._thumbnail_generation:
                processed += 1
                continue

            photo = ImageTk.PhotoImage(pil_image)
            self.thumbnail_refs[cache_key] = photo
            image_id = cache_key.rsplit(":", 1)[0]
            for slot in self._slot_pool:
                if slot.get("image_id") == image_id:
                    image_label = slot["image_label"]
                    image_label.configure(image=photo)
                    image_label.image = photo
            processed += 1

        self.after(50, self._poll_thumbnail_results)

    def _truncate_text(self, text: str, max_length: int) -> str:
        if len(text) <= max_length:
            return text
        return f"{text[: max_length - 1]}…"


def run() -> None:
    app = ImageExtractorApp()
    app.mainloop()
