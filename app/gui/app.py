from __future__ import annotations

import queue
import tkinter as tk
from pathlib import Path

from PIL import Image, ImageTk
from tkinter import messagebox, ttk

from app.docx_adapter import DocumentError
from app.models import ImageAsset
from app.services import SessionService

from .actions import ImageExtractorActionsMixin
from .grid import ImageExtractorGridMixin


class ImageExtractorApp(ImageExtractorActionsMixin, ImageExtractorGridMixin, tk.Tk):
    GRID_COLUMNS = 4
    MIN_GRID_COLUMNS = 1
    CARD_WIDTH = 210
    CARD_HEIGHT = 236
    CARD_PAD_X = 10
    CARD_PAD_Y = 10
    OVERSCAN_ROWS = 2
    PREVIEW_MIN_WIDTH = 320
    PREVIEW_MAX_WIDTH = 520
    PREVIEW_LAYOUT_GAP = 24
    CENTER_LAYOUT_PADDING_X = 20
    DOCUMENT_WATCH_INTERVAL_MS = 1200
    SIZE_FILTER_MODE_MAP = {"精确": "exact", "最小": "min", "最大": "max"}

    def __init__(self) -> None:
        super().__init__()
        self.title("WPS Image Extractor (Tkinter)")
        self.geometry("1380x860")
        self.minsize(420, 720)

        self.service = SessionService()
        self.query_var = tk.StringVar()
        self.size_mode_var = tk.StringVar(value="精确")
        self.width_filter_var = tk.StringVar(value="*")
        self.height_filter_var = tk.StringVar(value="*")
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
        self.image_viewer = None
        self._visible_images: list[ImageAsset] = []
        self._slot_pool: list[dict[str, object]] = []
        self._viewport_after_id: str | None = None
        self._filter_refresh_after_id: str | None = None
        self._preview_refresh_after_id: str | None = None
        self._body_layout_after_id: str | None = None
        self._document_watch_after_id: str | None = None
        self._empty_label_id: int | None = None
        self._last_render_range: tuple[int, int] | None = None
        self._last_preview_panel_size: tuple[int, int] | None = None
        self._last_body_layout_signature: tuple[int, int, int, int] | None = None
        self._allocated_grid_columns: int | None = None
        self._document_watch_signature: tuple[int, int] | None = None
        self._document_watch_in_progress = False
        self._document_watch_failed_signature: tuple[int, int] | None = None
        self._document_watch_last_notice: str | None = None
        self._thumbnail_result_queue: queue.Queue[tuple[int, str, Image.Image]] = queue.Queue()
        self._thumbnail_queue_empty_exception = queue.Empty
        self._thumbnail_pending: set[str] = set()
        self._thumbnail_generation = 0
        self._grid_columns = self.GRID_COLUMNS
        self._placeholder_thumbnail = ImageTk.PhotoImage(Image.new("RGBA", (170, 120), "#e5e7eb"))

        self._build_layout()
        self._bind_keyboard_navigation()
        self.query_var.trace_add("write", lambda *_args: self._schedule_filter_refresh())
        self.width_filter_var.trace_add("write", lambda *_args: self._schedule_filter_refresh())
        self.height_filter_var.trace_add("write", lambda *_args: self._schedule_filter_refresh())
        self.after(50, self._poll_thumbnail_results)
        self.after(self.DOCUMENT_WATCH_INTERVAL_MS, self._poll_document_change)

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
        ttk.Button(toolbar, text="查看已过滤...", command=self.view_filtered_images).pack(side="left", padx=(8, 0))
        ttk.Button(toolbar, text="导出过滤MD5...", command=self.export_filtered_md5s).pack(side="left", padx=(8, 0))
        ttk.Button(toolbar, text="导入过滤MD5...", command=self.import_filtered_md5s).pack(side="left", padx=(8, 0))
        ttk.Button(toolbar, text="导出所选...", command=self.export_selected).pack(side="left", padx=(8, 0))

        self.body = ttk.Frame(root)
        self.body.pack(fill="both", expand=True)
        self.body.grid_rowconfigure(0, weight=1)
        self.body.grid_columnconfigure(0, weight=0)
        self.body.grid_columnconfigure(1, weight=0)
        self.body.grid_columnconfigure(2, weight=1)
        self.body.bind("<Configure>", self._on_body_configure)

        self.left_panel = ttk.LabelFrame(self.body, text="文档信息", padding=10)
        self.left_panel.grid(row=0, column=0, sticky="ns")

        ttk.Label(self.left_panel, textvariable=self.doc_name_var, wraplength=220).pack(anchor="w")
        ttk.Label(self.left_panel, textvariable=self.total_var).pack(anchor="w", pady=(8, 0))
        ttk.Label(self.left_panel, textvariable=self.duplicate_var).pack(anchor="w", pady=(4, 0))
        ttk.Label(self.left_panel, textvariable=self.selection_var).pack(anchor="w", pady=(4, 0))
        ttk.Label(self.left_panel, textvariable=self.filtered_var).pack(anchor="w", pady=(4, 0))

        ttk.Separator(self.left_panel).pack(fill="x", pady=10)
        ttk.Label(self.left_panel, text="搜索文件名 / MD5 / 部件").pack(anchor="w")
        self.search_entry = ttk.Entry(self.left_panel, textvariable=self.query_var, width=26)
        self.search_entry.pack(fill="x", pady=(6, 0))

        ttk.Label(self.left_panel, text="尺寸过滤").pack(anchor="w", pady=(12, 0))
        self.size_mode_combo = ttk.Combobox(
            self.left_panel,
            textvariable=self.size_mode_var,
            values=list(self.SIZE_FILTER_MODE_MAP),
            state="readonly",
            width=26,
        )
        self.size_mode_combo.pack(fill="x", pady=(6, 0))
        self.size_mode_combo.bind("<<ComboboxSelected>>", lambda _event: self._schedule_filter_refresh(delay_ms=0))

        size_row = ttk.Frame(self.left_panel)
        size_row.pack(fill="x", pady=(6, 0))
        ttk.Label(size_row, text="宽").pack(side="left")
        self.width_filter_entry = ttk.Entry(size_row, textvariable=self.width_filter_var, width=10)
        self.width_filter_entry.pack(side="left", padx=(6, 8))
        ttk.Label(size_row, text="高").pack(side="left")
        self.height_filter_entry = ttk.Entry(size_row, textvariable=self.height_filter_var, width=10)
        self.height_filter_entry.pack(side="left", padx=(6, 0))
        ttk.Label(self.left_panel, text="输入正整数或 *，* 表示该维度不参与").pack(anchor="w", pady=(4, 0))

        ttk.Label(
            self.left_panel,
            text="交互说明：\n- 单击缩略图：仅预览\n- 复选框：勾选图片\n- 双击缩略图：定位文档\n- HJKL：Vim 式移动\n- O 键：单开当前图片\n- 双击预览栏：打开图片查看器",
            justify="left",
        ).pack(anchor="w", pady=(12, 0))

        self.center_panel = ttk.LabelFrame(self.body, text="缩略图区", padding=0)
        self.center_panel.grid(row=0, column=1, sticky="nsew", padx=(10, 10))

        canvas_area = ttk.Frame(self.center_panel)
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

        self.preview_panel = ttk.LabelFrame(self.body, text="预览", padding=10)
        self.preview_panel.grid(row=0, column=2, sticky="nsew")
        self.preview_panel.bind("<Configure>", self._on_preview_panel_configure)

        self.preview_image_area = ttk.Frame(self.preview_panel)
        self.preview_image_area.pack(fill="both", expand=True)
        self.preview_image_area.pack_propagate(False)

        self.preview_label = ttk.Label(
            self.preview_image_area,
            text="单击缩略图后显示大图预览\n双击此处打开图片查看器",
            anchor="center",
        )
        self.preview_label.place(relx=0.5, rely=0.5, anchor="center")
        self.preview_label.bind("<Double-Button-1>", lambda _event: self.open_preview_viewer_from_panel())
        self.preview_info_label = ttk.Label(self.preview_panel, textvariable=self.preview_info_var, justify="left", wraplength=260)
        self.preview_info_label.pack(fill="x", pady=(10, 0))
        self.after_idle(self._schedule_body_layout)

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

    def _sync_filter_state(self) -> bool:
        if self.service.session is None:
            return True
        try:
            mode = self.SIZE_FILTER_MODE_MAP.get(self.size_mode_var.get().strip())
            if mode is None:
                raise DocumentError("尺寸过滤模式无效。")
            self.service.set_text_query(self.query_var.get())
            self.service.set_size_filter(mode, self.width_filter_var.get(), self.height_filter_var.get())
        except DocumentError as exc:
            self.status_var.set(f"筛选条件无效：{exc}")
            messagebox.showwarning("筛选条件无效", str(exc), parent=self)
            return False
        return True

    def _focus_grid_area(self) -> None:
        if self.canvas.winfo_exists():
            self.after_idle(self.canvas.focus_set)

    def _schedule_filter_refresh(self, delay_ms: int = 150) -> None:
        if self._filter_refresh_after_id is not None:
            self.after_cancel(self._filter_refresh_after_id)
        self._filter_refresh_after_id = self.after(delay_ms, self._apply_filter_refresh)

    def _apply_filter_refresh(self) -> None:
        self._filter_refresh_after_id = None
        self.refresh_grid(recompute_visible=True, reset_x=False, force=False)

    def _required_center_width(self, columns: int) -> int:
        scrollbar_width = max(16, self.scrollbar.winfo_reqwidth())
        return max(self._column_span(), columns * self._column_span()) + scrollbar_width + 8

    def _compute_allocated_grid_columns(self, available_width: int) -> int:
        for columns in range(self.GRID_COLUMNS, self.MIN_GRID_COLUMNS - 1, -1):
            if self._required_center_width(columns) <= available_width:
                return columns
        return self.MIN_GRID_COLUMNS

    def _on_body_configure(self, _event: tk.Event) -> None:
        self._schedule_body_layout()

    def _schedule_body_layout(self, delay_ms: int = 0) -> None:
        if self._body_layout_after_id is not None:
            self.after_cancel(self._body_layout_after_id)
        self._body_layout_after_id = self.after(delay_ms, self._apply_body_layout)

    def _apply_body_layout(self) -> None:
        self._body_layout_after_id = None
        total_width = self.body.winfo_width()
        left_width = max(self.left_panel.winfo_width(), self.left_panel.winfo_reqwidth())
        total_content_width = max(
            self._required_center_width(self.MIN_GRID_COLUMNS),
            total_width - left_width - self.CENTER_LAYOUT_PADDING_X,
        )
        min_center_width = self._required_center_width(self.MIN_GRID_COLUMNS)
        layout_gap = min(
            self.PREVIEW_LAYOUT_GAP,
            max(0, total_content_width - min_center_width - self.PREVIEW_MIN_WIDTH),
        )
        preview_reserved_width = self.PREVIEW_MIN_WIDTH
        reserved_preview_width = min(
            preview_reserved_width,
            max(0, total_content_width - min_center_width - layout_gap),
        )
        center_available_width = max(min_center_width, total_content_width - reserved_preview_width - layout_gap)
        columns = self._compute_allocated_grid_columns(center_available_width)
        center_width = min(center_available_width, self._required_center_width(columns))
        preview_width = max(0, total_content_width - center_width - layout_gap)
        if preview_width > 0:
            preview_width_cap = max(0, total_content_width - min_center_width - layout_gap)
            preview_width = min(
                preview_width_cap,
                self.PREVIEW_MAX_WIDTH,
                max(self.PREVIEW_MIN_WIDTH, preview_width),
            )
            center_width = max(min_center_width, total_content_width - preview_width - layout_gap)

        signature = (total_width, left_width, center_width, preview_width)
        if signature == self._last_body_layout_signature:
            return
        self._last_body_layout_signature = signature
        self._allocated_grid_columns = columns
        self.body.grid_columnconfigure(1, minsize=center_width, weight=1)
        self.body.grid_columnconfigure(2, minsize=max(0, preview_width), weight=0)
        self._schedule_preview_refresh()

    def _on_preview_panel_configure(self, _event: tk.Event) -> None:
        current_size = (self.preview_panel.winfo_width(), self.preview_panel.winfo_height())
        if current_size == self._last_preview_panel_size:
            return
        self._last_preview_panel_size = current_size
        self._schedule_preview_refresh()

    def _schedule_preview_refresh(self, delay_ms: int = 80) -> None:
        if self._preview_refresh_after_id is not None:
            self.after_cancel(self._preview_refresh_after_id)
        self._preview_refresh_after_id = self.after(delay_ms, self._apply_preview_refresh)

    def _apply_preview_refresh(self) -> None:
        self._preview_refresh_after_id = None
        wraplength = max(220, self.preview_panel.winfo_width() - 24)
        self.preview_info_label.configure(wraplength=wraplength)
        self.refresh_preview()

    def _read_document_signature(self, document_path: Path) -> tuple[int, int] | None:
        try:
            stat = document_path.stat()
        except OSError:
            return None
        return (stat.st_mtime_ns, stat.st_size)

    def _reset_document_watch_state(self) -> None:
        self._document_watch_signature = None
        self._document_watch_failed_signature = None
        self._document_watch_last_notice = None

    def _update_document_watch_baseline(self) -> None:
        try:
            session = self.service.require_session()
        except DocumentError:
            self._reset_document_watch_state()
            return

        signature = self._read_document_signature(session.document_path)
        self._document_watch_signature = signature
        self._document_watch_failed_signature = None
        self._document_watch_last_notice = None

    def _schedule_document_watch(self, delay_ms: int | None = None) -> None:
        if self._document_watch_after_id is not None:
            self.after_cancel(self._document_watch_after_id)
        wait_ms = self.DOCUMENT_WATCH_INTERVAL_MS if delay_ms is None else delay_ms
        self._document_watch_after_id = self.after(wait_ms, self._poll_document_change)

    def _poll_document_change(self) -> None:
        self._document_watch_after_id = None
        try:
            session = self.service.require_session()
        except DocumentError:
            self._reset_document_watch_state()
            self._schedule_document_watch()
            return

        if self._document_watch_in_progress:
            self._schedule_document_watch()
            return

        current_signature = self._read_document_signature(session.document_path)
        if current_signature is None:
            notice = f"文档监听：暂时无法读取 {session.document_path.name}，等待下次检测。"
            if notice != self._document_watch_last_notice:
                self.status_var.set(notice)
                self._document_watch_last_notice = notice
            self._schedule_document_watch()
            return

        if self._document_watch_signature is None:
            self._document_watch_signature = current_signature
            self._document_watch_last_notice = None
            self._schedule_document_watch()
            return

        if current_signature == self._document_watch_signature:
            self._document_watch_failed_signature = None
            self._document_watch_last_notice = None
            self._schedule_document_watch()
            return

        self._document_watch_in_progress = True
        try:
            success = self.reload_document_preserving_state(auto_triggered=True)
        finally:
            self._document_watch_in_progress = False

        if success:
            self._update_document_watch_baseline()
        else:
            self._document_watch_failed_signature = current_signature

        self._schedule_document_watch()


def run() -> None:
    app = ImageExtractorApp()
    app.mainloop()
