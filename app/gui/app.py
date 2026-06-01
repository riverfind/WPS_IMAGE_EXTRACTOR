from __future__ import annotations

import queue
import tkinter as tk

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
        self._empty_label_id: int | None = None
        self._last_render_range: tuple[int, int] | None = None
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

        ttk.Label(left, text="尺寸过滤").pack(anchor="w", pady=(12, 0))
        self.size_mode_combo = ttk.Combobox(
            left,
            textvariable=self.size_mode_var,
            values=list(self.SIZE_FILTER_MODE_MAP),
            state="readonly",
            width=26,
        )
        self.size_mode_combo.pack(fill="x", pady=(6, 0))
        self.size_mode_combo.bind("<<ComboboxSelected>>", lambda _event: self._schedule_filter_refresh(delay_ms=0))

        size_row = ttk.Frame(left)
        size_row.pack(fill="x", pady=(6, 0))
        ttk.Label(size_row, text="宽").pack(side="left")
        self.width_filter_entry = ttk.Entry(size_row, textvariable=self.width_filter_var, width=10)
        self.width_filter_entry.pack(side="left", padx=(6, 8))
        ttk.Label(size_row, text="高").pack(side="left")
        self.height_filter_entry = ttk.Entry(size_row, textvariable=self.height_filter_var, width=10)
        self.height_filter_entry.pack(side="left", padx=(6, 0))
        ttk.Label(left, text="输入正整数或 *，* 表示该维度不参与").pack(anchor="w", pady=(4, 0))

        ttk.Label(
            left,
            text="交互说明：\n- 单击缩略图：仅预览\n- 复选框：勾选图片\n- 双击缩略图：定位文档\n- 双击预览栏：打开图片查看器",
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

        self.preview_label = ttk.Label(right, text="单击缩略图后显示大图预览\n双击此处打开图片查看器", anchor="center")
        self.preview_label.pack(fill="both", expand=True)
        self.preview_label.bind("<Double-Button-1>", lambda _event: self.open_preview_viewer_from_panel())
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


def run() -> None:
    app = ImageExtractorApp()
    app.mainloop()
