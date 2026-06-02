from __future__ import annotations

import threading
import tkinter as tk
from tkinter import ttk

from PIL import ImageTk

from app.docx_adapter import DocumentError
from app.models import ImageAsset


class ImageExtractorGridMixin:
    def _bind_keyboard_navigation(self) -> None:
        self.bind_all("<Up>", lambda event: self._on_grid_key(event, "up"))
        self.bind_all("<Down>", lambda event: self._on_grid_key(event, "down"))
        self.bind_all("<Left>", lambda event: self._on_grid_key(event, "left"))
        self.bind_all("<Right>", lambda event: self._on_grid_key(event, "right"))
        self.bind_all("<h>", lambda event: self._on_grid_key(event, "left"))
        self.bind_all("<j>", lambda event: self._on_grid_key(event, "down"))
        self.bind_all("<k>", lambda event: self._on_grid_key(event, "up"))
        self.bind_all("<l>", lambda event: self._on_grid_key(event, "right"))
        self.bind_all("<Prior>", lambda event: self._on_grid_key(event, "page_up"))
        self.bind_all("<Next>", lambda event: self._on_grid_key(event, "page_down"))
        self.bind_all("<Return>", self._on_enter_key)
        self.bind_all("<space>", self._on_space_key)
        self.bind_all("<o>", self._on_open_viewer_key)

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
        return widget_class in {"Entry", "TEntry", "Text", "Spinbox", "TSpinbox", "Combobox", "TCombobox"}

    def _on_enter_key(self, _event: tk.Event) -> str | None:
        if self._should_ignore_grid_key_event():
            return None
        try:
            current = self.service.current_preview()
        except DocumentError:
            return "break"
        if current is None:
            return "break"
        self.on_locate(current.id)
        return "break"

    def _on_space_key(self, _event: tk.Event) -> str | None:
        if self._should_ignore_grid_key_event():
            return None
        try:
            current = self.service.current_preview()
        except DocumentError:
            return "break"
        if current is None:
            return "break"
        new_selected = not current.selected
        self.on_check(current.id, new_selected)
        return "break"

    def _on_open_viewer_key(self, _event: tk.Event) -> str | None:
        if self._should_ignore_grid_key_event():
            return None
        try:
            current = self.service.current_preview()
        except DocumentError:
            return "break"
        if current is None:
            return "break"
        self.open_preview_viewer_from_panel()
        return "break"

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
        allocated_columns = getattr(self, "_allocated_grid_columns", None)
        if allocated_columns is not None:
            return allocated_columns
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

    def _refresh_visible_slot_styles(self) -> None:
        if not self._slot_pool or not self._visible_images:
            return

        visible_by_id = {image.id: image for image in self._visible_images}
        try:
            current_preview = self.service.current_preview()
        except DocumentError:
            return
        current_preview_id = current_preview.id if current_preview is not None else None

        for slot in self._slot_pool:
            image_id = slot.get("image_id")
            if not image_id:
                continue
            image = visible_by_id.get(image_id)
            if image is None:
                continue
            self._bind_slot(slot, image, current_preview_id)

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
            self.canvas.coords(window_id, x, y)
            self._bind_slot(slot, image, current_preview_id)
            if str(self.canvas.itemcget(window_id, "state")) != "normal":
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
            except self._thumbnail_queue_empty_exception:
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
