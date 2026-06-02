from __future__ import annotations

from pathlib import Path
from tkinter import filedialog, messagebox

from app.docx_adapter import DocumentError, NavigationResult, UnsupportedDocumentError
from app.models import ImageAsset

from .dialogs import DeleteConfirmDialog, FilteredImagesDialog, ImageViewerDialog


class ImageExtractorActionsMixin:
    def _display_size_mode(self, mode: str) -> str:
        for label, value in self.SIZE_FILTER_MODE_MAP.items():
            if value == mode:
                return label
        return "精确"

    def _snapshot_session_state(self) -> dict[str, object]:
        session = self.service.require_session()
        return {
            "query_text": self.query_var.get(),
            "size_filter_mode": self.SIZE_FILTER_MODE_MAP.get(self.size_mode_var.get().strip(), "exact"),
            "size_filter_width": self.width_filter_var.get(),
            "size_filter_height": self.height_filter_var.get(),
            "preview_image_id": session.preview_image_id,
            "selected_ids": {image.id for image in session.images if image.selected},
            "hidden_ids": {image.id for image in session.images if image.hidden},
        }

    def _restore_session_state(self, snapshot: dict[str, object]) -> None:
        session = self.service.require_session()
        self.query_var.set(str(snapshot["query_text"]))
        self.size_mode_var.set(self._display_size_mode(str(snapshot["size_filter_mode"])))
        self.width_filter_var.set(str(snapshot["size_filter_width"]))
        self.height_filter_var.set(str(snapshot["size_filter_height"]))

        selected_ids = snapshot["selected_ids"]
        hidden_ids = snapshot["hidden_ids"]
        if isinstance(selected_ids, set) and isinstance(hidden_ids, set):
            for image in session.images:
                image.hidden = image.id in hidden_ids
                image.selected = image.id in selected_ids

        preview_image_id = snapshot["preview_image_id"]
        if isinstance(preview_image_id, str) and session.get_image(preview_image_id) is not None:
            session.preview_image_id = preview_image_id

    def reload_document_preserving_state(self, *, auto_triggered: bool = False) -> bool:
        try:
            current_path = self.service.require_session().document_path
            snapshot = self._snapshot_session_state()
        except DocumentError as exc:
            if auto_triggered:
                self.status_var.set(f"自动重载失败：{exc}")
                return False
            messagebox.showwarning("提示", str(exc), parent=self)
            return False

        try:
            session = self.service.open_document(current_path)
        except (DocumentError, OSError) as exc:
            if auto_triggered:
                notice = f"检测到文档变更，但自动重载失败：{exc}"
                if notice != self._document_watch_last_notice:
                    self.status_var.set(notice)
                    self._document_watch_last_notice = notice
                return False
            messagebox.showerror("重新加载失败", str(exc), parent=self)
            self.status_var.set(f"重新加载失败：{exc}")
            return False

        self._restore_session_state(snapshot)
        self._reset_document_view_state()
        self._update_document_watch_baseline()
        if auto_triggered:
            notice = f"检测到文档变更，已自动重载 {session.document_path.name}，共提取 {session.total_images} 张图片。"
            self.status_var.set(notice)
            self._document_watch_last_notice = notice
        self.refresh_all(reset_grid_x=True, force_grid=True)
        self._focus_grid_area()
        return True

    def _current_preview_target_size(self) -> tuple[int, int]:
        image_area_width = self.preview_image_area.winfo_width()
        image_area_height = self.preview_image_area.winfo_height()
        if image_area_width > 32 and image_area_height > 32:
            available_width = max(220, image_area_width - 8)
            available_height = max(220, image_area_height - 8)
            return (available_width, available_height)

        return (320, 320)

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
        self._update_document_watch_baseline()
        self.status_var.set(f"已打开 {session.document_path.name}，共提取 {session.total_images} 张图片。")
        self.refresh_all(reset_grid_x=True, force_grid=True)
        self._focus_grid_area()

    def reload_document(self) -> None:
        success = self.reload_document_preserving_state(auto_triggered=False)
        if not success:
            return
        try:
            session = self.service.require_session()
        except DocumentError:
            return
        self.status_var.set(f"已重新加载 {session.document_path.name}，共提取 {session.total_images} 张图片。")

    def refresh_all(self, *, reset_grid_x: bool = False, force_grid: bool = True) -> None:
        try:
            session = self.service.require_session()
        except DocumentError:
            return
        if not self._sync_filter_state():
            return

        self.doc_name_var.set(session.document_path.name)
        self.total_var.set(f"总图片数：{session.total_images}")
        self.duplicate_var.set(f"重复组：{session.duplicate_group_count}")
        self.selection_var.set(f"已选：{session.selected_count}")
        self.filtered_var.set(f"已过滤：{session.hidden_count}")

        preview = self.service.current_preview()
        visible = self.service.visible_images()
        visible_ids = {image.id for image in visible}
        if visible and (preview is None or preview.hidden or preview.id not in visible_ids):
            self.service.set_preview(visible[0].id)
        self.refresh_grid(recompute_visible=True, reset_x=reset_grid_x, force=force_grid)
        self.refresh_preview()

    def refresh_grid(self, *, recompute_visible: bool = True, reset_x: bool = False, force: bool = False) -> None:
        try:
            session = self.service.require_session()
        except DocumentError:
            return
        if recompute_visible and not self._sync_filter_state():
            return

        visible_changed = False
        if recompute_visible:
            previous_ids = [image.id for image in self._visible_images]
            try:
                new_visible_images = self.service.visible_images()
            except DocumentError:
                return
            new_ids = [image.id for image in new_visible_images]
            visible_changed = new_ids != previous_ids
            self._visible_images = new_visible_images

        if visible_changed:
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

        self._update_grid_layout_metrics(reset_x=reset_x)
        self._ensure_slot_pool()
        self._schedule_viewport_render(force=force or visible_changed)

    def refresh_preview(self) -> None:
        try:
            current = self.service.current_preview()
        except DocumentError:
            return

        self.preview_info_label.configure(wraplength=max(220, self.preview_panel.winfo_width() - 24))

        if current is None:
            self.preview_label.configure(text="没有可预览图片", image="")
            self.preview_info_var.set("请先打开文档。")
            return

        preview = self.build_photo(current, self._current_preview_target_size(), self.preview_refs)
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

    def open_preview_viewer_from_panel(self) -> None:
        try:
            current = self.service.current_preview()
        except DocumentError as exc:
            messagebox.showwarning("提示", str(exc), parent=self)
            return
        if current is None:
            messagebox.showwarning("提示", "当前没有可查看的预览图片。", parent=self)
            return
        viewer = self._get_or_create_image_viewer(current)
        viewer.deiconify()
        viewer.lift()
        viewer.focus_force()

    def _get_or_create_image_viewer(self, image: ImageAsset) -> ImageViewerDialog:
        viewer = self.image_viewer
        if viewer is None or not viewer.winfo_exists():
            viewer = ImageViewerDialog(self, image)
            self.image_viewer = viewer
            return viewer
        viewer.set_image(image)
        return viewer

    def set_preview(self, image_id: str) -> None:
        try:
            self.service.set_preview(image_id)
        except DocumentError as exc:
            messagebox.showerror("预览失败", str(exc), parent=self)
            return
        self._refresh_visible_slot_styles()
        self.refresh_preview()
        self._focus_grid_area()

    def set_preview_from_viewer(self, image_id: str) -> bool:
        try:
            self.service.set_preview(image_id)
        except DocumentError:
            return False
        target_index = None
        for index, image in enumerate(self._visible_images):
            if image.id == image_id:
                target_index = index
                break
        self._refresh_visible_slot_styles()
        self.refresh_preview()
        if target_index is not None:
            self._scroll_index_into_view(target_index)
        return True

    def on_check(self, image_id: str, selected: bool) -> None:
        try:
            self.service.toggle_selection(image_id, selected)
            session = self.service.require_session()
        except DocumentError as exc:
            messagebox.showerror("勾选失败", str(exc), parent=self)
            return
        self.selection_var.set(f"已选：{session.selected_count}")
        self._refresh_visible_slot_styles()
        self._focus_grid_area()

    def select_all(self) -> None:
        if not self._sync_filter_state():
            return
        try:
            self.service.select_all_visible()
        except DocumentError as exc:
            messagebox.showwarning("提示", str(exc), parent=self)
            return
        self.status_var.set("已选中当前可见图片。")
        self._refresh_visible_slot_styles()
        self.refresh_all_stats()
        self._focus_grid_area()

    def invert_selection(self) -> None:
        if not self._sync_filter_state():
            return
        try:
            self.service.invert_visible()
        except DocumentError as exc:
            messagebox.showwarning("提示", str(exc), parent=self)
            return
        self.status_var.set("已对当前可见图片执行反选。")
        self._refresh_visible_slot_styles()
        self.refresh_all_stats()
        self._focus_grid_area()

    def on_select_duplicate_group(self, md5: str) -> None:
        try:
            count = self.service.select_same_md5(md5)
            session = self.service.require_session()
        except DocumentError as exc:
            messagebox.showwarning("提示", str(exc), parent=self)
            return
        self.selection_var.set(f"已选：{session.selected_count}")
        self.status_var.set(f"已通过重复组入口选中 {count} 张相同图片。")
        self._refresh_visible_slot_styles()
        self._focus_grid_area()

    def filter_selected(self) -> None:
        try:
            count = self.service.filter_selected()
        except DocumentError as exc:
            messagebox.showwarning("提示", str(exc), parent=self)
            return
        self.status_var.set(f"已隐藏 {count} 张选中图片。")
        self.refresh_all(reset_grid_x=False, force_grid=True)
        self._focus_grid_area()

    def clear_filters(self) -> None:
        try:
            self.service.clear_filters()
        except DocumentError as exc:
            messagebox.showwarning("提示", str(exc), parent=self)
            return
        self.status_var.set("已恢复全部过滤项。")
        self.refresh_all(reset_grid_x=False, force_grid=True)
        self._focus_grid_area()

    def view_filtered_images(self) -> None:
        try:
            filtered = self.service.filtered_images()
        except DocumentError as exc:
            messagebox.showwarning("提示", str(exc), parent=self)
            return
        if not filtered:
            messagebox.showwarning("提示", "当前没有已过滤图片。", parent=self)
            return

        dialog = FilteredImagesDialog(self, filtered)
        self.wait_window(dialog)
        self._focus_grid_area()

    def export_filtered_md5s(self) -> None:
        try:
            filtered = self.service.filtered_images()
        except DocumentError as exc:
            messagebox.showwarning("提示", str(exc), parent=self)
            return
        if not filtered:
            messagebox.showwarning("提示", "当前没有已过滤图片。", parent=self)
            return

        try:
            document_name = self.service.require_session().document_path.stem
        except DocumentError as exc:
            messagebox.showwarning("提示", str(exc), parent=self)
            return

        target = filedialog.asksaveasfilename(
            title="导出已过滤图片 MD5",
            defaultextension=".txt",
            initialfile=f"{document_name}_filtered_md5.txt",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")],
        )
        if not target:
            return

        try:
            count = self.service.export_filtered_md5s(Path(target))
        except (DocumentError, OSError) as exc:
            messagebox.showerror("导出失败", str(exc), parent=self)
            self.status_var.set(f"导出过滤 MD5 失败：{exc}")
            return

        self.status_var.set(f"已导出 {count} 条过滤 MD5 到 {Path(target).name}")
        messagebox.showinfo("导出完成", f"已导出 {count} 条 MD5。\n文件：{target}", parent=self)
        self._focus_grid_area()

    def import_filtered_md5s(self) -> None:
        try:
            self.service.require_session()
        except DocumentError as exc:
            messagebox.showwarning("提示", str(exc), parent=self)
            return

        source = filedialog.askopenfilename(
            title="导入过滤 MD5",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")],
        )
        if not source:
            return

        try:
            hidden_count, matched_md5_count = self.service.import_filtered_md5s(Path(source))
        except DocumentError as exc:
            messagebox.showerror("导入失败", str(exc), parent=self)
            self.status_var.set(f"导入过滤 MD5 失败：{exc}")
            return

        self.status_var.set(f"已按导入 MD5 过滤 {hidden_count} 张图片，命中 {matched_md5_count} 个 MD5。")
        messagebox.showinfo(
            "导入完成",
            f"本次新过滤图片：{hidden_count} 张\n命中 MD5：{matched_md5_count} 个\n文件：{source}",
            parent=self,
        )
        self.refresh_all(reset_grid_x=False, force_grid=True)
        self._focus_grid_area()

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
        self._focus_grid_area()

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
        self._update_document_watch_baseline()
        self.refresh_all(reset_grid_x=True, force_grid=True)
        self._focus_grid_area()

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
        self._focus_grid_area()

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
        self._reset_document_watch_state()
        self._thumbnail_generation += 1
        self._thumbnail_pending.clear()
        self._last_render_range = None
        self._destroy_slot_pool()
        self.preview_label.configure(image="", text="单击缩略图后显示大图预览\n双击此处打开图片查看器")
        self.preview_info_label.configure(wraplength=max(220, self.preview_panel.winfo_width() - 24))

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
