from __future__ import annotations

import os
import time
from pathlib import Path

from app.models import ImageAsset, ImageLocation

from .errors import DocumentError, NavigationResult


class OfficeAutomationMixin:
    document_path: Path

    def locate(self, image: ImageAsset) -> NavigationResult:
        try:
            import pythoncom  # type: ignore
            import win32com.client  # type: ignore
        except Exception:
            pythoncom = None
            win32com_client = None
        else:
            win32com_client = win32com.client

        if win32com_client is not None and image.location.part_name == "word/document.xml" and image.location.block_index:
            try:
                pythoncom.CoInitialize()
                app, app_name = self._open_writer_app(win32com_client)
                document = self._open_or_get_document(app, self.document_path)
                paragraph_index = min(max(1, image.location.block_index), document.Paragraphs.Count)
                paragraph_range = document.Paragraphs(paragraph_index).Range
                bookmark_name = "__wps_image_locator__"
                try:
                    document.Activate()
                except Exception:
                    pass
                try:
                    if document.Bookmarks.Exists(bookmark_name):
                        document.Bookmarks(bookmark_name).Delete()
                except Exception:
                    pass
                exact_located = self._locate_by_office_collection(document, app, image.location)
                text_located = False
                if not exact_located and image.location.text_hint:
                    text_located = self._locate_by_text_hint(document, app, image.location)
                try:
                    if not exact_located and not text_located:
                        document.Bookmarks.Add(bookmark_name, paragraph_range)
                except Exception:
                    pass
                try:
                    if not exact_located and not text_located:
                        app.Selection.GoTo(What=-1, Name=bookmark_name)
                except Exception:
                    pass
                try:
                    if not exact_located and not text_located:
                        paragraph_range.Select()
                except Exception:
                    pass
                try:
                    if not exact_located and not text_located:
                        app.Selection.SetRange(paragraph_range.Start, paragraph_range.End)
                except Exception:
                    pass
                try:
                    if not exact_located and not text_located:
                        app.Selection.Collapse(1)
                except Exception:
                    pass
                try:
                    if not exact_located and not text_located:
                        document.ActiveWindow.ScrollIntoView(app.Selection.Range, True)
                except Exception:
                    pass
                try:
                    if not exact_located and not text_located:
                        app.ActiveWindow.ScrollIntoView(app.Selection.Range, True)
                except Exception:
                    pass
                try:
                    if not exact_located and not text_located:
                        app.Selection.Range.Select()
                except Exception:
                    pass
                app.Visible = True
                try:
                    app.Activate()
                except Exception:
                    pass
                try:
                    if document.Bookmarks.Exists(bookmark_name):
                        document.Bookmarks(bookmark_name).Delete()
                except Exception:
                    pass
                if exact_located:
                    collection_label = image.location.office_collection_index or image.location.occurrence_index + 1
                    return NavigationResult(True, f"已通过 {app_name} 按图片对象序号定位到第 {collection_label} 项。")
                if text_located:
                    return NavigationResult(True, f"已通过 {app_name} 按文本锚点尝试定位。")
                return NavigationResult(True, f"已通过 {app_name} 尝试定位到第 {paragraph_index} 段附近。")
            except Exception as exc:  # pragma: no cover - depends on local Office/WPS environment.
                return NavigationResult(False, f"自动定位失败：{exc}")
            finally:
                if pythoncom is not None:
                    try:
                        pythoncom.CoUninitialize()
                    except Exception:
                        pass

        try:
            os.startfile(self.document_path)  # type: ignore[attr-defined]
        except OSError as exc:
            return NavigationResult(False, f"无法打开文档：{exc}")

        return NavigationResult(False, f"已打开文档，请按位置提示手动定位：{image.location.display_text}")

    def _open_writer_app(self, win32com_client) -> tuple[object, str]:
        errors: list[str] = []
        for prog_id, display_name in (
            ("kwps.Application", "WPS Writer"),
            ("Word.Application", "Microsoft Word"),
        ):
            try:
                app = win32com_client.GetActiveObject(prog_id)
            except Exception as active_exc:
                try:
                    app = win32com_client.Dispatch(prog_id)
                except Exception as dispatch_exc:
                    errors.append(f"{display_name}: {active_exc}; {dispatch_exc}")
                    continue
            try:
                return app, display_name
            except Exception as exc:
                errors.append(f"{display_name}: {exc}")
        detail = "；".join(errors) if errors else "未检测到可用的 COM 应用。"
        raise DocumentError(f"未找到可用于定位的 WPS/Word 应用。{detail}")

    def _ensure_document_not_open_for_write(self) -> None:
        try:
            import pythoncom  # type: ignore
            import win32com.client  # type: ignore
        except Exception:
            return

        try:
            pythoncom.CoInitialize()
            for prog_id, display_name in (
                ("kwps.Application", "WPS Writer"),
                ("Word.Application", "Microsoft Word"),
            ):
                try:
                    app = win32com.client.GetActiveObject(prog_id)
                except Exception:
                    continue
                open_document = self._get_open_document(app, self.document_path)
                if open_document is not None:
                    raise DocumentError(f"检测到文档当前仍在 {display_name} 中打开，请先关闭该文档后再执行删除。")
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def _get_open_document(self, app, document_path: Path):
        resolved_path = document_path.resolve()
        normalized_target = os.path.normcase(os.path.normpath(str(resolved_path)))
        try:
            documents = app.Documents
        except Exception:
            return None

        for index in range(1, documents.Count + 1):
            current = documents.Item(index)
            full_name = os.path.normcase(os.path.normpath(str(getattr(current, "FullName", ""))))
            if full_name == normalized_target:
                return current
        return None

    def _open_or_get_document(self, app, document_path: Path):
        try:
            documents = app.Documents
        except Exception as exc:
            raise DocumentError(f"当前应用不支持文档自动化接口：{exc}") from exc

        current = self._get_open_document(app, document_path)
        if current is not None:
            try:
                current.Activate()
            except Exception:
                pass
            return current

        resolved_path = document_path.resolve()
        document = documents.Open(str(resolved_path))
        for _ in range(20):
            try:
                _ = document.Paragraphs.Count
                try:
                    document.Activate()
                except Exception:
                    pass
                return document
            except Exception:
                time.sleep(0.1)
        return document

    def _locate_by_text_hint(self, document, app, location: ImageLocation) -> bool:
        candidates = self._build_text_candidates(location)
        if not candidates:
            return False
        try:
            paragraph_index = location.block_index
            if paragraph_index is not None:
                for radius in (0, 2, 6, 12):
                    paragraph_range = self._paragraph_window_range(document, paragraph_index, radius)
                    if paragraph_range is not None and self._find_first_candidate(document, paragraph_range, candidates, app):
                        return True
            return self._find_first_candidate(document, document.Content, candidates[:3], app)
        except Exception:
            return False

    def _locate_by_office_collection(self, document, app, location: ImageLocation) -> bool:
        collection_name = location.office_collection
        collection_index = location.office_collection_index
        if not collection_name or collection_index is None:
            return False
        try:
            collection = getattr(document, collection_name)
            item = collection(collection_index)
        except Exception:
            try:
                collection = getattr(document, collection_name)
                item = collection.Item(collection_index)
            except Exception:
                return False

        target_range = None
        if collection_name == "InlineShapes":
            try:
                item.Select()
            except Exception:
                pass
            try:
                target_range = item.Range
            except Exception:
                target_range = None
        else:
            try:
                item.Select()
            except Exception:
                pass
            try:
                target_range = item.Anchor
            except Exception:
                target_range = None

        if target_range is None:
            try:
                target_range = app.Selection.Range
            except Exception:
                target_range = None
        if target_range is None:
            return False

        try:
            app.Selection.SetRange(target_range.Start, target_range.End)
        except Exception:
            pass
        try:
            document.ActiveWindow.ScrollIntoView(target_range, True)
        except Exception:
            pass
        try:
            app.ActiveWindow.ScrollIntoView(target_range, True)
        except Exception:
            pass
        try:
            target_range.Select()
        except Exception:
            pass
        return True

    def _paragraph_window_range(self, document, paragraph_index: int, radius: int):
        try:
            paragraph_count = document.Paragraphs.Count
            start_index = max(1, paragraph_index - radius)
            end_index = min(paragraph_count, paragraph_index + radius)
            start_range = document.Paragraphs(start_index).Range
            end_range = document.Paragraphs(end_index).Range
            return document.Range(start_range.Start, end_range.End)
        except Exception:
            return None

    def _find_first_candidate(self, document, search_scope, candidates: list[str], app) -> bool:
        try:
            base_start = search_scope.Start
            base_end = search_scope.End
        except Exception:
            return False

        for candidate in candidates:
            try:
                search_range = document.Range(base_start, base_end)
                finder = search_range.Find
                found = finder.Execute(
                    FindText=candidate,
                    Forward=True,
                    Wrap=0,
                    Format=False,
                    MatchCase=False,
                    MatchWholeWord=False,
                    MatchWildcards=False,
                    MatchSoundsLike=False,
                    MatchAllWordForms=False,
                )
            except Exception:
                continue
            if not found:
                continue
            try:
                search_range.Select()
            except Exception:
                pass
            try:
                app.Selection.SetRange(search_range.Start, search_range.End)
            except Exception:
                pass
            try:
                app.ActiveWindow.ScrollIntoView(search_range, True)
            except Exception:
                pass
            return True
        return False
