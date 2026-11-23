"""用 Word 实时预览：原稿只读打开，排版后另存并刷新。

为什么不内嵌：Qt 渲染不出 Word 的真实排版（字体、行距、分页、页眉页脚），
已经改了三版都不对。Word 自己才是唯一真实的渲染器。
"""
from __future__ import annotations

import shutil
import tempfile
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QMessageBox, QWidget

ORIGINAL_NAME = "原稿_预览.docx"
RENDERED_NAME = "排版预览.docx"


class WordPreview:
    """管理临时目录里的两份文档，并用 Word 并排打开。"""

    def __init__(self, source: Path, parent: QWidget | None = None):
        self.source = Path(source)
        self.parent = parent
        self.dir = Path(tempfile.mkdtemp(prefix="thesis_preview_"))
        self.original = self.dir / ORIGINAL_NAME
        self.rendered = self.dir / RENDERED_NAME
        self._word = None
        self._doc_o = None
        self._doc_r = None
        self._ok = False
        self._we_started = False

    # ---------------------------------------------------------------- #
    @property
    def ready(self) -> bool:
        return self._ok

    @property
    def dir_path(self) -> Path:
        return self.dir

    def _app(self):
        import win32com.client

        try:
            return win32com.client.GetActiveObject("Word.Application")
        except Exception:
            self._we_started = True
            return win32com.client.Dispatch("Word.Application")

    def _docs_to_close(self):
        """收集需要关闭的文档对象。

        不能只依赖保存下来的 self._doc_o/_doc_r：刷新预览时旧引用会失效，
        Close 静默抛异常后窗口就留在那儿了（实测会留下一个"阅读模式"窗口）。
        所以再按文件名在当前 Word 实例里兜底找一遍。
        """
        found = []
        for doc in (self._doc_o, self._doc_r):
            if doc is not None:
                found.append(doc)
        try:
            names = {ORIGINAL_NAME, RENDERED_NAME}
            for d in list(self._word.Documents):
                try:
                    if d.Name in names and d not in found:
                        found.append(d)
                except Exception:
                    continue
        except Exception:
            pass
        return found

    # ---------------------------------------------------------------- #
    def open(self, render_fn) -> bool:
        """生成排版预览，并把两份文档并排打开。"""
        try:
            shutil.copyfile(self.source, self.original)
            render_fn(self.rendered)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self.parent, "生成预览失败", str(exc))
            return False

        try:
            word = self._app()
            word.Visible = True
            word.DisplayAlerts = 0
            self._doc_o = word.Documents.Open(str(self.original), ReadOnly=True)
            self._doc_r = word.Documents.Open(str(self.rendered), ReadOnly=False)
            try:
                word.Windows.Arrange(3)   # 垂直并排
            except Exception:
                pass
            self._word = word
            self._ok = True
            return True
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(
                self.parent, "打开 Word 失败",
                f"{exc}\n\n请确认已安装 Microsoft Word。",
            )
            return False

    def refresh(self, render_fn) -> bool:
        """重新生成排版预览并在 Word 里重开（原稿不动）。

        顺序很重要：必须先关掉 Word 里的预览文档再写文件，
        否则 Windows 的文件锁会让保存报 PermissionError。
        """
        if not self._ok:
            return False
        # 先关掉 Word 里的预览文档再写文件，否则文件锁会让保存报 PermissionError。
        # 按文件名兜底，避免旧 COM 引用失效后关不掉。
        closed = False
        for doc in (self._doc_r,):
            try:
                if doc is not None:
                    doc.Close(False)
                    closed = True
            except Exception:
                pass
        if not closed:
            try:
                for d in list(self._word.Documents):
                    if d.Name == RENDERED_NAME:
                        d.Close(False)
            except Exception:
                pass
        self._doc_r = None

        try:
            render_fn(self.rendered)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self.parent, "生成预览失败", str(exc))
            return False
        try:
            self._doc_r = self._word.Documents.Open(str(self.rendered), ReadOnly=False)
            try:
                self._doc_r.Activate()
            except Exception:
                pass
            try:
                self._word.Windows.Arrange(3)
            except Exception:
                pass
            return True
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self.parent, "刷新预览失败", str(exc))
            return False

    # ---------------------------------------------------------------- #
    def close(self) -> None:
        """只关闭预览用的两份文档，不退出用户可能正在用的 Word。"""
        for doc in self._docs_to_close():
            try:
                doc.Close(False)
            except Exception:
                pass
        self._doc_o = self._doc_r = None
        self._ok = False

    def cleanup(self) -> None:
        self.close()
        # 如果这个 Word 是我们为了预览才拉起来的，且已经没有文档了，就退出它，
        # 免得在后台留一个看不见的 WINWORD.EXE。用户自己开着的 Word 不动。
        if self._we_started and self._word is not None:
            try:
                if self._word.Documents.Count == 0:
                    self._word.Quit()
            except Exception:
                pass
            self._word = None
        try:
            shutil.rmtree(self.dir, ignore_errors=True)
        except Exception:
            pass
