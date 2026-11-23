"""把每个二级菜单逐个打开并截图，用于以用户视角检查易用性。

用法:
    .venv/Scripts/python.exe build/shot_dialogs.py <docx路径>
产出:
    build/dlg_<key>.png
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from PySide6.QtCore import Qt, QCoreApplication
from PySide6.QtWidgets import QApplication, QFileDialog

from detect.ui.dialogs import (
    AboutDialog,
    IndentDialog,
    LayoutDialog,
    NumberingDialog,
    PageDialog,
    RoleDialog,
    RolesDialog,
)
from detect.ui.main_window import MainWindow
from detect.ui.settings import EDITABLE_ROLES, ROLE_LABELS
from detect.ui.style import apply

OUT = ROOT / "build"


def build_dialog(win, key):
    s = win.settings
    if key in EDITABLE_ROLES:
        return RoleDialog(ROLE_LABELS.get(key, key), s.roles[key], win)
    if key == "page":
        return PageDialog(s.page, win)
    if key in ("indent", "not_indent"):
        return IndentDialog(s, win, focus_not_indent=(key == "not_indent"))
    if key in ("num_figure", "num_table"):
        from detect.ui.dialogs import CaptionNumberDialog
        from detect.ui.main_window import NUM_KIND
        return CaptionNumberDialog(s, win, kind=NUM_KIND[key])
    if key == "num_heading":
        from detect.ui.main_window import NUM_KIND
        return NumberingDialog(s, win, kind=NUM_KIND[key])
    if key == "table_style":
        from detect.ui.dialogs import TableDialog
        return TableDialog(s, win)
    if key in ("layout_section", "layout_pagenum", "layout_header"):
        from detect.ui.main_window import LAYOUT_KIND
        return LayoutDialog(s, win, kind=LAYOUT_KIND[key])
    if key == "roles":
        return RolesDialog(win.analysis, win.doc, win)
    if key == "about":
        return AboutDialog(win)
    return None


def main() -> int:
    src = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "assets" / "sample-filled.docx"
    QCoreApplication.setAttribute(Qt.AA_ShareOpenGLContexts, True)
    app = QApplication([])
    apply(app)

    win = MainWindow()
    win.resize(1500, 900)
    QFileDialog.getOpenFileName = staticmethod(lambda *a, **k: (str(src), ""))
    win.open_file()
    app.processEvents()
    win.hide()          # 只留卡片在屏幕上，截图更干净

    keys = ["page", "indent", "not_indent", "table_style"] + EDITABLE_ROLES + \
           ["num_heading", "num_figure", "num_table",
            "layout_section", "layout_pagenum", "layout_header",
            "roles", "about"]
    for key in keys:
        dlg = build_dialog(win, key)
        if dlg is None:
            continue
        dlg.show()
        app.processEvents()
        app.processEvents()
        name = f"dlg_{key}.png"
        dlg.grab().save(str(OUT / name))
        print("saved", name, dlg.geometry().width(), "x", dlg.geometry().height())
        dlg.close()
    win.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
