"""真实平台截图：用系统真实字体渲染，导出主窗口和左右两栏的图片。

用法:
    .venv/Scripts/python.exe build/shot.py <docx路径> [输出前缀]

注意：不设 QT_QPA_PLATFORM=offscreen，这样拿到的是真实字体和真实分页。
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from PySide6.QtCore import Qt, QCoreApplication
from PySide6.QtWidgets import QApplication, QFileDialog

from detect.ui.main_window import MainWindow
from detect.ui.style import apply


def main() -> int:
    src = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "assets" / "sample-filled.docx"
    prefix = sys.argv[2] if len(sys.argv) > 2 else "shot"

    QCoreApplication.setAttribute(Qt.AA_ShareOpenGLContexts, True)
    try:
        QApplication.setHighDpiScaleFactorRoundingPolicy(
            Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
        )
    except Exception:
        pass
    app = QApplication([])
    apply(app)

    win = MainWindow()
    win.resize(1500, 900)
    win.show()
    QFileDialog.getOpenFileName = staticmethod(lambda *a, **k: (str(src), ""))
    win.open_file()
    app.processEvents()

    out = ROOT / "build"
    win.grab().save(str(out / f"{prefix}_main.png"))
    win.before_view.grab().save(str(out / f"{prefix}_before.png"))
    win.after_view.grab().save(str(out / f"{prefix}_after.png"))

    bd, ad = win.before_view.document(), win.after_view.document()
    print(f"src={src.name}")
    print(f"pages: before={win.before_view.page_count()} after={win.after_view.page_count()}")
    print(f"labels: {win.after_view._labels[:8]} ... {win.after_view._labels[-3:]}")
    print(f"chars: before={bd.characterCount()} after={ad.characterCount()}")
    win.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
