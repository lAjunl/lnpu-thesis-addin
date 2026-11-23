"""启动桌面界面。

用法:
    .venv/Scripts/python.exe -m detect.ui.app
或:
    .venv/Scripts/python.exe run_ui.py
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

# 高 DPI：必须在 QApplication 之前设置环境变量，否则 Windows 会做位图拉伸，整个界面发糊
os.environ.setdefault("QT_AUTO_SCREEN_SCALE_FACTOR", "1")
os.environ.setdefault("QT_ENABLE_HIGHDPI_SCALING", "1")
os.environ.setdefault("QT_SCALE_FACTOR_ROUNDING_POLICY", "PassThrough")

from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication

from .assets import app_icon_path
from .main_window import MainWindow
from .style import apply


def main() -> int:
    try:
        QApplication.setHighDpiScaleFactorRoundingPolicy(
            Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
        )
    except Exception:
        pass

    app = QApplication(sys.argv)
    for attr in (
        "AA_EnableHighDpiScaling",
        "AA_UseHighDpiPixmaps",
        "AA_ShareOpenGLContexts",
    ):
        try:
            app.setAttribute(getattr(Qt, attr), True)
        except Exception:
            pass

    app.setApplicationName("论文格式助手")
    icon = QIcon(app_icon_path())
    app.setWindowIcon(icon)
    apply(app)

    win = MainWindow()
    win.setWindowIcon(icon)
    win.show()

    # 命令行直接跟一个 docx 路径：双击/右键打开时自动载入
    for arg in sys.argv[1:]:
        p = Path(arg)
        if p.suffix.lower() == ".docx" and p.is_file():
            win.open_path(p)
            break

    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
