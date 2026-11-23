"""界面资源路径：开发期读项目目录，打包后读 PyInstaller 解包目录。"""
from __future__ import annotations

import sys
from pathlib import Path


def asset_path(name: str) -> str:
    base = getattr(sys, "_MEIPASS", None)
    if base:
        return str(Path(base) / "assets" / "icons" / name)
    return str(Path(__file__).resolve().parent.parent.parent / "assets" / "icons" / name)


def app_icon_path() -> str:
    return asset_path("icon-256.png")
