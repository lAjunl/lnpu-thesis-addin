"""把 OLE 公式（MathType 等）的 EMF 矢量图转成 PNG。

Qt 的 QImage 不支持 EMF，所以预览里 MathType 公式全是空白。
Windows 自带的 GDI+ 能渲染 EMF，这里用 ctypes 调它，
转换结果按 (文件内容哈希, 目标宽度) 缓存到 %TEMP%。
"""
from __future__ import annotations

import ctypes
import hashlib
import os
import tempfile
from pathlib import Path

PNG_CLSID = b"\x06\xf4\x7c\x55\x04\x1a\xd3\x11\x9a\x73\x00\x00\xf8\x1e\xf3\x2e"

_cache: dict[tuple, Path] = {}


class _StartupInput(ctypes.Structure):
    _fields_ = [("GdiplusVersion", ctypes.c_uint32),
                ("DebugEventCallback", ctypes.c_void_p),
                ("SuppressBackgroundThread", ctypes.c_int),
                ("SuppressExternalCodecs", ctypes.c_int)]


def _gdiplus():
    return ctypes.WinDLL("gdiplus.dll")


def emf_to_png(blob: bytes, target_w: int = 700) -> Path | None:
    if not blob:
        return None
    key = (hashlib.md5(blob).hexdigest(), int(target_w))
    hit = _cache.get(key)
    if hit is not None and hit.exists():
        return hit
    src = Path(tempfile.mkdtemp(prefix="emf_")) / "src.emf"
    src.write_bytes(blob)
    dst = src.with_suffix(".png")
    try:
        if not _convert(str(src), str(dst), int(target_w)):
            return None
    except Exception:
        return None
    if len(_cache) > 64:
        _cache.clear()
    _cache[key] = dst
    return dst


def _convert(src: str, dst: str, target_w: int) -> bool:
    g = _gdiplus()
    token = ctypes.c_void_p()
    si = _StartupInput(1, None, 0, 0)
    if g.GdiplusStartup(ctypes.byref(token), ctypes.byref(si), None) != 0:
        return False
    try:
        img = ctypes.c_void_p()
        if g.GdipLoadImageFromFile(src, ctypes.byref(img)) != 0 or not img:
            return False
        try:
            w = ctypes.c_uint()
            h = ctypes.c_uint()
            g.GdipGetImageWidth(img, ctypes.byref(w))
            g.GdipGetImageHeight(img, ctypes.byref(h))
            sw, sh = max(1, w.value), max(1, h.value)
            # 允许放大：EMF 的像素尺寸往往远小于"按磅显示所需"的尺寸
            # （实测 MathType 导出的 EMF 是 32 px/磅，而按磅显示只要 1 px/磅）。
            # 旧代码写死 min(1.0, …) 永不放大，公式的细笔画被拉虚。
            scale = (target_w / float(sw)) if target_w else 1.0
            scale = max(0.05, min(8.0, scale))
            dw, dh = max(1, int(sw * scale)), max(1, int(sh * scale))

            bmp = ctypes.c_void_p()
            if g.GdipCreateBitmapFromScan0(dw, dh, 0, 0x0026200A, None,
                                           ctypes.byref(bmp)) != 0:
                return False
            try:
                gr = ctypes.c_void_p()
                if g.GdipGetImageGraphicsContext(bmp, ctypes.byref(gr)) != 0:
                    return False
                try:
                    g.GdipSetCompositingMode(gr, 1)          # CompositingModeSourceOver
                    g.GdipSetInterpolationMode(gr, 7)       # HighQualityBicubic
                    g.GdipDrawImageRectI(gr, img, 0, 0, dw, dh)
                finally:
                    g.GdipDeleteGraphics(gr)
                return g.GdipSaveImageToFile(bmp, dst,
                                             PNG_CLSID, None) == 0
            finally:
                g.GdipDisposeImage(bmp)
        finally:
            g.GdipDisposeImage(img)
    finally:
        g.GdiplusShutdown(token)
