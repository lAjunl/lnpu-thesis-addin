# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置：生成「论文格式助手」Windows 程序。

用法:
    .venv/Scripts/python.exe -m PyInstaller --noconfirm --clean lnpu.spec
产物:
    dist/论文格式助手/论文格式助手.exe
"""

block_cipher = None

# 只带界面需要的 Qt 模块，砍掉用不到的（能省 100MB+）
hidden = [
    "PySide6.QtSvg",
    "PySide6.sip",
    "pythoncom",
    "pywintypes",
    "win32com",
    "win32com.client",
    "PIL",
    "PIL.Image",
    "detect",
    "detect.ui",
    "detect.ui.app",
    "detect.ui.main_window",
    "detect.ui.dialogs",
    "detect.ui.render",
    "detect.ui.settings",
    "detect.ui.preview",
    "detect.ui.pageview",
    "detect.ui.widgets",
    "detect.ui.word_preview",
    "detect.ui.assets",
    "detect.features",
    "detect.classify",
    "detect.reflow",
    "detect.sections",
    "detect.numbering",
    "detect.recognize",
    "detect.check",
    "detect.preview_render",
]

excluded = [
    "PySide6.QtQml", "PySide6.QtQuick", "PySide6.QtQuickWidgets",
    "PySide6.Qt3DCore", "PySide6.Qt3DRender", "PySide6.QtMultimedia",
    "PySide6.QtMultimediaWidgets", "PySide6.QtNetwork",
    "PySide6.QtWebEngineCore", "PySide6.QtWebEngineWidgets", "PySide6.QtWebEngineQuick",
    "PySide6.QtTest", "PySide6.QtDesigner", "PySide6.QtSql",
    "PySide6.QtBluetooth", "PySide6.QtNfc", "PySide6.QtPositioning",
    "PySide6.QtSerialPort", "PySide6.QtCharts", "PySide6.QtDataVisualization",
    "PySide6.QtOpenGL", "PySide6.QtOpenGLWidgets", "PySide6.QtPdf",
    "PySide6.QtPdfWidgets", "PySide6.QtHttpServer", "PySide6.QtTextToSpeech",
    "PySide6.QtSensors", "PySide6.QtPrintSupport", "PySide6.QtHelp",
    "PySide6.QtOpenGL", "tkinter", "matplotlib", "numpy",
]

a = Analysis(
    ["run_ui.py"],
    pathex=[],
    binaries=[],
    datas=[("assets/icons/icon-256.png", "assets/icons")],
    hiddenimports=hidden,
    hookspath=[],
    runtime_hooks=[],
    excludes=excluded,
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="论文格式助手",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    icon="assets/icons/app.ico",
    manifest_filename="dpi.manifest",
    # 把版本写进 exe 的 Windows 文件属性（资源管理器 → 属性 → 详细信息）
    version="build/version.txt",
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    name="论文格式助手",
)
