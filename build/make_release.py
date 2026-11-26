"""打分享包：绿色版 exe + README + 示例文档 + 截图。

文件名带版本号（版本出处：detect/__init__.py）。
"""
import shutil, sys, zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from detect import __version__          # noqa: E402

OUT = ROOT / "release"
OUT.mkdir(exist_ok=True)
zip_path = OUT / f"论文格式助手-v{__version__}.zip"

# 清空旧的exe目录副本
stage = OUT / "_stage"
if stage.exists():
    shutil.rmtree(stage)
(stage / "论文格式助手").mkdir(parents=True)

# 1) 绿色版程序
src_dir = ROOT / "dist" / "论文格式助手"
dst_dir = stage / "论文格式助手"
n = 0
for p in src_dir.rglob("*"):
    if p.is_file():
        rel = p.relative_to(src_dir)
        (dst_dir / rel).parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(p, dst_dir / rel)
        n += 1
print("复制程序文件:", n)

# 2) 文档与示例
shutil.copy2(ROOT / "README.md", stage / "README.md")
changelog = ROOT / "CHANGELOG.md"
if changelog.is_file():
    shutil.copy2(changelog, stage / "CHANGELOG.md")
(stage / "docs").mkdir(exist_ok=True)
shutil.copy2(ROOT / "docs" / "screenshot.png", stage / "docs" / "screenshot.png")
sample = ROOT / "assets" / "sample-filled.docx"
if sample.is_file():
    shutil.copy2(sample, stage / "示例文档.docx")

# 3) 打包
if zip_path.exists():
    zip_path.unlink()
legacy = OUT / "论文格式助手.zip"      # 旧的无版本号文件名
if legacy.exists():
    legacy.unlink()
with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
    for p in sorted(stage.rglob("*")):
        if p.is_file():
            z.write(p, p.relative_to(stage).as_posix())
print("包:", zip_path, f"{zip_path.stat().st_size/1024/1024:.1f} MB")
shutil.rmtree(stage)
