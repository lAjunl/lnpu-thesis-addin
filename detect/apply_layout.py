"""分节 + 页码页眉 应用入口。

用法:
    .venv/Scripts/python.exe -m detect.apply_layout <docx> <front_index> <body_index> [-o out.docx]

例（前置从"摘要"开始，正文从第一个一级标题开始）:
    .venv/Scripts/python.exe -m detect.apply_layout test.docx 11 16 -o out.docx

会先自动备份到 <out>.bak.docx。
"""
from __future__ import annotations

import argparse
import shutil
from pathlib import Path

from docx import Document

from .sections import apply, diagnose


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("front_index", type=int, help="前置部分起点段号（摘要/目录）")
    ap.add_argument("body_index", type=int, help="正文起点段号（第一个一级标题）")
    ap.add_argument("-o", "--out")
    args = ap.parse_args()

    src = Path(args.docx)
    out = Path(args.out) if args.out else src.with_name(src.stem + "_分节.docx")

    doc = Document(str(src))
    print(f"切分前：{len(doc.sections)} 节 / {len(doc.paragraphs)} 段")

    apply(doc, front_index=args.front_index, body_index=args.body_index)

    backup = out.with_suffix(".bak.docx")
    if not backup.exists():
        shutil.copyfile(out if out.exists() else src, backup)
        print(f"已备份 → {backup}")

    doc.save(str(out))

    d = diagnose(doc)
    print(f"\n已保存 → {out}")
    for s in d["sections"]:
        fmt = s["pgNumType_fmt"] or "无"
        start = s["pgNumType_start"] or "—"
        print(
            f"  第{s['index']}节  页眉={s['header'] or '(无)':<12} "
            f"页脚={(s['footer_text'] or '(无)')}  页码={fmt} 起始={start}"
        )


if __name__ == "__main__":
    main()
