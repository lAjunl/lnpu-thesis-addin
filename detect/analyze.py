"""识别结果导出（结构化 JSON + 摘要）。

用法:
    .venv/Scripts/python.exe -m detect.analyze <docx> [-o out.json] [--use-style]
"""
from __future__ import annotations

import argparse
from pathlib import Path

from docx import Document

from .classify import analyze, is_heading_role


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("-o", "--out")
    ap.add_argument("--use-style", action="store_true")
    args = ap.parse_args()

    src = Path(args.docx)
    doc = Document(str(src))
    a = analyze(doc, source=src.name, use_style=args.use_style)

    print(f"文件: {src.name}   段落 {a.paragraph_count}")
    print(f"编号体系: {a.scheme}   {a.scheme_stats}")
    print(f"层级推断: {a.level_note}")
    print(f"\n识别摘要:")
    for k, v in sorted(a.summary()["by_role"].items(), key=lambda x: -x[1]):
        print(f"  {k:<18}{v}")

    heads = [d for d in a.detections if is_heading_role(d.role)]
    print(f"\n标题 {len(heads)} 个:")
    for d in heads[:60]:
        lv = f"{d.level} 级" if d.level else "未分级"
        conf = f"{d.level_confidence:.0%}" if d.level else "—"
        print(f"  #{d.index:<4d} {lv:<8} 把握 {conf:<5} {d.reasons[0][:28]:<30} | {a.features[d.index].text[:46]}")
    if len(heads) > 60:
        print(f"  … 其余 {len(heads)-60} 个")

    out = args.out or str(src.with_suffix(".analysis.json"))
    a.to_json(out)
    print(f"\n已写出 {out}")


if __name__ == "__main__":
    main()
