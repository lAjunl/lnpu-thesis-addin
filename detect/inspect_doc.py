"""文档结构勘察：打印一份 docx 的段落、样式、格式特征，用于设计识别规则。

用法:
    .venv/Scripts/python.exe detect/inspect_doc.py <docx路径>
"""
from __future__ import annotations

import sys
from collections import Counter
from pathlib import Path

from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_LINE_SPACING


def east_asia_font(run) -> str:
    rPr = run._r.find(qn("w:rPr"))
    if rPr is None:
        return ""
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        return ""
    return rFonts.get(qn("w:eastAsia")) or ""


def outline_level(p) -> int | None:
    """python-docx 的 ParagraphFormat 未暴露大纲级别，直接读 w:outlineLvl。"""
    pPr = p._p.find(qn("w:pPr"))
    if pPr is None:
        return None
    el = pPr.find(qn("w:outlineLvl"))
    if el is None:
        return None
    try:
        return int(el.get(qn("w:val")))
    except (TypeError, ValueError):
        return None


def para_features(p) -> dict:
    pf = p.paragraph_format
    runs = [r for r in p.runs if r.text.strip()]
    sizes = [r.font.size.pt for r in runs if r.font.size is not None]
    bolds = [bool(r.bold) for r in runs]
    return {
        "text": p.text.strip(),
        "style": p.style.name if p.style is not None else "",
        "align": str(pf.alignment),
        "size": max(sizes) if sizes else None,
        "bold": any(bolds),
        "outline": outline_level(p),
        "line_rule": str(pf.line_spacing_rule),
        "line": pf.line_spacing,
        "space_before": pf.space_before.pt if pf.space_before else None,
        "space_after": pf.space_after.pt if pf.space_after else None,
        "first_indent": pf.first_line_indent.pt if pf.first_line_indent else None,
        "east_asia": east_asia_font(runs[0]) if runs else "",
        "ascii": runs[0].font.name if runs else "",
        "in_table": False,
        "has_drawing": p._p.find(".//" + qn("w:drawing")) is not None
        or p._p.find(".//" + qn("w:pict")) is not None,
    }


def main() -> None:
    path = Path(sys.argv[1])
    doc = Document(str(path))

    print(f"文件: {path.name}  ({path.stat().st_size / 1024:.0f} KB)")
    print(f"段落总数: {len(doc.paragraphs)}  表格: {len(doc.tables)}  "
          f"节: {len(doc.sections)}  内嵌图: {len(doc.inline_shapes)}")

    # 样式使用统计
    styles = Counter(p.style.name for p in doc.paragraphs if p.style is not None)
    print("\n=== 样式使用 ===")
    for name, n in styles.most_common(20):
        print(f"  {name!r}: {n}")

    feats = [para_features(p) for p in doc.paragraphs]

    print("\n=== 大纲级别分布 ===")
    oc = Counter(str(f["outline"]) for f in feats)
    for k, v in oc.most_common():
        print(f"  outline={k}: {v}")

    print("\n=== 字号分布 ===")
    sc = Counter(f["size"] for f in feats)
    for k, v in sorted(sc.items(), key=lambda x: (x[0] is None, x[0])):
        print(f"  {k}pt: {v}")

    print("\n=== 行距规则分布 ===")
    lc = Counter(f["line_rule"] for f in feats)
    for k, v in lc.most_common():
        print(f"  {k}: {v}")

    print("\n=== 前 60 段（序号 | 样式 | 字号 | 加粗 | 对齐 | outline | 文本）===")
    for i, f in enumerate(feats[:60]):
        t = f["text"][:40].replace("\n", " ")
        print(
            f"#{i:4d} [{f['style'][:12]:12s}] {str(f['size']):>5} b={int(f['bold'])} "
            f"al={str(f['align'])[:12]:12s} ol={str(f['outline']):>4} | {t}"
        )

    print("\n=== 含图形的段落 ===")
    dr = [i for i, f in enumerate(feats) if f["has_drawing"]]
    print(f"  共 {len(dr)} 段: {dr[:20]}")

    print("\n=== 以 图/表 开头的段落 ===")
    for i, f in enumerate(feats):
        t = f["text"]
        if t and t[0] in "图表" and len(t) < 60:
            print(f"  #{i:4d} [{f['style'][:10]:10s}] {str(f['size']):>5} "
                  f"al={str(f['align'])[:10]:10s} | {t[:50]}")

    print("\n=== 以数字编号开头的段落（前 40）===")
    import re

    n = 0
    for i, f in enumerate(feats):
        if re.match(r"^\s*\d+(\.\d+)*\s", f["text"]):
            print(f"  #{i:4d} [{f['style'][:10]:10s}] {str(f['size']):>5} "
                  f"b={int(f['bold'])} al={str(f['align'])[:10]:10s} | {f['text'][:46]}")
            n += 1
            if n >= 40:
                break


if __name__ == "__main__":
    main()
