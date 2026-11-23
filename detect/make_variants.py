"""构造"无格式稿"变体，用于压力测试结构识别器。

用法:
    .venv/Scripts/python.exe detect/make_variants.py <docx路径>

产出（build/ 下）:
    variant_a_plain.docx   彻底无格式：所有样式引用与格式属性清空
    variant_b_flat.docx    样式拍平：把生效格式写成直接格式，再移除样式引用
    ground_truth.json      每段的语义角色（来自原文件的样式 + 文献区判定）

变体 B 模拟"作者手动设了字号/加粗，但从来没建过样式"的真实手写稿；
变体 A 是最坏情况：全文一个字号，只剩文本和编号。
"""
from __future__ import annotations

import json
import shutil
import tempfile
import sys
from pathlib import Path

from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

from .features import (
    extract_features,
    style_chain,
    _size_pt,
    _bold_of,
    _east_asia_of,
    _ascii_of,
    _align_of,
    _line_of,
)
from .evaluate import semantic_truth

FMT_TAGS_PPR = ["spacing", "jc", "ind", "outlineLvl", "pStyle", "keepNext", "keepLines"]
FMT_TAGS_RPR = ["rFonts", "sz", "szCs", "b", "bCs", "i", "iCs", "color", "u", "rStyle"]


def _el(parent, tag: str):
    if parent is None:
        return None
    return parent.find(qn("w:" + tag))


def _drop_children(el, tags) -> None:
    if el is None:
        return
    for tag in tags:
        for node in el.findall(qn("w:" + tag)):
            el.remove(node)


def _fresh(src_path: Path) -> Document:
    """复制到临时文件再打开。

    不能用 copy.deepcopy(Document)：python-docx 的 Document/Part 持有 lxml
    元素，deepcopy 不会真正复制部件树，改动不会反映到 save() 的结果里
    （实测变体文件和原文一模一样）。
    """
    tmp = Path(tempfile.mkdtemp(prefix="variant_")) / "copy.docx"
    shutil.copyfile(src_path, tmp)
    return Document(str(tmp)), tmp


def make_plain(src_path: Path) -> Document:
    """变体 A：清空所有格式，只留文本。"""
    doc, _ = _fresh(src_path)
    for p in doc.paragraphs:
        p_el = p._p
        pPr = _el(p_el, "pPr")
        _drop_children(pPr, FMT_TAGS_PPR)
        for r in p.runs:
            rPr = _el(r._r, "rPr")
            _drop_children(rPr, FMT_TAGS_RPR)
    # 页眉页脚同样清空
    for sec in doc.sections:
        for part in (sec.header, sec.footer):
            for p in part.paragraphs:
                pPr = _el(p._p, "pPr")
                _drop_children(pPr, FMT_TAGS_PPR)
                for r in p.runs:
                    _drop_children(_el(r._r, "rPr"), FMT_TAGS_RPR)
    return doc


def make_flat(src_path: Path) -> Document:
    """变体 B：把生效格式写成直接格式，再移除样式引用。

    解析用 features.py 的独立函数，不经过分类器，避免循环依赖。
    """
    doc, _ = _fresh(src_path)
    defaults = doc.element.find(".//" + qn("w:docDefaults"))
    defaults_rpr = _el(defaults, "rPr")

    for p in doc.paragraphs:
        style = p.style
        chain = style_chain(style) if style is not None else []
        pPr = _el(p._p, "pPr")

        size = _size_pt(p, pPr, chain, defaults_rpr)
        bold = _bold_of(p, pPr, chain)
        ea = _east_asia_of(p, pPr, chain)
        ascii_f = _ascii_of(p, pPr, chain)
        align = _align_of(pPr, chain)
        rule, value = _line_of(pPr, chain)

        # 先把段落级格式清空（避免段落 pPr 覆盖 run 直接格式）
        _drop_children(pPr, FMT_TAGS_PPR)

        # 行距/对齐写成段落直接格式
        pf = p.paragraph_format
        if rule == "exactly" and value is not None:
            pf.line_spacing_rule = 4  # EXACTLY
            pf.line_spacing = Pt(value)
        elif rule == "multiple" and value is not None:
            pf.line_spacing_rule = 5  # MULTIPLE
            pf.line_spacing = value
        if align:
            pf.alignment = {
                "left": 0,
                "center": 1,
                "right": 2,
                "both": 3,
            }.get(align)

        # run 级格式
        for r in p.runs:
            if not r.text.strip():
                continue
            if size is not None:
                r.font.size = Pt(size)
            if bold is not None:
                r.font.bold = bold
            if ascii_f:
                r.font.name = ascii_f
            if ea:
                rPr = r._r.get_or_add_rPr()
                rFonts = rPr.find(qn("w:rFonts"))
                if rFonts is None:
                    from docx.oxml import OxmlElement

                    rFonts = OxmlElement("w:rFonts")
                    rPr.insert(0, rFonts)
                rFonts.set(qn("w:eastAsia"), ea)

        # 移除样式引用（保留 Normal）
        p.style = doc.styles["Normal"]

    return doc


def main() -> None:
    src_path = Path(sys.argv[1])
    out_dir = Path(__file__).resolve().parent.parent / "build"
    out_dir.mkdir(exist_ok=True)

    src = Document(str(src_path))
    feats = extract_features(src)
    truth, meta = semantic_truth(feats)
    (out_dir / "ground_truth.json").write_text(
        json.dumps(
            {
                "source": src_path.name,
                "ref_start": meta["ref_start"],
                "paragraphs": [
                    {"index": f.index, "role": r, "style": f.style_name, "text": f.text[:80]}
                    for f, r in zip(feats, truth)
                ],
            },
            ensure_ascii=False,
            indent=1,
        ),
        encoding="utf-8",
    )
    print(f"ground truth 已写出（文献区起点 {meta['ref_start']}）")

    a = make_plain(src_path)
    (out_dir / "variant_a_plain.docx").write_bytes(_save(a))
    print("变体 A（彻底无格式）已生成")

    b = make_flat(src_path)
    (out_dir / "variant_b_flat.docx").write_bytes(_save(b))
    print("变体 B（样式拍平为直接格式）已生成")


def _save(doc: Document) -> bytes:
    import io

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


if __name__ == "__main__":
    main()
