"""上下标：保留原文的角标，并把"CO2""m2"这类纯文本里的数字转成角标。

为什么要做：科技论文里化学式（CO2、H2S、CH4）和面积/体积单位（m2、cm3）
的数字必须是下标/上标，但很多稿子是从网页或 Word 里拷出来的纯文本，
数字是正常大小。工具得能识别并修正。

保守策略：只认"元素符号 + 数字"的化学式形态和"长度单位 + 数字"，
并有排除表（A4、B5 这类不是化学式），避免把 "Figure 2" 改坏。
"""
from __future__ import annotations

import copy
import re

# 常见元素符号（覆盖石油/化工/环境类论文够用）
ELEMENTS = {
    "H", "He", "Li", "Be", "B", "C", "N", "O", "F", "Ne", "Na", "Mg", "Al",
    "Si", "P", "S", "Cl", "Ar", "K", "Ca", "Sc", "Ti", "V", "Cr", "Mn", "Fe",
    "Co", "Ni", "Cu", "Zn", "Ga", "Ge", "As", "Se", "Br", "Kr", "Rb", "Sr",
    "Y", "Zr", "Nb", "Mo", "Tc", "Ru", "Rh", "Pd", "Ag", "Cd", "In", "Sn",
    "Sb", "Te", "I", "Xe", "Cs", "Ba", "La", "Ce", "Nd", "Sm", "Eu", "Gd",
    "Tb", "Dy", "Ho", "Er", "Tm", "Yb", "Lu", "Hf", "Ta", "W", "Re", "Os",
    "Ir", "Pt", "Au", "Hg", "Tl", "Pb", "Bi", "Th", "Pa", "U", "Np", "Pu",
}

# 不是化学式的"字母+数字"，不要动
EXCLUDE = {
    "A4", "A3", "A5", "B4", "B5", "C4", "C5", "D4", "D5", "E4", "E5",
    "F4", "F5", "G4", "G5", "H4", "H5", "K4", "K5", "M4", "M5", "N4",
    "N5", "P4", "P5", "R4", "R5", "S4", "S5", "T4", "T5", "U4", "U5",
    "V4", "V5", "W4", "W5", "X4", "X5", "Y4", "Y5", "Z4", "Z5", "B2",
    "C3", "D3", "E3", "G3", "H3", "K3", "M3", "N3", "P3", "R3", "S3",
    "T3", "U3", "V3", "W3", "X3", "Y3", "Z3",
}

# 化学式：大写开头，由"元素符号+数字"重复组成，且整体含数字
RE_FORMULA = re.compile(r"\b((?:[A-Z][a-z]?\d*)+)\b")
# 长度/面积/体积单位：m2、cm3、km2、mm2 …
RE_UNIT = re.compile(r"\b(cm|mm|km|dm|um|nm|m)(\d+)\b")

RE_DIGITS = re.compile(r"\d+")


def _formula_sub(m: re.Match) -> str:
    """占位：实际拆分在 split_marks 里做。"""
    return m.group(1)


def split_marks(text: str):
    """把文本切成 [(片段, 角标)]，角标 0=正常 1=上标 -1=下标。"""
    pieces: list[list] = []
    pos = 0

    def push(t, mark=0):
        if t:
            pieces.append([t, mark])

    for m in RE_FORMULA.finditer(text):
        token = m.group(1)
        if token in EXCLUDE or not any(c.isdigit() for c in token):
            continue
        push(text[pos:m.start()])
        i = 0
        while i < len(token):
            ch = token[i]
            if ch.isdigit():
                j = i
                while j < len(token) and token[j].isdigit():
                    j += 1
                push(token[i:j], -1)
                i = j
            else:
                push(ch, 0)
                i += 1
        pos = m.end()
    push(text[pos:])

    # 单位上标（m2、cm3）：只处理没被化学式吃掉的片段
    merged: list[list] = []
    for t, mark in pieces:
        if mark == 0 and t:
            last = 0
            for m in RE_UNIT.finditer(t):
                if m.start() > 0 and t[m.start() - 1].isalnum():
                    continue
                merged.append([t[last:m.start()], 0])
                merged.append([m.group(1), 0])
                merged.append([m.group(2), 1])
                last = m.end()
            merged.append([t[last:], 0])
        else:
            merged.append([t, mark])
    return [(t, m) for t, m in merged if t]


def _set_vert(run, mark: int) -> None:
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    rPr = run._r.get_or_add_rPr()
    for el in rPr.findall(qn("w:vertAlign")):
        rPr.remove(el)
    if mark:
        va = OxmlElement("w:vertAlign")
        va.set(qn("w:val"), "superscript" if mark > 0 else "subscript")
        rPr.append(va)


def apply_marks_to_paragraph(paragraph) -> int:
    """把一个段落里的"CO2""m2"拆成角标 run，保留原格式。返回新增的 run 数。"""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    total = 0
    for run in list(paragraph.runs):
        if run.font.superscript or run.font.subscript:
            continue                      # 原文已有角标，不动
        pieces = split_marks(run.text)
        if len(pieces) <= 1:
            continue
        rPr = run._r.find(qn("w:rPr"))
        anchor = run._r
        made = 0
        for text, mark in pieces:
            new_r = OxmlElement("w:r")
            if rPr is not None:
                new_r.append(copy.deepcopy(rPr))
            if mark:
                rPr_new = new_r.find(qn("w:rPr"))
                if rPr_new is None:
                    rPr_new = OxmlElement("w:rPr")
                    new_r.insert(0, rPr_new)
                va = OxmlElement("w:vertAlign")
                va.set(qn("w:val"), "superscript" if mark > 0 else "subscript")
                rPr_new.append(va)
            t = OxmlElement("w:t")
            t.set(qn("xml:space"), "preserve")
            t.text = text
            new_r.append(t)
            anchor.addnext(new_r)
            anchor = new_r
            made += 1
        run._r.getparent().remove(run._r)
        total += made
    return total


def apply_marks(doc) -> int:
    """全文档处理。幂等：已有角标的 run 会跳过。"""
    total = 0
    for para in doc.paragraphs:
        total += apply_marks_to_paragraph(para)
    return total


def apply_marks_text(text: str):
    """给预览用：返回 [(文本, 角标)]。"""
    return split_marks(text)

