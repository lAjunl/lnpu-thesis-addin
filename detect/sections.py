"""分节与页码页眉。

把文档整理成三节（学校规范）：
    封面（无页眉、无页码）
    前置（页眉 + 大写罗马数字页码）
    正文（页眉 + 阿拉伯数字页码，从 1 起）

风险提示：这一步会改动文档结构。调用方必须先备份，并且只应该在用户确认后执行。
"""
from __future__ import annotations

from dataclasses import dataclass, field

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .reflow import set_run_font

# ------------------------------------------------------------------ #
# 默认规范
# ------------------------------------------------------------------ #

DEFAULT_HEADER_TEXT = "硕士学位论文"

# 页眉横线线型 → w:val
LINE_STYLES = {"single": "single", "double": "double",
              "dashed": "dashed", "dotted": "dotted",
              "dashDot": "dashDot", "wave": "wave"}


@dataclass
class SectionPlan:
    """一节的目标状态。"""
    header_text: str | None      # None = 清空（无页眉）
    footer_page: bool            # 页脚是否放页码
    number_format: str | None    # 'upperRoman' | 'decimal' | None
    number_start: int | None = None
    link_to_previous: bool = False


PLAN_THESIS = [
    SectionPlan(header_text=None, footer_page=False, number_format=None),
    SectionPlan(header_text=DEFAULT_HEADER_TEXT, footer_page=True,
                number_format="upperRoman", number_start=1),
    SectionPlan(header_text=DEFAULT_HEADER_TEXT, footer_page=True,
                number_format="decimal", number_start=1),
]


# ------------------------------------------------------------------ #
# 底层 XML
# ------------------------------------------------------------------ #

def _sub(parent, tag: str):
    return parent.find(qn("w:" + tag))


def _remove(parent, tag: str) -> None:
    if parent is None:
        return
    for el in parent.findall(qn("w:" + tag)):
        parent.remove(el)


def _clear_paragraph(paragraph) -> None:
    """删掉段落里除 pPr 之外的所有子元素（保留段落本身）。"""
    for child in list(paragraph._p):
        if child.tag == qn("w:pPr"):
            continue
        paragraph._p.remove(child)


def _set_page_format(section, fmt: str | None, start: int | None) -> None:
    sectPr = section._sectPr
    _remove(sectPr, "pgNumType")
    if fmt is None:
        return
    el = OxmlElement("w:pgNumType")
    el.set(qn("w:fmt"), fmt)
    if start is not None:
        el.set(qn("w:start"), str(start))
    pgMar = _sub(sectPr, "pgMar")
    if pgMar is not None:
        pgMar.addnext(el)
    else:
        sectPr.append(el)


def _add_page_field(paragraph) -> None:
    """在段落末尾追加 PAGE 域。"""
    def run_with(child):
        r = OxmlElement("w:r")
        r.append(child)
        return r

    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")

    for item in (begin, instr, sep, text, end):
        paragraph._p.append(run_with(item))


# ------------------------------------------------------------------ #
# 分节
# ------------------------------------------------------------------ #

def insert_break_before(paragraph):
    """在指定段落之前插入分节符。

    做法：把 sectPr 挂到前一段的 pPr 上——Word 里"某段的 pPr 带 sectPr"
    表示该段是其所在节的最后一段，因此下一段自然成为新节的开头。
    这样不会新增段落、不会打乱分页。

    若前一段已经带有 sectPr（说明它已经是某一节的结尾），
    才退而求其次插入一个空段落来承载分节符。
    """
    prev = paragraph._p.getprevious()
    if prev is not None and prev.tag == qn("w:p"):
        pPr = prev.find(qn("w:pPr"))
        if pPr is None:
            pPr = OxmlElement("w:pPr")
            prev.insert(0, pPr)
        if pPr.find(qn("w:sectPr")) is None:
            sectPr = OxmlElement("w:sectPr")
            pPr.append(sectPr)
            return sectPr

    # 退化路径：插入空段落承载分节符
    new_p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    sectPr = OxmlElement("w:sectPr")
    pPr.append(sectPr)
    new_p.append(pPr)
    paragraph._p.addprevious(new_p)
    return sectPr


def split_document(doc, front_index: int, body_index: int) -> list:
    """在 front_index（前置起点）与 body_index（正文起点）处切分。

    返回三节的列表 [封面, 前置, 正文]。
    若文档已经是 3 节，则不再切分，直接返回（只配置）。
    """
    if not (0 < front_index < body_index < len(doc.paragraphs)):
        raise ValueError(f"切分点非法：front={front_index} body={body_index} n={len(doc.paragraphs)}")

    existing = len(doc.sections)
    if existing >= 3:
        return list(doc.sections)

    # 先切后面的，避免插入元素导致前面的索引漂移
    insert_break_before(doc.paragraphs[body_index])
    insert_break_before(doc.paragraphs[front_index])

    sections = list(doc.sections)
    if len(sections) != 3:
        raise RuntimeError(f"切分后应为 3 节，实际 {len(sections)} 节")
    return sections


# ------------------------------------------------------------------ #
# 配置
# ------------------------------------------------------------------ #

def _set_header_line(paragraph, width_pt: float, color: str,
                     style: str = "single", gap_pt: float = 0.0) -> None:
    """页眉文字下面那条横线。0 = 不要线。sz 的单位是 1/8 磅。

    w:space 是"横线距正文的距离"（磅）——默认写死 1 会导致文字贴着线，
    所以这里用设置里的 header_gap。
    """
    pPr = paragraph._p.get_or_add_pPr()
    _remove(pPr, "pBdr")
    if width_pt <= 0:
        return
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), LINE_STYLES.get(style, "single"))
    bottom.set(qn("w:sz"), str(max(1, int(round(width_pt * 8)))))
    bottom.set(qn("w:space"), str(max(0, int(round(gap_pt)))))
    bottom.set(qn("w:color"), (color or "auto").replace("#", ""))
    pBdr.append(bottom)
    pPr.append(pBdr)


def _apply_header(doc, settings) -> None:
    """页眉文字的字体/字号/加粗/字色 + 横线（含线与文字的距离）。"""
    width = float(getattr(settings, "header_line", 0) or 0)
    color = getattr(settings, "header_line_color", "000000") or "000000"
    style = getattr(settings, "header_line_style", "single")
    gap = float(getattr(settings, "header_gap", 4.0) or 0)
    cjk = getattr(settings, "header_font", "") or "宋体"
    latin = getattr(settings, "header_ascii", "") or "Times New Roman"
    size = float(getattr(settings, "header_size", 10.5) or 10.5)
    bold = bool(getattr(settings, "header_bold", False))
    hcolor = (getattr(settings, "header_color", "") or "").replace("#", "")
    for section in doc.sections:
        header = section.header
        if not header.paragraphs:
            continue
        p = header.paragraphs[0]
        if not p.runs:
            continue
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            set_run_font(run, cjk, latin, size, bold, hcolor)
        _set_header_line(p, width, color, style, gap)


def configure(doc, plans: list[SectionPlan] | None = None,
              settings=None) -> list:
    """按 plans 配置每一节。"""
    plans = plans or PLAN_THESIS
    sections = list(doc.sections)
    if len(plans) == 1 and len(sections) > 1:
        # 期刊论文只有一节计划，但原文可能已经有若干节 → 全部套用同一计划
        plans = plans * len(sections)
    if len(sections) != len(plans):
        raise ValueError(f"节数 {len(sections)} 与计划数 {len(plans)} 不符")

    for i, (section, plan) in enumerate(zip(sections, plans)):
        # 页眉
        header = section.header
        header.is_linked_to_previous = plan.link_to_previous
        if header.paragraphs:
            p = header.paragraphs[0]
            _clear_paragraph(p)
            if plan.header_text:
                p.add_run(plan.header_text)
        # 页脚：页码一律居中（学校规范与 Word 默认都是居中，
        # 旧版没设对齐 → 页码贴在左边）
        footer = section.footer
        footer.is_linked_to_previous = plan.link_to_previous
        if footer.paragraphs:
            p = footer.paragraphs[0]
            _clear_paragraph(p)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if plan.footer_page:
                _add_page_field(p)
        _set_page_format(section, plan.number_format, plan.number_start)

    if settings is not None:
        _apply_header(doc, settings)
    return sections


def plans_from_settings(settings) -> list[SectionPlan]:
    """按设置生成各节计划。enable_header=False 时所有节都不写页眉。"""
    profile = getattr(settings, "profile", "thesis")
    text = settings.header_text if getattr(settings, "enable_header", True) else None
    nums = getattr(settings, "enable_page_numbers", True)

    if profile == "journal":
        # 期刊论文：一节到底，页码从头阿拉伯数字，默认无页眉
        return [SectionPlan(header_text=text, footer_page=nums,
                            number_format="decimal" if nums else None,
                            number_start=1)]

    if not nums:
        return [
            SectionPlan(header_text=None, footer_page=False, number_format=None),
            SectionPlan(header_text=text, footer_page=False, number_format=None),
            SectionPlan(header_text=text, footer_page=False, number_format=None),
        ]
    return [
        SectionPlan(header_text=None, footer_page=False, number_format=None),
        SectionPlan(header_text=text, footer_page=True,
                    number_format="upperRoman", number_start=1),
        SectionPlan(header_text=text, footer_page=True,
                    number_format="decimal", number_start=1),
    ]


def apply(doc, front_index: int, body_index: int,
          plans: list[SectionPlan] | None = None,
          settings=None) -> list:
    """切分 + 配置。会改动文档结构，调用前请先备份。"""
    split_document(doc, front_index, body_index)
    return configure(doc, plans, settings)


def diagnose(doc) -> dict:
    """读回分节与页码状态，用于验证。"""
    out = []
    for i, s in enumerate(doc.sections):
        sectPr = s._sectPr
        fmt = None
        start = None
        el = _sub(sectPr, "pgNumType")
        if el is not None:
            fmt = el.get(qn("w:fmt"))
            start = el.get(qn("w:start"))
        hdr = s.header.paragraphs[0].text.strip() if s.header.paragraphs else ""
        ftr = s.footer.paragraphs[0].text.strip() if s.footer.paragraphs else ""
        out.append({
            "index": i + 1,
            "header": hdr,
            "footer_text": ftr,
            "pgNumType_fmt": fmt,
            "pgNumType_start": start,
            "header_linked": s.header.is_linked_to_previous,
            "footer_linked": s.footer.is_linked_to_previous,
        })
    return {"sections": out}
