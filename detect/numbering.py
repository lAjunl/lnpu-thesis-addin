"""自动编号：标题号 / 图号 / 表号。

规则（学校规范）：
    标题：1 / 1.1 / 1.1.1，按层级和出现顺序生成
    图：图N.M（N = 当前章号，M = 该章内第几图）
    表：表N.M（同上）

已有编号的段落默认不动；调用方要在应用前把预览给用户看，
因为这一步会改文字内容。
"""
from __future__ import annotations

import copy
import re
from dataclasses import dataclass

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .classify import Detection, effective_role, is_heading_role


# 段落开头已有的编号。不剥掉的话会出现"图4.1 Figure 1. ..."、"7.2 7.2 xxx"这种重复。
RE_CAP_LABEL = re.compile(
    r"^\s*(?:图|表|Figure|Fig\.?|Table)\s*\d+(?:\s*[.\-–—]\s*\d+)*\s*[.、．:：]?\s*",
    re.I,
)
# 只剥"编号词"本身（用于已经判定为图注/表注的段落，连不带编号的
# "Fig.  (a) …" / "Fig.  Keyword …" 前缀也要剥掉，否则会出现
# "图1.1 Fig. 1 …"这种重复）
RE_CAP_WORD = re.compile(
    r"^\s*(?:图|表|Figure|Fig\.?|Table)\b[\s.．:：)】\]]*", re.I)
RE_CAP_EN = re.compile(r"^\s*(Fig(?:ure)?|Table)\b", re.I)
RE_HEAD_LABEL = re.compile(
    r"^\s*(?:\d+(?:\s*\.\s*\d+)*\s*[.、．]?|[一二三四五六七八九十百]+、)\s+"
)
# 从已有编号里取层级数字（1.2 → [1, 2]）
RE_NUM_PARTS = re.compile(r"^(\d+(?:\.\d+)*)")
# 已有图号/表号：图3.2 / 图 3.2 / Fig. 3.2 / Table 3-2
RE_CAP_NUM = re.compile(
    r"^\s*(?:图|表|Figure|Fig\.?|Table)\s*(\d+)(?:\s*[.\-–—]\s*(\d+))?",
    re.I,
)


def _heading_parts(number_text) -> list[int] | None:
    """把已有标题编号解析成层级数字；解析不了返回 None。"""
    if not number_text:
        return None
    m = RE_NUM_PARTS.match(number_text.strip())
    if not m:
        return None
    try:
        return [int(p) for p in m.group(1).split(".")]
    except ValueError:
        return None


def _caption_parts(text) -> tuple[int, int | None] | None:
    """把已有图号/表号解析成 (N, M?)；解析不了返回 None。"""
    m = RE_CAP_NUM.match((text or "").strip())
    if not m:
        return None
    return int(m.group(1)), (int(m.group(2)) if m.group(2) else None)


def _strip_leading(text: str, rx: re.Pattern, limit: int = 3) -> str:
    for _ in range(limit):
        new = rx.sub("", text, count=1)
        if new == text:
            break
        text = new
    return text


def clean_caption(text: str, aggressive: bool = False) -> str:
    """剥掉图注/表注开头已有的图号表号（可能中文英文各写一遍）。

    aggressive=True 用在"这段已经判定为图注/表注"的场合：连不带编号的
    `Fig.  (a) …` / `Fig.  Keyword …` 前缀也剥掉。
    """
    out = _strip_leading(text, RE_CAP_LABEL)
    if aggressive:
        out = _strip_leading(out, RE_CAP_WORD, limit=1)
    return out


def _auto_caption_prefix(detections, feats, kind: str) -> str | None:
    """原稿图注/表注普遍用 Fig. / Table 开头时，沿用英文前缀。

    期刊稿原稿是 "Fig. …"，排版还硬塞一个"图1.1"既重复又别扭；
    但用户如果在「图号」卡片里手动改过前缀，就以设置为准
    （那时 Settings.caption_prefix_auto 会被置为 False）。
    """
    role = "figure_caption" if kind == "figure" else "table_caption"
    texts = [feats[d.index].text for d in detections if effective_role(d) == role
             and feats[d.index].char_len > 0]
    if len(texts) < 2:
        return None
    if kind == "figure":
        if all(re.match(r"^\s*Figure\b", t) for t in texts):
            return "Figure"
        if all(re.match(r"^\s*Fig\.\s*", t) or t.strip().lower() == "fig."
               for t in texts):
            return "Fig."
        if all(re.match(r"^\s*Fig\b", t) for t in texts):
            return "Fig"
        return None
    if all(re.match(r"^\s*Table\b", t, re.I) for t in texts):
        return "Table"
    return None


def clean_heading(text: str) -> str:
    """剥掉标题开头已有的编号。"""
    return _strip_leading(text, RE_HEAD_LABEL, limit=1)


def _delete_leading_chars(paragraph, n: int) -> None:
    """从段落开头删掉 n 个字符（跨 run），保留 run 的格式。"""
    for run in paragraph.runs:
        if n <= 0:
            break
        t = run.text
        if not t:
            continue
        cut = min(n, len(t))
        run.text = t[cut:]
        n -= cut


@dataclass
class NumberPlan:
    index: int          # 段落序号
    text: str           # 将插入的编号文字
    kind: str           # "heading" / "figure" / "table"
    old: str | None     # 原有编号（若有）


def _has_number(d: Detection) -> bool:
    return bool(d.number)


def plan_heading_numbers(detections: list[Detection],
                         body_index: int | None = None,
                         settings=None) -> list[NumberPlan]:
    """按层级生成标题编号。

    `number_keep_existing`（默认开）时**保留原稿已有的编号**：读到已有编号
    就把计数器对齐到它，只给缺编号的段落补号；关掉则全部重排。

    body_index 之后才开始编号：封面/前置部分（摘要、Highlights 之类）
    不参与章号，否则会出现 "0.1 Highlights" 这种章号为 0 的编号。

    sep 取自 settings.heading_sep（章号与节号之间的分隔：1.1 / 1-1 / 11）。
    """
    counters = {1: 0, 2: 0, 3: 0}
    sep = getattr(settings, "heading_sep", ".") if settings is not None else "."
    if sep is None:
        sep = ""
    keep = bool(getattr(settings, "number_keep_existing", True)) \
        if settings is not None else True
    plans: list[NumberPlan] = []
    for d in detections:
        role = effective_role(d)
        if not (role.startswith("heading") and role != "heading"):
            continue
        if body_index is not None and d.index < body_index:
            continue
        lvl = int(role[-1])
        # 还没遇到一级标题就出现次级标题时，按一级处理，
        # 否则会编出 "0.1" 这种章号为 0 的号（实测期刊稿会踩到）
        if lvl > 1 and counters[1] == 0:
            lvl = 1
        if keep:
            parts = _heading_parts(d.number)
            if parts:
                # 原稿写了编号：原样保留，计数器对齐到它，后面接着编
                for i in range(1, 4):
                    counters[i] = parts[i - 1] if i <= len(parts) else 0
                continue
        counters[lvl] += 1
        for deeper in range(lvl + 1, 4):
            counters[deeper] = 0
        parts = [str(counters[i]) for i in range(1, lvl + 1)]
        plans.append(NumberPlan(d.index, sep.join(parts) + " ", "heading", d.number))
    return plans


def plan_caption_numbers(detections: list[Detection], kind: str,
                         settings=None, feats=None) -> list[NumberPlan]:
    """给图注/表注生成编号。

    默认 图N.M / 表N.M（N=章号，M=本章第几个）。可用设置改：
    前缀文字、分隔符（. / - / 空）、是否按章、起始号。
    设置 caption_prefix_auto（默认开）时，原稿图注普遍用 Fig./Table 开头
    就沿用英文前缀，避免「图1.1 Fig. 1 …」这种重复。

    `number_keep_existing`（默认开）时，原稿已有的图号/表号原样保留，
    计数器对齐到它，只给缺编号的图注补号。
    """
    default = ((getattr(settings, "caption_fig_prefix", "图") or "图")
               if kind == "figure"
               else (getattr(settings, "caption_tab_prefix", "表") or "表"))
    label = default
    if feats is not None and getattr(settings, "caption_prefix_auto", False):
        auto = _auto_caption_prefix(detections, feats, kind)
        if auto:
            label = auto
    sep = getattr(settings, "caption_sep", ".") or ""
    by_chapter = getattr(settings, "caption_by_chapter", True)
    start = int(getattr(settings, "caption_start", 1) or 1)
    keep = bool(getattr(settings, "number_keep_existing", True)) \
        if settings is not None else True
    chapter = 0
    per_chapter = start - 1
    plans: list[NumberPlan] = []
    for d in detections:
        role = effective_role(d)
        if by_chapter and role == "heading1":
            chapter += 1
            per_chapter = start - 1
            continue
        if role != ("figure_caption" if kind == "figure" else "table_caption"):
            continue
        if keep:
            txt = feats[d.index].text if feats is not None else (d.number or "")
            got = _caption_parts(txt)
            if got is not None:
                n, m = got
                if by_chapter:
                    if m is not None:
                        chapter, per_chapter = n, m
                    else:
                        per_chapter = n
                else:
                    per_chapter = n
                continue
        per_chapter += 1
        n = chapter if by_chapter else max(1, chapter)
        # 中文前缀"图1.1"不加空格；英文前缀 "Fig. 1.1" / "Figure 1.1" 加空格
        gap = "" if not label.isascii() or label == "" else " "
        num = (f"{label}{gap}{n}{sep}{per_chapter}" if by_chapter
               else f"{label}{gap}{per_chapter}")
        plans.append(NumberPlan(d.index, f"{num} ", kind, d.number))
    return plans


def _prefix_run(paragraph, text: str) -> None:
    """在段落开头插入一个带文字的 run，尽量沿用第一个 run 的格式。"""
    src = paragraph.runs[0]._r if paragraph.runs else None
    new_r = OxmlElement("w:r")
    if src is not None:
        rPr = src.find(qn("w:rPr"))
        if rPr is not None:
            new_r.append(copy.deepcopy(rPr))
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    new_r.append(t)
    if src is not None:
        src.addprevious(new_r)
    else:
        paragraph._p.append(new_r)


def apply_plans(doc, plans: list[NumberPlan]) -> int:
    """把编号插进文档。返回插入数量。"""
    n = 0
    for plan in plans:
        para = doc.paragraphs[plan.index]
        # 原文已有的编号先剥掉，避免"图4.1 Figure 1."这种重复
        if plan.kind in ("figure", "table"):
            old = "".join(r.text for r in para.runs)
            # aggressive：这段已经判定为图注/表注，连 "Fig.  " 这种
            # 不带编号的前缀也剥掉，否则会变成「图1.1 Fig. 1 …」
            cleaned = clean_caption(old, aggressive=True)
            if cleaned != old:
                _delete_leading_chars(para, len(old) - len(cleaned))
        elif plan.kind == "heading":
            old = "".join(r.text for r in para.runs)
            cleaned = clean_heading(old)
            if cleaned != old:
                _delete_leading_chars(para, len(old) - len(cleaned))
        _prefix_run(para, plan.text)
        n += 1
    return n


def summarize(plans: list[NumberPlan]) -> str:
    from collections import Counter

    if not plans:
        return "无"
    c = Counter(p.kind for p in plans)
    bits = []
    if c.get("heading"):
        bits.append(f"标题号 {c['heading']} 个")
    if c.get("figure"):
        bits.append(f"图号 {c['figure']} 个")
    if c.get("table"):
        bits.append(f"表号 {c['table']} 个")
    return "、".join(bits)
