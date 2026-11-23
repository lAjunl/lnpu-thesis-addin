"""段落特征提取。

关键设计：格式特征取"生效值"——段落直接格式 → 段落样式的 pPr/rPr →
basedOn 链 → docDefaults。真实文档里大量段落没有任何直接格式，
字号全部来自样式，只读 run.font.size 会得到一片 None。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph


# ------------------------------------------------------------------ #
# 底层 XML 辅助
# ------------------------------------------------------------------ #

def _find_in_pPr(pPr, local: str):
    if pPr is None:
        return None
    return pPr.find(qn("w:" + local))


def _rpr_of(el) -> Any:
    """取元素（w:pPr 或 w:rPr 或 w:style）下的 w:rPr。"""
    if el is None:
        return None
    return el.find(qn("w:rPr"))


def _style_pPr(style) -> Any:
    return style.element.find(qn("w:pPr"))


def _style_rPr(style) -> Any:
    el = style.element.find(qn("w:rPr"))
    if el is not None:
        return el
    # 有些样式把 rPr 放在基于样式的末尾；再退回 linked style
    return None


def _get_w(el, local: str) -> str | None:
    if el is None:
        return None
    v = el.get(qn("w:" + local))
    return v


def _half_pt_to_pt(el) -> float | None:
    v = _get_w(el, "val")
    if v is None:
        return None
    try:
        return int(v) / 2.0
    except ValueError:
        return None


def _first_not_none(values):
    for v in values:
        if v is not None:
            return v
    return None


# ------------------------------------------------------------------ #
# 样式链
# ------------------------------------------------------------------ #

def style_chain(style, max_depth: int = 12):
    chain = []
    cur = style
    seen = set()
    while cur is not None and len(chain) < max_depth:
        if id(cur) in seen:
            break
        seen.add(id(cur))
        chain.append(cur)
        try:
            cur = cur.base_style
        except Exception:
            cur = None
    return chain


# ------------------------------------------------------------------ #
# 特征
# ------------------------------------------------------------------ #

@dataclass
class ParaFeatures:
    index: int
    text: str
    char_len: int
    style_name: str
    style_id: str

    align: str | None = None
    size_pt: float | None = None
    bold: bool | None = None
    italic: bool | None = None
    east_asia: str = ""
    ascii_font: str = ""
    line_rule: str | None = None
    line_value: float | None = None
    space_before_pt: float | None = None
    space_after_pt: float | None = None
    first_indent_pt: float | None = None
    outline_level: int | None = None

    has_image: bool = False
    has_math: bool = False          # 段落里含公式对象（OMML）
    has_ole: bool = False           # 段落里含 OLE 对象（MathType 公式等）
    in_table: bool = False
    is_first: bool = False

    number_prefix: str | None = None
    number_level: int | None = None
    starts_fig_tab: str | None = None  # "figure" | "table"
    ends_with_punct: bool = False
    word_count: int = 0

    prev_image_idx: int | None = None
    next_image_idx: int | None = None

    reasons: list[str] = field(default_factory=list)


_ALIGN_MAP = {
    "left": "left",
    "center": "center",
    "right": "right",
    "both": "both",
    "justify": "both",
    "distribute": "both",
}


def _align_of(pPr, chain) -> str | None:
    jc = _find_in_pPr(pPr, "jc")
    if jc is not None:
        v = _get_w(jc, "val")
        if v:
            return _ALIGN_MAP.get(v, v)
    for st in chain:
        jc = _find_in_pPr(_style_pPr(st), "jc")
        if jc is not None:
            v = _get_w(jc, "val")
            if v:
                return _ALIGN_MAP.get(v, v)
    return None


def _size_pt(p: Paragraph, pPr, chain, doc_defaults_rpr) -> float | None:
    # 1) 段落直接格式的 run
    sizes = [r.font.size.pt for r in p.runs if r.font.size is not None]
    if sizes:
        return max(sizes)
    # 2) 段落 pPr 的 rPr
    v = _half_pt_to_pt(_find_in_pPr(_rpr_of(pPr), "sz"))
    if v is not None:
        return v
    # 3) 样式链
    for st in chain:
        v = _half_pt_to_pt(_find_in_pPr(_style_rPr(st), "sz"))
        if v is not None:
            return v
    # 4) docDefaults
    return _half_pt_to_pt(_find_in_pPr(doc_defaults_rpr, "sz"))


def _bold_of(p: Paragraph, pPr, chain) -> bool | None:
    for r in p.runs:
        if r.bold is not None:
            return bool(r.bold)
    for el in (_rpr_of(pPr), *[_style_rPr(st) for st in chain]):
        b = _find_in_pPr(el, "b")
        if b is not None:
            v = _get_w(b, "val")
            if v is None:
                return True
            return v not in ("0", "false", "off")
    return None


def _east_asia_of(p: Paragraph, pPr, chain) -> str:
    for r in p.runs:
        rPr = r._r.find(qn("w:rPr"))
        if rPr is not None:
            rf = rPr.find(qn("w:rFonts"))
            if rf is not None and rf.get(qn("w:eastAsia")):
                return rf.get(qn("w:eastAsia"))
    candidates = [_rpr_of(pPr), *[_style_rPr(st) for st in chain]]
    for el in candidates:
        if el is None:
            continue
        rf = el.find(qn("w:rFonts"))
        if rf is not None and rf.get(qn("w:eastAsia")):
            return rf.get(qn("w:eastAsia"))
    return ""


def _ascii_of(p: Paragraph, pPr, chain) -> str:
    for r in p.runs:
        if r.font.name:
            return r.font.name
    candidates = [_rpr_of(pPr), *[_style_rPr(st) for st in chain]]
    for el in candidates:
        if el is None:
            continue
        rf = el.find(qn("w:rFonts"))
        if rf is not None and rf.get(qn("w:ascii")):
            return rf.get(qn("w:ascii"))
    return ""


def _line_of(pPr, chain):
    """返回 (rule, value)。rule ∈ exactly / multiple / at_least / single / None。"""
    for el in [pPr, *[_style_pPr(st) for st in chain]]:
        sp = _find_in_pPr(el, "spacing")
        if sp is None:
            continue
        rule = _get_w(sp, "lineRule") or "auto"
        raw = _get_w(sp, "line")
        if raw is None:
            # Word 默认 auto/240；只有显式写了 lineRule 才有意义
            if _get_w(sp, "lineRule") is not None:
                return rule, None
            continue
        try:
            n = int(raw)
        except ValueError:
            continue
        if rule in ("exact", "atLeast"):
            return ("exactly" if rule == "exact" else "at_least"), n / 20.0
        if rule == "auto":
            return "multiple", n / 240.0
        return rule, n
    return None, None


def _space_pt(el, local: str) -> float | None:
    v = _get_w(el, local)
    if v is None:
        return None
    try:
        return int(v) / 20.0
    except ValueError:
        return None


def _outline_of(pPr, chain) -> int | None:
    for el in [pPr, *[_style_pPr(st) for st in chain]]:
        ol = _find_in_pPr(el, "outlineLvl")
        if ol is not None:
            v = _get_w(ol, "val")
            if v is not None:
                try:
                    return int(v)
                except ValueError:
                    pass
    return None


def _has_image(p: Paragraph) -> bool:
    """真正的插图。注意 w:object 是 OLE 对象（MathType 公式），不算图。"""
    p_el = p._p
    for tag in ("w:drawing", "w:pict"):
        if p_el.find(".//" + qn(tag)) is not None:
            return True
    return False


VML = "{urn:schemas-microsoft-com:vml}"


def _has_ole(p: Paragraph) -> bool:
    return p._p.find(".//" + qn("w:object")) is not None


def ole_size(p: Paragraph):
    """MathType 等 OLE 对象的原始尺寸（磅）。读 w:object 的 dxaOrig/dyaOrig（twips）。"""
    for el in p._p.iter(qn("w:object")):
        w = el.get(qn("w:dxaOrig"))
        h = el.get(qn("w:dyaOrig"))
        if w and h:
            try:
                return int(w) / 20.0, int(h) / 20.0
            except ValueError:
                return None
    return None


def image_size(p: Paragraph):
    """插图的原始尺寸（磅）。读 wp:extent（EMU，1 磅 = 12700 EMU）。

    有了原始尺寸，预览就能按 Word 里的真实大小显示，并且保留多倍像素，
    不会因为"像素数 = 磅数"在高 DPI 屏上被放大而发虚。
    """
    for el in p._p.iter(qn("wp:extent")):
        cx, cy = el.get("cx"), el.get("cy")
        if cx and cy:
            try:
                return int(cx) / 12700.0, int(cy) / 12700.0
            except ValueError:
                return None
    return None


def equation_label(p: Paragraph) -> str:
    r"""公式编号，比如 "(1)"、"(4-1)"。

    MathType 的编号是 Word 域（`SEQ MTEqn`），文字写在 `w:instrText` 里，
    所以 `paragraph.text` 读不到——预览里就"看不到编号"。
    这里把域指令（SEQ / MACROBUTTON / \* MERGEFORMAT）排除掉，
    只留域外面和域结果里的正文（"(" + "1" + ")"）。
    """
    parts = []
    for el in p._p.iter():
        if el.tag == qn("w:t"):
            parts.append(el.text or "")
        elif el.tag == qn("w:instrText"):
            t = el.text or ""
            s = t.strip()
            if not s:
                continue
            if ("\\" in t or "MERGEFORMAT" in t
                    or s.upper().startswith(("SEQ ", "MACROBUTTON"))):
                continue
            parts.append(t)
    return "".join(parts).strip()


def ole_image_blob(p: Paragraph):
    """MathType 之类 OLE 公式的图片形态（v:imagedata 的 r:id）。"""
    for el in p._p.iter(VML + "imagedata"):
        rid = el.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        if not rid:
            continue
        try:
            part = p.part.related_parts.get(rid)
        except Exception:
            part = None
        if part is not None:
            try:
                return part.blob
            except Exception:
                return None
    return None


def _has_math(p: Paragraph) -> bool:
    """段落里是否含 Word 公式对象（OMML）。"""
    p_el = p._p
    return (p_el.find(".//" + qn("m:oMath")) is not None
            or p_el.find(".//" + qn("m:oMathPara")) is not None)


def _math_text(p: Paragraph) -> str:
    """把公式里的 m:t 文本抽出来，给预览做占位显示。"""
    parts = [el.text or "" for el in p._p.iter(qn("m:t"))]
    return "".join(parts).strip()


def _in_table(p: Paragraph) -> bool:
    el = p._p.getparent()
    while el is not None:
        if el.tag == qn("w:tbl"):
            return True
        el = el.getparent()
    return False


# ------------------------------------------------------------------ #
# 编号识别
# ------------------------------------------------------------------ #

import re

RE_NUM = re.compile(r"^\s*(\d+(?:\.\d+)*)\s*[.、．]?\s*")
RE_HEADING_NUMBER = re.compile(r"^\s*(\d+(?:\.\d+)*)\s*[.、．]?\s+\S")
RE_FIG_ZH = re.compile(r"^\s*图\s*\d+", re.IGNORECASE)
RE_TAB_ZH = re.compile(r"^\s*表\s*\d+", re.IGNORECASE)
RE_FIG_EN = re.compile(r"^\s*(Figure|Fig\.?)\s*\d+", re.IGNORECASE)
RE_TAB_EN = re.compile(r"^\s*Table\s*\d+", re.IGNORECASE)
RE_EQN = re.compile(r"^[（(]\s*式\s*\d+(?:\.\d+)?\s*[)）]")
RE_REF = re.compile(r"^\s*\[\d+\]")
RE_TOC = re.compile(r"^\s*(目\s*录|Contents)\b")


def extract_features(doc: Document) -> list[ParaFeatures]:
    styles = list(doc.styles)
    defaults = doc.element.find(".//" + qn("w:docDefaults"))
    defaults_rpr = _rpr_of(defaults)

    cache: dict[int, list] = {}
    out: list[ParaFeatures] = []

    for i, p in enumerate(doc.paragraphs):
        style = p.style
        chain = style_chain(style) if style is not None else []
        pPr = p._p.find(qn("w:pPr"))

        text = p.text.strip()
        f = ParaFeatures(
            index=i,
            text=text,
            char_len=len(text.replace(" ", "")),
            style_name=style.name if style is not None else "",
            style_id=style.style_id if style is not None else "",
            is_first=(i == 0),
            in_table=_in_table(p),
            has_image=_has_image(p),
            has_math=_has_math(p),
            has_ole=_has_ole(p),
        )

        f.align = _align_of(pPr, chain)
        f.size_pt = _size_pt(p, pPr, chain, defaults_rpr)
        f.bold = _bold_of(p, pPr, chain)
        f.east_asia = _east_asia_of(p, pPr, chain)
        f.ascii_font = _ascii_of(p, pPr, chain)
        f.line_rule, f.line_value = _line_of(pPr, chain)
        f.outline_level = _outline_of(pPr, chain)

        sp = _find_in_pPr(pPr, "spacing")
        if sp is not None:
            f.space_before_pt = _space_pt(sp, "before")
            f.space_after_pt = _space_pt(sp, "after")
        ind = _find_in_pPr(pPr, "ind")
        if ind is not None:
            f.first_indent_pt = _space_pt(ind, "firstLine")

        # 编号
        m = RE_NUM.match(text)
        if m:
            f.number_prefix = m.group(1)
            f.number_level = m.group(1).count(".") + 1
        if RE_FIG_ZH.match(text) or RE_FIG_EN.match(text):
            f.starts_fig_tab = "figure"
        elif RE_TAB_ZH.match(text) or RE_TAB_EN.match(text):
            f.starts_fig_tab = "table"

        f.ends_with_punct = bool(text) and text[-1] in "。．.；;！!？?"
        f.word_count = len(text.split())

        out.append(f)

    # 最近图片距离
    img_idxs = [f.index for f in out if f.has_image]
    for f in out:
        prev = next((x for x in reversed(img_idxs) if x < f.index), None)
        nxt = next((x for x in img_idxs if x > f.index), None)
        f.prev_image_idx = prev
        f.next_image_idx = nxt

    return out


def body_baseline(feats: list[ParaFeatures]) -> dict:
    """从"长段落"里统计正文基线格式（字号/行距/对齐/是否加粗）。"""
    from collections import Counter

    long_paras = [f for f in feats if f.char_len >= 40 and not f.has_image]
    if not long_paras:
        long_paras = [f for f in feats if f.char_len >= 20] or feats
    sizes = Counter(f.size_pt for f in long_paras if f.size_pt is not None)
    lines = Counter(f.line_rule for f in long_paras if f.line_rule is not None)
    aligns = Counter(f.align for f in long_paras if f.align is not None)
    bolds = Counter(f.bold for f in long_paras)
    return {
        "size": sizes.most_common(1)[0][0] if sizes else None,
        "line_rule": lines.most_common(1)[0][0] if lines else None,
        "line_value": Counter(
            f.line_value for f in long_paras if f.line_value is not None
        ).most_common(1)[0][0] if long_paras else None,
        "align": aligns.most_common(1)[0][0] if aligns else None,
        "bold": bolds.most_common(1)[0][0] if bolds else False,
        "n": len(long_paras),
    }
