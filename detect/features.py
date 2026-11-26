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


class StyleIndex:
    """按文档一次性建立样式索引，避免 python-docx 每次读 p.style / tbl.style
    线性扫描整个 styles.xml（含 default_for 的 O(N×S) 灾难）。
    """

    def __init__(self, doc: Document):
        from docx.enum.style import WD_STYLE_TYPE

        w_type = qn("w:type")
        self.by_type_id: dict[tuple[str, str], Any] = {}
        self.by_id: dict[str, Any] = {}
        for st in doc.styles:
            sid = st.style_id
            if not sid:
                continue
            self.by_id[sid] = st
            tp = st.element.get(w_type) or ""
            if tp:
                self.by_type_id[(tp, sid)] = st
        try:
            self.default_para = doc.styles.default(WD_STYLE_TYPE.PARAGRAPH)
        except Exception:
            self.default_para = None
        try:
            self.default_table = doc.styles.default(WD_STYLE_TYPE.TABLE)
        except Exception:
            self.default_table = None
        self._chains: dict[str, list] = {}
        self._w_pPr = qn("w:pPr")
        self._w_pStyle = qn("w:pStyle")
        self._w_tblPr = qn("w:tblPr")
        self._w_tblStyle = qn("w:tblStyle")
        self._w_basedOn = qn("w:basedOn")
        self._w_val = qn("w:val")

    def chain_of(self, style) -> list:
        if style is None:
            return []
        key = (style.element.get(qn("w:type")) or "") + ":" + (style.style_id or str(id(style)))
        cached = self._chains.get(key)
        if cached is not None:
            return cached
        chain = []
        cur = style
        seen = set()
        while cur is not None and len(chain) < 12:
            ck = (cur.element.get(qn("w:type")) or "") + ":" + (cur.style_id or str(id(cur)))
            if ck in seen:
                break
            seen.add(ck)
            chain.append(cur)
            bo = cur.element.find(self._w_basedOn)
            next_id = bo.get(self._w_val) if bo is not None else None
            if next_id:
                tp = cur.element.get(qn("w:type")) or ""
                cur = self.by_type_id.get((tp, next_id)) or self.by_id.get(next_id)
            else:
                cur = None
        self._chains[key] = chain
        return chain

    def para_style(self, p: Paragraph):
        pPr = p._p.find(self._w_pPr)
        if pPr is not None:
            ps = pPr.find(self._w_pStyle)
            if ps is not None:
                sid = ps.get(self._w_val)
                if sid:
                    st = self.by_type_id.get(("paragraph", sid)) or self.by_id.get(sid)
                    if st is not None:
                        return st
        return self.default_para

    def para_chain(self, p: Paragraph) -> list:
        return self.chain_of(self.para_style(p))

    def table_style(self, tbl):
        tblPr = tbl._tbl.find(self._w_tblPr)
        if tblPr is not None:
            ts = tblPr.find(self._w_tblStyle)
            if ts is not None:
                sid = ts.get(self._w_val)
                if sid:
                    st = self.by_type_id.get(("table", sid)) or self.by_id.get(sid)
                    if st is not None:
                        return st
        return self.default_table

    def table_chain(self, tbl) -> list:
        return self.chain_of(self.table_style(tbl))


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
    is_toc_field: bool = False      # Word 自动目录条目（PAGEREF _Toc 域）
    left_indent_pt: float | None = None
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


def _visible_runs(p: Paragraph) -> list[tuple[Any, str]]:
    """返回段落里真正有可见文字的 (w:r, text)（含超链接内部，排除域指令）。

    Word 目录标题「目  录」后面常紧跟一个隐藏的 `TOC \\o "1-2"` 域指令 run
    （没有任何 w:t 文字、字号却是 12pt）；旧代码遍历 p.runs 会把这个隐藏 run
    当成段落字号，把 16pt 标题错认成 12pt。同时 TOC 条目的正文藏在
    w:hyperlink 里，不在 p.runs 里，这里一并扫出来。
    """
    out = []
    for r_el in p._p.iter(qn("w:r")):
        txt = "".join(t.text or "" for t in r_el.findall(qn("w:t")))
        if txt:
            out.append((r_el, txt))
    return out


def _size_pt(p: Paragraph, pPr, chain, doc_defaults_rpr) -> float | None:
    """段落字号。run 级按**字符数加权**取众数，不用 max()——
    一个 14 磅的空格 run 不该把整段（10 磅正文）顶成标题。"""
    counter: dict[float, int] = {}
    for r_el, txt in _visible_runs(p):
        if not txt.strip():
            continue
        pt = _half_pt_to_pt(_find_in_pPr(_rpr_of(r_el), "sz"))
        if pt is None:
            continue
        counter[pt] = counter.get(pt, 0) + len(txt.strip())
    if counter:
        return sorted(counter.items(), key=lambda kv: (-kv[1], kv[0]))[0][0]
    # 2) 段落 pPr 的 rPr
    v = _half_pt_to_pt(_find_in_pPr(_rpr_of(pPr), "sz"))
    if v is not None:
        return v
    # 3) 样式链
    for st in chain:
        v = _half_pt_to_pt(_find_in_pPr(_style_rPr(st), "sz"))
        if v is not None:
            return v
    # 4) docDefaults（很多稿子的默认字体/字号只写在这里）
    return _half_pt_to_pt(_find_in_pPr(doc_defaults_rpr, "sz"))


def _bold_of(p: Paragraph, pPr, chain, doc_defaults_rpr=None) -> bool | None:
    inherited: bool | None = None
    for el in (_rpr_of(pPr), *[_style_rPr(st) for st in chain],
               doc_defaults_rpr):
        b = _find_in_pPr(el, "b")
        if b is not None:
            v = _get_w(b, "val")
            inherited = True if v is None else v not in ("0", "false", "off")
            break

    # 按非空字符数加权判定整段是否加粗：避免仅开头标签加粗
    # （如 "摘要：..." / "Keywords: ..."）被误判为整段加粗
    bold_chars = 0
    total_chars = 0
    any_explicit = False
    for r_el, txt in _visible_runs(p):
        s = txt.strip()
        if not s:
            continue
        n = len(s)
        total_chars += n
        b = _find_in_pPr(_rpr_of(r_el), "b")
        if b is not None:
            any_explicit = True
            v = _get_w(b, "val")
            is_b = True if v is None else v not in ("0", "false", "off")
        else:
            is_b = bool(inherited)
        if is_b:
            bold_chars += n

    if total_chars > 0 and any_explicit:
        return bold_chars > total_chars / 2.0
    return inherited


def _east_asia_of(p: Paragraph, pPr, chain, doc_defaults_rpr=None) -> str:
    for r_el, txt in _visible_runs(p):
        if not txt.strip():
            continue
        rPr = _rpr_of(r_el)
        if rPr is not None:
            rf = rPr.find(qn("w:rFonts"))
            if rf is not None and rf.get(qn("w:eastAsia")):
                return rf.get(qn("w:eastAsia"))
    for el in (_rpr_of(pPr), *[_style_rPr(st) for st in chain],
               doc_defaults_rpr):
        if el is None:
            continue
        rf = el.find(qn("w:rFonts"))
        if rf is not None and rf.get(qn("w:eastAsia")):
            return rf.get(qn("w:eastAsia"))
    return ""


def _ascii_of(p: Paragraph, pPr, chain, doc_defaults_rpr=None) -> str:
    for r_el, txt in _visible_runs(p):
        if not txt.strip():
            continue
        rPr = _rpr_of(r_el)
        if rPr is not None:
            rf = rPr.find(qn("w:rFonts"))
            if rf is not None:
                v = rf.get(qn("w:ascii")) or rf.get(qn("w:hAnsi"))
                if v:
                    return v
    for el in (_rpr_of(pPr), *[_style_rPr(st) for st in chain],
               doc_defaults_rpr):
        if el is None:
            continue
        rf = el.find(qn("w:rFonts"))
        if rf is not None and rf.get(qn("w:ascii")):
            return rf.get(qn("w:ascii"))
    return ""


def _line_of(pPr, chain, defaults_pPr=None):
    """返回 (rule, value)。rule ∈ exactly / multiple / at_least / single / None。"""
    for el in [pPr, *[_style_pPr(st) for st in chain], defaults_pPr]:
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


def _parse_css_pt(val: str) -> float | None:
    m = re.match(r"^\s*([0-9]+(?:\.[0-9]+)?)\s*(pt|in|cm|mm|px)?\s*$", val or "", re.I)
    if not m:
        return None
    v = float(m.group(1))
    u = (m.group(2) or "pt").lower()
    if u == "pt":
        return v
    if u == "in":
        return v * 72.0
    if u == "cm":
        return v * 72.0 / 2.54
    if u == "mm":
        return v * 72.0 / 25.4
    if u == "px":
        return v * 0.75
    return v


def vml_size_pt(el) -> tuple[float, float] | None:
    """从 VML v:shape / v:rect 的 style 属性读尺寸（如 width:96.4pt;height:85.15pt）。"""
    cur = el
    for _ in range(4):
        if cur is None:
            break
        st = cur.get("style") or ""
        if "width" in st and "height" in st:
            w = h = None
            for part in st.split(";"):
                k, _, v = part.partition(":")
                k = k.strip().lower()
                if k == "width":
                    w = _parse_css_pt(v)
                elif k == "height":
                    h = _parse_css_pt(v)
            if w and h:
                return w, h
        cur = cur.getparent()
    return None


def image_size(p: Paragraph):
    """插图的原始尺寸（磅）。先读 DrawingML wp:extent（EMU），再兜底 VML style。

    有了原始尺寸，预览就能按 Word 里的真实大小显示，并且保留多倍像素，
    不会因为"像素数 = 磅数"在高 DPI 屏上被放大而发虚。
    """
    for el in p._p.iter(qn("wp:extent")):
        cx, cy = el.get("cx"), el.get("cy")
        if cx and cy:
            try:
                return int(cx) / 12700.0, int(cy) / 12700.0
            except ValueError:
                pass
    for el in p._p.iter(VML + "imagedata"):
        sz = vml_size_pt(el)
        if sz:
            return sz
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


def _is_toc_field(p: Paragraph) -> bool:
    """Word 自动目录生成的条目（内含 PAGEREF _Toc 域指令）。"""
    for el in p._p.iter(qn("w:instrText")):
        t = el.text or ""
        if "PAGEREF" in t.upper() and "_TOC" in t.upper():
            return True
    return False


def _left_indent_of(pPr, chain) -> float | None:
    for el in [pPr, *[_style_pPr(st) for st in chain]]:
        ind = _find_in_pPr(el, "ind")
        if ind is None:
            continue
        v = _space_pt(ind, "left")
        if v is not None:
            return v
        v = _space_pt(ind, "start")
        if v is not None:
            return v
    return None


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
    # docDefaults 在 styles.xml 里（不在 document.xml），层级是 docDefaults/rPrDefault/rPr。
    # 以前写成 doc.element.find(...) 永远取到 None —— 样式派生格式的稿子会大面积读空。
    defaults_rpr = None
    defaults_pPr = None
    try:
        defaults = doc.styles.element.find(qn("w:docDefaults"))
        if defaults is not None:
            ppr_default = defaults.find(qn("w:pPrDefault"))
            if ppr_default is not None:
                defaults_pPr = ppr_default.find(qn("w:pPr"))
            rpr_default = defaults.find(qn("w:rPrDefault"))
            if rpr_default is not None:
                defaults_rpr = rpr_default.find(qn("w:rPr"))
    except Exception:      # noqa: BLE001
        defaults_rpr = None
        defaults_pPr = None

    si = StyleIndex(doc)
    w_pPr = qn("w:pPr")
    out: list[ParaFeatures] = []

    for i, p in enumerate(doc.paragraphs):
        style = si.para_style(p)
        chain = si.chain_of(style)
        pPr = p._p.find(w_pPr)

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
            is_toc_field=_is_toc_field(p),
        )

        f.align = _align_of(pPr, chain)
        f.size_pt = _size_pt(p, pPr, chain, defaults_rpr)
        f.bold = _bold_of(p, pPr, chain, defaults_rpr)
        f.east_asia = _east_asia_of(p, pPr, chain, defaults_rpr)
        f.ascii_font = _ascii_of(p, pPr, chain, defaults_rpr)
        f.line_rule, f.line_value = _line_of(pPr, chain, defaults_pPr)
        f.outline_level = _outline_of(pPr, chain)
        f.left_indent_pt = _left_indent_of(pPr, chain)

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

    # 最近图片距离（前后两次线性扫，O(N)）
    last_img: int | None = None
    for f in out:
        f.prev_image_idx = last_img
        if f.has_image:
            last_img = f.index
    next_img: int | None = None
    for f in reversed(out):
        f.next_image_idx = next_img
        if f.has_image:
            next_img = f.index

    return out


def body_baseline(feats: list[ParaFeatures]) -> dict:
    """从"长段落"里统计正文基线格式（字号/行距/对齐/是否加粗）。"""
    from collections import Counter

    long_paras = [f for f in feats if f.char_len >= 40 and not f.has_image]
    if not long_paras:
        long_paras = [f for f in feats if f.char_len >= 20] or feats
    sizes = Counter(f.size_pt for f in long_paras if f.size_pt is not None)
    lines = Counter(f.line_rule for f in long_paras if f.line_rule is not None)
    values = Counter(f.line_value for f in long_paras if f.line_value is not None)
    aligns = Counter(f.align for f in long_paras if f.align is not None)
    bolds = Counter(f.bold for f in long_paras)
    return {
        "size": sizes.most_common(1)[0][0] if sizes else None,
        "line_rule": lines.most_common(1)[0][0] if lines else None,
        # 段落都没写行距时 values 是空的，按 long_paras 判断会 IndexError
        "line_value": values.most_common(1)[0][0] if values else None,
        "align": aligns.most_common(1)[0][0] if aligns else None,
        "bold": bolds.most_common(1)[0][0] if bolds else False,
        "n": len(long_paras),
    }


_DOC_CACHE: dict[tuple[str, float], str] = {}


def ensure_docx(path) -> "Path":
    """若是 .docx 原样返回；若是旧版 .doc，用本机 Word 静默另存为临时 .docx。"""
    from pathlib import Path

    p = Path(path)
    if p.suffix.lower() == ".docx":
        return p
    if p.suffix.lower() != ".doc":
        raise ValueError("只支持 .docx 或 .doc 格式的 Word 文档。")

    key = (str(p.resolve()), p.stat().st_mtime)
    hit = _DOC_CACHE.get(key)
    if hit and Path(hit).exists():
        return Path(hit)

    import tempfile
    try:
        import win32com.client
    except Exception as exc:  # noqa: BLE001
        raise RuntimeError(
            "无法自动转换 .doc（缺少 pywin32/Word），请先在 Word 中另存为 .docx 再打开。"
        ) from exc

    tmp_dir = Path(tempfile.gettempdir()) / "thesis_doc_conv"
    tmp_dir.mkdir(parents=True, exist_ok=True)
    out = tmp_dir / f"{p.stem}_{abs(hash(key)) & 0xFFFF:04x}.docx"
    if out.exists():
        out.unlink()

    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(str(p.resolve()), ReadOnly=True, AddToRecentFiles=False)
        try:
            doc.SaveAs2(str(out.resolve()), FileFormat=16)  # wdFormatXMLDocument
        finally:
            doc.Close(False)
    finally:
        word.Quit()

    _DOC_CACHE[key] = str(out)
    return out

