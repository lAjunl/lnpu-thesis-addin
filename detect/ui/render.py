"""双栏预览：分页渲染（QTextDocument + 逐页绘制）。

重写的原因（旧版实测的缺陷）：
- 旧版是一条无限滚动的文本流，没有纸张 / 页边距 / 页眉 / 页码 / 分页。
  学位论文的版式判断完全依赖页面，所以它看起来不像论文。
- 旧版把"固定 22 磅"换算成"相对字号的百分比"，而 Qt 的百分比是相对
  字体自然行高算的，行距系统性偏大约 20%；左右两栏还各用一套公式。
- 旧版只取中文字体，设置里的 ascii 字体从头到尾没被用过。
- 旧版最多只画 400 段，长文档的预览是残缺的。

单位约定：QTextDocument 的布局单位 = "点"。字号统一用 setPixelSize(pt)
强制 1 单位 = 1 点，避免高 DPI 下 Qt 按 96/120dpi 换算导致分页漂移。
"""
from __future__ import annotations

from pathlib import Path

import re

from docx.oxml.ns import qn

from PySide6.QtCore import QSize, QSizeF, Qt, QUrl
from PySide6.QtGui import (
    QColor,
    QTextFormat,
    QTextFrameFormat,
    QFont,
    QFontMetricsF,
    QImage,
    QTextBlock,
    QTextBlockFormat,
    QTextCharFormat,
    QTextCursor,
    QTextDocument,
    QTextImageFormat,
    QTextLength,
    QTextTableCellFormat,
    QTextTableFormat,
)

from ..classify import effective_role
from ..labels import LABEL_ROLES
from ..numbering import (
    clean_caption,
    clean_heading,
    plan_caption_numbers,
    plan_heading_numbers,
)

def _fixed_height_kind() -> int:
    """QTextBlockFormat.FixedHeight 的枚举值。

    必须走 FixedHeight：ProportionalHeight 的百分比基准是字体自然行高
    （实测 lineSpacing，不是 height()），换字体就漂，固定磅值根本对不上。
    """
    try:
        return int(QTextBlockFormat.LineHeightTypes.FixedHeight.value)
    except Exception:
        pass
    try:
        return int(QTextBlockFormat.FixedHeight)
    except Exception:
        return 2


_FIXED = _fixed_height_kind()


MM_TO_PT = 72.0 / 25.4
PAPER_SIZES = {
    "A4": (595.276, 841.890),
    "Letter": (612.0, 792.0),
}

MAX_PARAS = 6000        # 安全上限，不再让正常论文被截断
MAX_TABLES = 120
MAX_TABLE_ROWS = 25
MAX_TABLE_CELL = 80

INK_COLOR = QColor("#1d1d1f")
FALLBACK_CJK = "宋体"
FALLBACK_LATIN = "Times New Roman"


# ------------------------------------------------------------------ #
# 页面几何
# ------------------------------------------------------------------ #

def page_geometry(settings):
    """返回 (纸张 pt, (上右下左) pt, 正文区 pt)。"""
    pg = settings.page
    paper = PAPER_SIZES.get(getattr(settings, "paper", "A4"), PAPER_SIZES["A4"])
    t = float(pg.get("top", 25.0)) * MM_TO_PT
    b = float(pg.get("bottom", 25.0)) * MM_TO_PT
    l = float(pg.get("left", 30.0)) * MM_TO_PT
    r = float(pg.get("right", 30.0)) * MM_TO_PT
    return (paper[0], paper[1]), (t, r, b, l), (paper[0] - l - r, paper[1] - t - b)


# ------------------------------------------------------------------ #
# 字体
# ------------------------------------------------------------------ #

_FONT_CACHE: dict[tuple, QFont] = {}
_CF_CACHE: dict[tuple, QTextCharFormat] = {}
_METRIC_CACHE: dict[tuple, float] = {}


def _font(family: str, size: float, bold: bool) -> QFont:
    key = (family or FALLBACK_LATIN, int(round(size)), bool(bold))
    f = _FONT_CACHE.get(key)
    if f is None:
        f = QFont()
        f.setFamily(key[0])
        # 1 单位 = 1 点，见模块 docstring
        f.setPixelSize(max(1, key[1]))
        f.setWeight(QFont.Weight.Bold if key[2] else QFont.Weight.Normal)
        _FONT_CACHE[key] = f
    return f


def _char_format(family: str, size: float, bold: bool,
                 color: str = "") -> QTextCharFormat:
    key = (family or FALLBACK_LATIN, int(round(size)), bool(bold), color or "")
    cf = _CF_CACHE.get(key)
    if cf is None:
        cf = QTextCharFormat()
        cf.setFont(_font(family, size, bold))
        if color:
            try:
                cf.setForeground(QColor("#" + color))
            except Exception:
                pass
        _CF_CACHE[key] = cf
    return cf


def _line_pt(rule: str | None, value: float, natural: float) -> float:
    """把设置里的行距规则换算成"磅"（预览的 FixedHeight 单位）。"""
    if rule == "single":
        return natural
    if rule == "multiple":
        return 1.5 * natural
    if rule == "multiple2":
        return 2.0 * natural
    if rule == "minimum":
        return max(float(value), natural)
    return float(value)          # exact / 默认


def _natural_pt(family: str, size: float) -> float:
    """该字体在该字号下的单倍行距（布局单位）。

    用 lineSpacing() 而不是 height()：前者才是 Word"单倍行距"的对应量。
    QFontMetricsF 构造很贵，按 (字体, 字号) 缓存——全文档通常只有几种组合。
    """
    key = (family or FALLBACK_LATIN, int(round(size)))
    v = _METRIC_CACHE.get(key)
    if v is None:
        try:
            m = QFontMetricsF(_font(family, size, False))
            v = max(1.0, m.lineSpacing() or m.height())
        except Exception:
            v = size * 1.2
        _METRIC_CACHE[key] = v
    return v


def _is_cjk(ch: str) -> bool:
    o = ord(ch)
    return (
        0x2E80 <= o <= 0x9FFF        # CJK 部首 .. 统一表意文字
        or 0xF900 <= o <= 0xFAFF     # 兼容表意文字
        or 0x20000 <= o <= 0x2FA1F   # 扩展 B 及以上
        or 0x3000 <= o <= 0x303F     # CJK 标点
        or 0xFF00 <= o <= 0xFFEF     # 全角形式
    )


_CJK_CLASS = ("一-鿿"        # CJK 部首 .. 统一表意文字
              "豈-﫿"        # 兼容表意文字
              "\U00020000-\U0002FA1F"   # 扩展 B 及以上
              "　-〿"        # CJK 标点
              "＀-￯"        # 全角形式
              )
RE_SEG = re.compile(f"([{_CJK_CLASS}]+|[^{_CJK_CLASS}]+)")


def _segments(text: str):
    """按"中文字体 / 西文字体"把文本切成段（正则切分，避免逐字符 Python 循环）。"""
    out = []
    for seg in RE_SEG.findall(text):
        if seg:
            out.append((_is_cjk(seg[0]), seg))
    return out


def _align_flag(align: str | None) -> Qt.Alignment:
    return {
        "left": Qt.AlignLeft,
        "center": Qt.AlignHCenter,
        "right": Qt.AlignRight,
        "both": Qt.AlignJustify,
    }.get(align or "left", Qt.AlignLeft)


# ------------------------------------------------------------------ #
# 块序列
# ------------------------------------------------------------------ #

def iter_blocks(doc) -> list[tuple[str, int]]:
    out: list[tuple[str, int]] = []
    p_i = 0
    t_i = 0
    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            out.append(("p", p_i))
            p_i += 1
        elif child.tag == qn("w:tbl"):
            out.append(("tbl", t_i))
            t_i += 1
    return out


_MC_NS = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
_MC_ALT = _MC_NS + "AlternateContent"
_MC_CHOICE = _MC_NS + "Choice"
_MC_FALLBACK = _MC_NS + "Fallback"
_O_RELID = "{urn:schemas-microsoft-com:office:office}relid"


def _drawing_size(el) -> tuple[float, float] | None:
    for ex in el.iter(qn("wp:extent")):
        cx, cy = ex.get("cx"), ex.get("cy")
        if cx and cy:
            try:
                return int(cx) / 12700.0, int(cy) / 12700.0
            except ValueError:
                pass
    return None


def _rel_blob(doc, rid: str | None) -> bytes | None:
    if not rid:
        return None
    try:
        part = doc.part.related_parts.get(rid)
        return part.blob if part is not None else None
    except Exception:
        return None


_W_DRAWING = qn("w:drawing")
_W_PICT = qn("w:pict")
_W_OBJ = qn("w:object")
_A_BLIP = qn("a:blip")
_R_EMBED = qn("r:embed")
_W_T = qn("w:t")


def element_images(doc, root_el) -> list[tuple[bytes, tuple[float, float] | None]]:
    """提取一个 XML 节点（段落或表格单元格）里的所有插图 (blob, 原始尺寸pt)。

    兼容两种 Word 存图方式：
    - DrawingML（w:drawing → a:blip），现代 docx 默认；
    - VML（w:pict → v:imagedata），老版 .doc 转 .docx 的封面校徽/插图多走这一路。
    跳过 w:object（MathType OLE 公式由公式通路画）和 mc:AlternateContent 里重复的 mc:Fallback。
    """
    # 快速短路：绝大多数表格单元格/段落不含图片，C 层查找可省去 Python 递归遍历
    if root_el.find(".//" + _W_DRAWING) is None and root_el.find(".//" + _W_PICT) is None:
        return []

    from ..features import vml_size_pt

    out: list[tuple[bytes, tuple[float, float] | None]] = []

    def walk(el) -> None:
        tag = el.tag
        if tag == _W_OBJ:
            return
        if tag == _MC_ALT:
            choice = el.find(_MC_CHOICE)
            fallback = el.find(_MC_FALLBACK)
            before = len(out)
            if choice is not None:
                walk(choice)
            if len(out) == before and fallback is not None:
                walk(fallback)
            return
        if tag == _W_DRAWING:
            sz = _drawing_size(el)
            for blip in el.iter(_A_BLIP):
                blob = _rel_blob(doc, blip.get(_R_EMBED))
                if blob:
                    out.append((blob, sz))
            return
        if tag == _W_PICT:
            for im in el.iter(_VML_IMG):
                rid = im.get(_RID) or im.get(_O_RELID)
                blob = _rel_blob(doc, rid)
                if blob:
                    out.append((blob, vml_size_pt(im)))
            return
        for ch in el:
            walk(ch)

    walk(root_el)
    return out


def paragraph_images(doc, paragraph) -> list[tuple[bytes, tuple[float, float] | None]]:
    return element_images(doc, paragraph._p)


def paragraph_image_blob(doc, paragraph) -> bytes | None:
    imgs = paragraph_images(doc, paragraph)
    return imgs[0][0] if imgs else None


# ------------------------------------------------------------------ #
# 行距
# ------------------------------------------------------------------ #

def _line_pt_before(rule: str | None, value: float | None, natural: float) -> float:
    """原稿侧：把 Word 的 spacing 规则换算成"点"。

    features 里 exactly/atLeast 的 value 已经是磅；auto/multiple 的 value 是倍数。
    """
    if rule in ("exactly", "at_least"):
        return float(value) if value else natural
    if rule == "multiple":
        return float(value or 1.0) * natural
    return natural             # auto / 未写


# ------------------------------------------------------------------ #
# 插入
# ------------------------------------------------------------------ #

def _insert_runs(cur: QTextCursor, text: str, cjk: str, latin: str,
                 size: float, bold: bool, color: str = "",
                 marks: bool = False) -> None:
    if marks:
        from ..marks import split_marks
        for seg_text, mk in split_marks(text):
            sz = size * 0.7 if mk else size
            for is_cjk, seg in _segments(seg_text):
                cf = _char_format(cjk if is_cjk else latin, sz, bold, color)
                if mk > 0:
                    cf.setVerticalAlignment(
                        QTextCharFormat.VerticalAlignment.AlignSuperScript)
                elif mk < 0:
                    cf.setVerticalAlignment(
                        QTextCharFormat.VerticalAlignment.AlignSubScript)
                cur.insertText(seg, cf)
        return
    for is_cjk, seg in _segments(text):
        cur.insertText(seg, _char_format(cjk if is_cjk else latin, size, bold, color))


def _para_breaks(p_el) -> tuple[bool, bool, bool]:
    """检查段落中的显式分页/分节符。

    返回 (break_before, break_after, has_page_or_sect_break)：
    - w:pPr/w:pageBreakBefore（val 非 0/false/off）→ 段前分页
    - w:pPr/w:sectPr（type 非 continuous）→ 段后分节换页
    - w:r/w:br[@w:type="page"] → 按在段落中相对可见内容的位置判定段前或段后分页
    """
    break_before = False
    break_after = False
    ppr = p_el.find(qn("w:pPr"))
    if ppr is not None:
        pbb = ppr.find(qn("w:pageBreakBefore"))
        if pbb is not None and (pbb.get(qn("w:val")) or "").lower() not in (
            "0", "false", "off"
        ):
            break_before = True
        sp = ppr.find(qn("w:sectPr"))
        if sp is not None:
            tp = sp.find(qn("w:type"))
            sv = (tp.get(qn("w:val")) or "").lower() if tp is not None else ""
            if sv != "continuous":
                break_after = True

    w_br = qn("w:br")
    w_type = qn("w:type")
    if p_el.find(".//" + w_br) is not None:
        w_t = _W_T
        content_tags = (_W_DRAWING, _W_PICT, _W_OBJ, qn("m:oMath"), qn("m:oMathPara"))
        seen_content = False
        seen_br_before_content = False
        has_any_br = False
        for el in p_el.iter():
            tag = el.tag
            if tag == w_br and (el.get(w_type) or "").lower() == "page":
                has_any_br = True
                if seen_content:
                    break_after = True
                else:
                    seen_br_before_content = True
            elif tag == w_t and (el.text or "").strip():
                seen_content = True
            elif tag in content_tags:
                seen_content = True
        if has_any_br:
            if seen_br_before_content and seen_content:
                break_before = True
            elif not seen_content:
                break_after = True

    return break_before, break_after, (break_before or break_after)


def _insert_paragraph(cur: QTextCursor, text: str, cjk: str, latin: str,
                      size: float, bold: bool, line_pt: float,
                      align: str | None, indent_pt: float = 0.0,
                      color: str = "", marks: bool = False,
                      label_bold: bool = False,
                      left_margin: float = 0.0,
                      page_break_before: bool = False,
                      reuse_first: bool = False) -> None:
    bf = QTextBlockFormat()
    bf.setAlignment(_align_flag(align))
    # 直接给磅值：FixedHeight 的单位就是布局单位（1 单位 = 1 磅）
    bf.setLineHeight(max(1.0, line_pt), _FIXED)
    bf.setLeftMargin(max(0.0, left_margin))
    bf.setRightMargin(0)
    if indent_pt:
        bf.setTextIndent(indent_pt)
    if page_break_before:
        bf.setPageBreakPolicy(QTextFormat.PageBreakFlag.PageBreak_AlwaysBefore)
    if reuse_first:
        cur.setBlockFormat(bf)
    else:
        cur.insertBlock(bf)
    if not text:
        cur.setBlockCharFormat(_char_format(cjk, size, bold, color))
        return
    # 段落开头的 "Abstract:" / "关键词:" 只加粗标签那几个字
    if label_bold:
        from ..labels import split_label
        label, _ = split_label(text)
        if label is not None:
            _insert_runs(cur, label, cjk, latin, size, True, color, False)
            _insert_runs(cur, text[len(label):], cjk, latin, size, bold,
                         color, marks)
            return
    _insert_runs(cur, text, cjk, latin, size, bold, color, marks)


def _load_image(blob: bytes, content_w: float,
                natural: tuple[float, float] | None = None) -> QImage | None:
    """解码 + 折算 devicePixelRatio（见 _insert_image 的说明）。"""
    if natural:
        im = _scaled_image(blob, content_w * 6, keep_hi=True)
    else:
        im = _scaled_image(blob, content_w)
    if im is None:
        return None
    if natural and im.width() > 0:
        w_pt = max(1.0, float(natural[0]))
        h_pt = float(natural[1]) if natural[1] else w_pt
        if w_pt > content_w:               # 比栏宽还宽就缩到栏宽
            h_pt *= content_w / w_pt
            w_pt = content_w
        im.setDevicePixelRatio(im.width() / w_pt)
    return im


def _insert_image(cur: QTextCursor, blob: bytes, content_w: float,
                  natural: tuple[float, float] | None = None,
                  align: str | None = "center") -> None:
    """插入图片（默认居中，单独一行）。"""
    try:
        im = _load_image(blob, content_w, natural)
        if im is None:
            return
        bf = QTextBlockFormat()
        bf.setAlignment(_align_flag(align or "center"))
        cur.insertBlock(bf)
        cur.insertImage(im)
    except Exception:
        pass


def _insert_equation(cur: QTextCursor, im, number: str,
                     content_w: float, latin: str, size: float, bold: bool,
                     color: str) -> None:
    """公式 + 编号：编号右对齐，和公式同一行（Word 的排法）。

    im 可以是单张图，也可以是同段里的多张（按顺序横排）。
    用 1 行 2 列的无边框表格：左列居中放公式，右列右对齐放编号。
    没有编号就只画公式。
    """
    ims = im if isinstance(im, (list, tuple)) else [im]
    if not number:
        bf = QTextBlockFormat()
        bf.setAlignment(Qt.AlignHCenter)
        cur.insertBlock(bf)
        for one in ims:
            cur.insertImage(one)
        return
    try:
        num_w = 80.0
        tf = QTextTableFormat()
        tf.setBorder(0)
        tf.setCellPadding(2)
        tf.setCellSpacing(0)
        tf.setWidth(QTextLength(QTextLength.FixedLength, content_w))
        tf.setColumnWidthConstraints([
            QTextLength(QTextLength.FixedLength, max(20.0, content_w - num_w)),
            QTextLength(QTextLength.FixedLength, num_w),
        ])
        t = cur.insertTable(1, 2, tf)
        c0 = t.cellAt(0, 0)
        cur0 = c0.firstCursorPosition()
        bf0 = QTextBlockFormat()
        bf0.setAlignment(Qt.AlignHCenter)
        cur0.setBlockFormat(bf0)
        for one in ims:
            cur0.insertImage(one)

        c1 = t.cellAt(0, 1)
        cur1 = c1.firstCursorPosition()
        bf1 = QTextBlockFormat()
        bf1.setAlignment(Qt.AlignRight)
        cur1.setBlockFormat(bf1)
        cur1.insertText(number, _char_format(latin, size, bold, color))
        _leave_table(cur, t)
    except Exception:
        bf = QTextBlockFormat()
        bf.setAlignment(Qt.AlignHCenter)
        cur.insertBlock(bf)
        for one in ims:
            cur.insertImage(one)


_EQ_NUM_RE = re.compile(r"^[\(（]\s*[\w.\-–—]+\s*[\)）][.。]?$")
_VML_IMG = "{urn:schemas-microsoft-com:vml}imagedata"
_RID = ("{http://schemas.openxmlformats.org/officeDocument/2006/"
        "relationships}id")


def _has_field(p) -> bool:
    return (p._p.find(".//" + qn("w:fldChar")) is not None
            or p._p.find(".//" + qn("w:fldSimple")) is not None)


def _object_natural(obj) -> tuple[float, float] | None:
    """w:object 的原始尺寸（磅），来自 dxaOrig/dyaOrig（twips）。"""
    w, h = obj.get(qn("w:dxaOrig")), obj.get(qn("w:dyaOrig"))
    if not w or not h:
        return None
    try:
        return int(w) / 20.0, int(h) / 20.0
    except ValueError:
        return None


def _object_blob(p, obj) -> bytes | None:
    """某个 w:object 自己的公式图片（v:imagedata 的 r:id）。"""
    for im in obj.iter(_VML_IMG):
        rid = im.get(_RID)
        if not rid:
            continue
        try:
            part = p.part.related_parts.get(rid)
            return part.blob if part is not None else None
        except Exception:
            return None
    return None


def _formula_images(p, content_w: float, size: float, latin: str,
                    bold: bool) -> list[QImage]:
    """把段落里的公式按文档顺序渲染成图片：OMML 走自绘，MathType 走 EMF。

    一段里可以有多条（行内公式常见「式(1)和式(2)」），顺序和段落一致。
    """
    out: list[QImage] = []
    m_math = qn("m:oMath")
    m_para = qn("m:oMathPara")
    w_obj = qn("w:object")
    for el in p._p.iter():
        if el.tag in (m_math, m_para):
            par = el.getparent()
            if el.tag == m_math and par is not None and par.tag == m_para:
                continue            # oMathPara 是容器，只画外层
            from ..omml import render_element
            try:
                im = render_element(el, size, latin, bold)
            except Exception:      # noqa: BLE001  单条公式画不出来不影响整页
                im = None
            if im is not None:
                out.append(im)
        elif el.tag == w_obj:
            blob = _object_blob(p, el)
            if not blob:
                continue
            from ..emf import emf_to_png
            natural = _object_natural(el)
            target = int(natural[0] * 6) if natural else int(content_w * 6)
            png = emf_to_png(blob, max(64, target))
            if png is None:
                continue
            im = _load_image(Path(png).read_bytes(), content_w, natural)
            if im is not None:
                out.append(im)
    return out


def _standalone_formula(p, f, n_images: int) -> bool:
    """这段公式是「独占一行的显示公式」还是「句子里的行内公式」。

    显示公式：段里只有公式对象 + 编号（"(1)" / "(4-1)"，编号常是 Word 域），
    这类要按「中间公式 + 右侧编号」排。
    行内公式：公式前后还有正文，必须留在同一行里。
    """
    txt = (f.text or "").replace("\t", " ").strip()
    if not txt:
        return True
    if n_images != 1:
        return False
    if _EQ_NUM_RE.match(txt):
        return True
    # 有公式编号域时，段里的文字可能只剩编号（域结果），也算显示公式
    if _has_field(p):
        from ..features import equation_label
        label = equation_label(p).replace("\t", " ").strip()
        return bool(label) and txt == label and len(txt) <= 12
    return False


def _run_line_text(r) -> str:
    """一个 w:r 里能显示成文字的部分（w:t / 制表符 / 换行）。"""
    out = []
    for ch in r:
        if ch.tag == qn("w:t"):
            out.append(ch.text or "")
        elif ch.tag == qn("w:tab"):
            out.append("    ")
        elif ch.tag in (qn("w:br"), qn("w:cr")):
            out.append(" ")
    return "".join(out)


_INLINE_CONTAINERS = {qn(t) for t in (
    "w:hyperlink", "w:smartTag", "w:ins", "w:sdt", "w:sdtContent",
    "w:fldSimple", "w:customXml")}


def _insert_inline_formula(cur: QTextCursor, p, ims: list[QImage],
                           cjk: str, latin: str, size: float, bold: bool,
                           color: str) -> None:
    """行内公式：按原段落顺序把「文字 → 公式图 → 文字」铺进同一段落。"""
    state = {"i": 0, "used": False}

    def walk(el) -> None:
        for ch in el:
            tag = ch.tag
            if tag in (qn("m:oMath"), qn("m:oMathPara")):
                if state["i"] < len(ims):
                    cur.insertImage(ims[state["i"]])
                    state["used"] = True
                state["i"] += 1
                continue
            if tag == qn("w:r"):
                if ch.find(".//" + qn("w:object")) is not None:
                    if state["i"] < len(ims):
                        cur.insertImage(ims[state["i"]])
                        state["used"] = True
                    state["i"] += 1
                else:
                    txt = _run_line_text(ch)
                    if txt:
                        _insert_runs(cur, txt, cjk, latin, size, bold, color)
                continue
            if tag in _INLINE_CONTAINERS:
                walk(ch)

    walk(p._p)
    if not state["used"]:
        for im in ims:
            cur.insertImage(im)
            return


_IMG_CACHE: dict[tuple, QImage] = {}


def _scaled_image(blob: bytes, content_w: float, keep_hi: bool = False) -> QImage | None:
    """解码并按栏宽缩放。

    解码是刷新里最贵的一步（实测 14 张 4000px 大图 = 0.96 秒，占整个刷新的 2/3），
    而同一份文档每次刷新拿到的都是同样的图，所以按 (内容, 目标宽度) 缓存。
    keep_hi=True 时不缩到栏宽（公式/插图要保留多倍像素，清晰度优先）。
    """
    if not blob:
        return None
    key = (hash(blob), int(content_w), bool(keep_hi))
    hit = _IMG_CACHE.get(key)
    if hit is not None:
        return hit
    im = QImage()
    im.loadFromData(blob)
    if im.isNull():
        # Qt 原生不解码 EMF/WMF（老版 Word 常见矢量校徽/图），走 GDI+ 转 PNG
        from ..emf import emf_to_png
        target = max(128, int(content_w * (4 if keep_hi else 2)))
        png = emf_to_png(blob, target)
        if png is not None:
            im = QImage()
            im.loadFromData(Path(png).read_bytes())
    if im.isNull():
        return None
    if not keep_hi and im.width() > content_w:
        im = im.scaledToWidth(int(content_w), Qt.SmoothTransformation)
    if len(_IMG_CACHE) > 64:
        _IMG_CACHE.clear()
    _IMG_CACHE[key] = im
    return im


def _leave_table(cur: QTextCursor, table) -> None:
    """insertTable() 会把光标留在单元格内，必须显式移出，
    否则后面的段落/线会画进单元格里。"""
    try:
        cur.setPosition(table.lastPosition() + 1)
    except Exception:
        cur.movePosition(QTextCursor.End)


_RULE_URL = "rule://ink"


def _ensure_rule_resource(td: QTextDocument) -> None:
    if td is None:
        return
    if not td.resource(QTextDocument.ResourceType.ImageResource, QUrl(_RULE_URL)):
        im = QImage(QSize(4, 4), QImage.Format.Format_ARGB32)
        im.fill(INK_COLOR)
        td.addResource(QTextDocument.ResourceType.ImageResource,
                       QUrl(_RULE_URL), im)


def _rule_block(cur: QTextCursor, content_w: float, width_pt: float,
                w: float | None = None, off: float = 0.0,
                page_break_before: bool = False,
                reuse_empty: bool = True) -> None:
    """画一条横线：用浮点高度的 QTextImageFormat 拉伸纯色资源。

    相比整像素 QImage（1px 在缩放 60%~75% 时容易被抽丢、0.75pt 与 1.5pt 拉不开差距），
    QTextImageFormat 支持精确浮点线宽，在各种缩放比下都能稳定区分粗细线。
    """
    w = float(w) if w else content_w
    h_pt = max(0.75, float(width_pt))
    _ensure_rule_resource(cur.document())
    fmt = QTextImageFormat()
    fmt.setName(_RULE_URL)
    fmt.setWidth(max(2.0, w))
    fmt.setHeight(h_pt)
    bf = QTextBlockFormat()
    bf.setAlignment(Qt.AlignLeft)
    bf.setLeftMargin(off)
    bf.setRightMargin(max(0.0, content_w - off - w))
    bf.setLineHeight(max(2.0, h_pt + 0.8), _FIXED)
    if page_break_before:
        bf.setPageBreakPolicy(QTextFormat.PageBreakFlag.PageBreak_AlwaysBefore)
    # 表格/frame 后面通常已经留了一个空块，直接复用——否则线会和
    # 表格之间空出一行（实测约 16 磅）
    blk = cur.block()
    if reuse_empty and blk.isValid() and not blk.text().strip():
        cur.setBlockFormat(bf)
        cur.insertImage(fmt)
    else:
        cur.insertBlock(bf)
        cur.insertImage(fmt)


def _insert_cell_text(cursor: QTextCursor, text: str, cjk: str, latin: str,
                      size: float, bold: bool, color: str) -> None:
    for is_cjk, seg in _segments(text):
        cursor.insertText(seg, _char_format(cjk if is_cjk else latin,
                                            size, bold, color))


def _leave_frame(cur: QTextCursor, frame) -> None:
    """把光标移出 QTextFrame（表格要包在 frame 里才能左右偏移）。"""
    try:
        cur.setPosition(frame.lastPosition() + 1)
    except Exception:
        cur.movePosition(QTextCursor.End)


def _frame_table(cur: QTextCursor, rows: list, cols: int, content_w: float,
                 line_pt: float, cjk: str, latin: str, size: float, bold: bool,
                 color: str, align: str, w: float, off: float,
                 border_pt: float = 0.0,
                 cell_imgs: list | None = None,
                 first_row_bold: bool | None = None,
                 page_break_before: bool = False) -> None:
    """把一个表格包进无边框 frame 里，用 frame 的左右边距控制表格位置。"""
    col_w = w / max(1, cols)
    tf = QTextTableFormat()
    if border_pt > 0:
        # Qt 的 QTextTableFormat.setBorder 小于 1.0 会被截成 0 而不画线，所以兜底 >= 1.0
        tf.setBorder(max(1.0, float(border_pt)))
        tf.setBorderStyle(QTextFrameFormat.BorderStyle.BorderStyle_Solid)
        tf.setBorderBrush(INK_COLOR)
    else:
        tf.setBorder(0)
    tf.setCellPadding(3)
    tf.setCellSpacing(0)
    tf.setWidth(QTextLength(QTextLength.FixedLength, w))
    tf.setColumnWidthConstraints(
        [QTextLength(QTextLength.FixedLength, col_w)] * cols
    )
    ff = QTextFrameFormat()
    ff.setBorder(0)
    ff.setPadding(0)
    ff.setLeftMargin(off)
    ff.setRightMargin(max(0.0, content_w - off - w))
    if page_break_before:
        ff.setPageBreakPolicy(QTextFormat.PageBreakFlag.PageBreak_AlwaysBefore)
    fr = cur.insertFrame(ff)
    cur.setPosition(fr.firstPosition())
    t = cur.insertTable(len(rows), cols, tf)
    img_max_w = max(20.0, col_w - 8.0)
    for r, texts in enumerate(rows):
        row_bold = (first_row_bold if (r == 0 and first_row_bold is not None)
                    else bold)
        for c in range(cols):
            cell = t.cellAt(r, c)
            cfmt = QTextTableCellFormat()
            cfmt.setFont(_font(cjk, size, row_bold))
            cell.setFormat(cfmt)
            cc = cell.firstCursorPosition()
            bf = QTextBlockFormat()
            bf.setAlignment(_align_flag(align))
            cc.setBlockFormat(bf)
            txt = texts[c] if c < len(texts) else ""
            if txt:
                _insert_cell_text(cc, txt, cjk, latin, size, row_bold, color)
            if cell_imgs and r < len(cell_imgs) and c < len(cell_imgs[r]):
                for blob, nat in cell_imgs[r][c]:
                    im = _load_image(blob, img_max_w, nat)
                    if im is not None:
                        if txt:
                            cc.insertBlock(bf)
                        cc.insertImage(im)
    _leave_frame(cur, fr)


def _one_row_table(cur: QTextCursor, texts: list, content_w: float,
                   line_pt: float, cjk: str = FALLBACK_CJK,
                   latin: str = FALLBACK_LATIN, size: float = 10.5,
                   bold: bool = False, color: str = "",
                   align: str = "left", w: float | None = None,
                   off: float = 0.0, cell_imgs: list | None = None,
                   page_break_before: bool = False) -> None:
    w = float(w) if w else content_w
    _frame_table(cur, [list(texts)], len(texts), content_w, line_pt,
                 cjk, latin, size, bold, color, align, w, off,
                 border_pt=0.0,
                 cell_imgs=[cell_imgs] if cell_imgs is not None else None,
                 page_break_before=page_break_before)


def _orig_table_style(table) -> tuple[str, float, float]:
    """推断原稿表格的框线样式（供原稿侧预览用），避免套用右边设置把原稿框线隐藏。"""
    tblPr = getattr(table._tbl, "tblPr", None)
    b = tblPr.find(qn("w:tblBorders")) if tblPr is not None else None
    if b is None:
        return "grid", 1.0, 0.75
    vals = {}
    for side in ("top", "bottom", "left", "right", "insideH", "insideV"):
        el = b.find(qn("w:" + side))
        v = (el.get(qn("w:val")) or "").lower() if el is not None else ""
        vals[side] = v not in ("", "none", "nil")
    has_h = vals["top"] or vals["bottom"]
    has_v = vals["left"] or vals["right"] or vals["insideV"]
    if has_h and not has_v and not vals["insideH"]:
        return "three_line", 1.5, 0.75
    if any(vals.values()):
        return "grid", 1.0, 0.75
    # 某些表只在单元格 tcBorders 上画线
    for row in table.rows[:2]:
        for cell in row.cells[:4]:
            tcPr = cell._tc.find(qn("w:tcPr"))
            tb = tcPr.find(qn("w:tcBorders")) if tcPr is not None else None
            if tb is not None:
                for ch in tb:
                    if (ch.get(qn("w:val")) or "").lower() not in ("", "none", "nil"):
                        return "grid", 1.0, 0.75
    return "none", 0.0, 0.0


def table_width_pt(table, content_w: float) -> float:
    """原稿表格的宽度（磅）。读 w:tblW；没写就按栏宽。"""
    tblPr = getattr(table._tbl, "tblPr", None)
    if tblPr is None:
        return content_w
    el = tblPr.find(qn("w:tblW"))
    if el is None:
        return content_w
    try:
        v = int(el.get(qn("w:w")))
    except (TypeError, ValueError):
        return content_w
    t = (el.get(qn("w:type")) or "").lower()
    if t == "pct":
        return content_w * max(0.05, min(1.0, v / 5000.0))
    if t == "auto":
        return content_w
    return v / 20.0          # dxa(twips) → 磅


def _insert_table(doc, cur: QTextCursor, t_i: int, content_w: float,
                  line_pt: float = 1.5, cjk: str = FALLBACK_CJK,
                  latin: str = FALLBACK_LATIN, size: float = 10.5,
                  bold: bool = False, color: str = "", align: str = "left",
                  pos: str = "center", width_pt: float | None = None,
                  style: str = "three_line",
                  line_inner: float = 0.75,
                  header_bold: bool | None = None,
                  table=None,
                  page_break_before: bool = False,
                  reuse_first: bool = False) -> None:
    """预览表格：支持三线表（两粗一细）、网格表、无线表，并渲染单元格里的图片。

    三线表为什么拆分绘制：PySide6 的 QTextTableCellFormat 分边边框不绘制，
    所以拆成「顶线(粗) + 表头行 + 栏目线(细) + 数据行 + 底线(粗)」，列宽固定以对齐。
    网格表直接用 QTextTableFormat.setBorder() 画全框线。
    表格位置（pos）靠外层 QTextFrame 的左右边距控制。
    """
    if table is None:
        try:
            table = doc.tables[t_i]
        except IndexError:
            return
    tr_lst = table._tbl.tr_lst[:MAX_TABLE_ROWS]
    if not tr_lst:
        return
    # 直接按 XML tr/tc 读单元格文本与插图，避开 python-docx row.cells 网格展开的 O(R×C) 开销
    row_tcs = [tr.tc_lst for tr in tr_lst]
    cols = max(1, max(len(tcs) for tcs in row_tcs))
    w = max(20.0, min(float(width_pt or content_w), content_w))
    off = {"left": 0.0,
           "center": (content_w - w) / 2.0,
           "right": content_w - w}.get(pos, 0.0)
    h_bold = bold if header_bold is None else bool(header_bold)

    all_texts: list[list[str]] = []
    all_imgs: list[list[list]] = []
    for tcs in row_tcs:
        r_txts: list[str] = []
        r_imgs: list[list] = []
        for tc in tcs:
            txt = "".join(t.text or "" for t in tc.iter(_W_T)).strip().replace("\n", " ")[:MAX_TABLE_CELL]
            r_txts.append(txt)
            r_imgs.append(element_images(doc, tc))
        while len(r_txts) < cols:
            r_txts.append("")
            r_imgs.append([])
        all_texts.append(r_txts)
        all_imgs.append(r_imgs)

    if style == "grid":
        bw = max(1.0, float(line_inner or line_pt or 1.0))
        _frame_table(cur, all_texts, cols, content_w, line_pt, cjk, latin,
                     size, bold, color, align, w, off,
                     border_pt=bw, cell_imgs=all_imgs,
                     first_row_bold=h_bold,
                     page_break_before=page_break_before)
        cur.insertBlock()
        return

    if style == "none":
        _frame_table(cur, all_texts, cols, content_w, 0.0, cjk, latin,
                     size, bold, color, align, w, off,
                     border_pt=0.0, cell_imgs=all_imgs,
                     first_row_bold=h_bold,
                     page_break_before=page_break_before)
        cur.insertBlock()
        return

    # three_line：两粗一细（顶/底线 outer，栏目线 inner）
    outer = float(line_pt or 0)
    inner = float(line_inner or 0) if float(line_inner or 0) > 0 else outer
    if outer > 0:
        # 顶线不复用表前的空段落，避免吞掉正文/封面里的刻意空行
        _rule_block(cur, content_w, outer, w, off,
                    page_break_before=page_break_before,
                    reuse_empty=reuse_first)              # 顶线（粗）
    _one_row_table(cur, all_texts[0], content_w, outer,
                   cjk, latin, size, h_bold, color, align, w, off,
                   cell_imgs=all_imgs[0],
                   page_break_before=(page_break_before and outer <= 0))
    if inner > 0:
        _rule_block(cur, content_w, inner, w, off)        # 栏目线（细）
    if len(all_texts) > 1:
        _frame_table(cur, all_texts[1:], cols, content_w, outer, cjk, latin,
                     size, bold, color, align, w, off,
                     border_pt=0.0, cell_imgs=all_imgs[1:])
    if outer > 0:
        _rule_block(cur, content_w, outer, w, off)        # 底线（粗）
    cur.insertBlock()


# ------------------------------------------------------------------ #
# 主入口
# ------------------------------------------------------------------ #

def render_side(doc, analysis, settings, side: str, content: tuple[float, float],
                body_index: int | None = None, overrides: dict | None = None):
    """渲染一侧。返回 (QTextDocument, {段索引: QTextBlock})。"""
    feats = analysis.features
    blocks_src = iter_blocks(doc)
    content_w, content_h = content

    head_plans: dict[int, str] = {}
    fig_plans: dict[int, str] = {}
    tab_plans: dict[int, str] = {}
    if side == "after":
        if getattr(settings, "enable_number_heading", False):
            head_plans = {p.index: p.text for p in plan_heading_numbers(
                analysis.detections, body_index, settings)}
        if getattr(settings, "enable_number_figure", False):
            fig_plans = {p.index: p.text for p in plan_caption_numbers(
                analysis.detections, "figure", settings, feats)}
        if getattr(settings, "enable_number_table", False):
            tab_plans = {p.index: p.text for p in plan_caption_numbers(
                analysis.detections, "table", settings, feats)}
    marks_on = bool(getattr(settings, "enable_marks", True))
    tbl_style = getattr(settings, "table_style", "three_line") or "three_line"
    table_outer = float(getattr(settings, "table_line", 1.5) or 0)
    table_inner = float(getattr(settings, "table_line_inner", 0.75) or 0)
    tbl_cjk = getattr(settings, "table_font", "") or FALLBACK_CJK
    tbl_latin = getattr(settings, "table_ascii", "") or FALLBACK_LATIN
    tbl_size = float(getattr(settings, "table_font_size", 10.5) or 10.5)
    tbl_bold = bool(getattr(settings, "table_bold", False))
    tbl_header_bold = bool(getattr(settings, "table_header_bold", tbl_bold))
    tbl_color = getattr(settings, "table_color", "") or ""
    tbl_align = getattr(settings, "table_align", "left") or "left"
    tbl_pos = getattr(settings, "table_pos", "center") or "center"

    td = QTextDocument()
    td.setPageSize(QSizeF(content_w, content_h))
    td.setDocumentMargin(0)
    td.setDefaultFont(_font(FALLBACK_LATIN, 12, False))

    # 一次性快照段落/表格与样式索引，避免循环内反复构建 python-docx 列表和扫 styles.xml
    paras = doc.paragraphs
    tables = doc.tables
    style_idx = None
    inspect_tbl = None
    if side == "before" and tables:
        from ..features import StyleIndex
        from ..recognize import inspect_table_format
        style_idx = StyleIndex(doc)
        inspect_tbl = inspect_table_format

    cur = QTextCursor(td)
    block_of: dict[int, QTextBlock] = {}
    shown = 0
    shown_tbl = 0
    started = False
    pending_break = False

    # 裁掉文末连续的无意义空段，避免在文档末尾平白多出空白页；
    # 文档开头与中间的空行（如封面留白、段落间隔）全部如实保留。
    last_content = -1
    for idx_b in range(len(blocks_src) - 1, -1, -1):
        k_b, i_b = blocks_src[idx_b]
        if k_b == "tbl":
            last_content = idx_b
            break
        fb = feats[i_b]
        if fb.text.strip() or fb.has_image or fb.has_math or fb.has_ole:
            last_content = idx_b
            break
    blocks_trim = blocks_src[:last_content + 1] if last_content >= 0 else []

    for kind, i in blocks_trim:
        if kind == "tbl":
            if shown_tbl < MAX_TABLES and i < len(tables):
                tbl = tables[i]
                pb_tbl = pending_break and started
                if side == "before":
                    st, ow, iw = _orig_table_style(tbl)
                    orig = inspect_tbl(doc, [tbl], style_index=style_idx)
                    _insert_table(
                        doc, cur, i, content_w, ow,
                        orig.get("font") or FALLBACK_CJK,
                        orig.get("ascii") or FALLBACK_LATIN,
                        float(orig.get("size") or 10.5),
                        bool(orig.get("bold")), "",
                        orig.get("align") or "left",
                        orig.get("pos") or "center",
                        table_width_pt(tbl, content_w),
                        style=st, line_inner=iw,
                        header_bold=bool(orig.get("header_bold")),
                        table=tbl,
                        page_break_before=pb_tbl,
                        reuse_first=not started,
                    )
                else:
                    _insert_table(
                        doc, cur, i, content_w, table_outer,
                        tbl_cjk, tbl_latin, tbl_size, tbl_bold, tbl_color,
                        tbl_align, tbl_pos,
                        table_width_pt(tbl, content_w),
                        style=tbl_style, line_inner=table_inner,
                        header_bold=tbl_header_bold,
                        table=tbl,
                        page_break_before=pb_tbl,
                        reuse_first=not started,
                    )
                shown_tbl += 1
                started = True
                pending_break = False
            continue

        f = feats[i]
        para_obj = paras[i]
        br_before, br_after, has_br = _para_breaks(para_obj._p)
        is_blank = (not f.text.strip() and not f.has_image
                    and not f.has_math and not f.has_ole)

        if is_blank:
            if has_br:
                # 仅含分页符/分节符的空段：切到下一页，不在新页顶部多出一行空白
                if started:
                    if pending_break:
                        # 连续两个纯分页段：保留一张有意的空白页
                        sz_b = f.size_pt or 12.0
                        nat_b = _natural_pt(f.ascii_font or FALLBACK_LATIN, sz_b)
                        lp_b = _line_pt_before(f.line_rule, f.line_value, nat_b)
                        _insert_paragraph(
                            cur, "", f.east_asia or FALLBACK_CJK,
                            f.ascii_font or FALLBACK_LATIN, sz_b, False, lp_b,
                            f.align, 0.0, "", False, False, 0.0,
                            page_break_before=True, reuse_first=False,
                        )
                    pending_break = True
                continue
            if shown >= MAX_PARAS:
                continue
            # 纯空行：原稿侧与排版后侧未改格式（other）均沿用原稿行高，
            # 既保留封面/段间留白，又不把紧凑空行误撑成正文行距
            cjk = f.east_asia or FALLBACK_CJK
            latin = f.ascii_font or FALLBACK_LATIN
            size = f.size_pt or 12.0
            natural = _natural_pt(latin, size)
            line_pt = _line_pt_before(f.line_rule, f.line_value, natural)
            if side == "after":
                d = analysis.detections[i]
                role = effective_role(d)
                ov = (overrides or {}).get(i)
                base = settings.roles.get(role)
                if ov or base:
                    p = dict(base or settings.roles.get("body") or {})
                    if ov:
                        p.update({k: v for k, v in ov.items() if v is not None})
                    size = float(p.get("size", size))
                    natural = _natural_pt(p.get("ascii") or latin, size)
                    line_pt = _line_pt(p.get("line_rule") or "exact",
                                       float(p.get("line", 22)), natural)
            _insert_paragraph(
                cur, "", cjk, latin, size, False, line_pt, f.align,
                0.0, "", False, False, 0.0,
                page_break_before=(pending_break and started),
                reuse_first=not started,
            )
            started = True
            pending_break = False
            continue

        if not f.has_image and shown >= MAX_PARAS:
            continue

        left_margin = 0.0
        if side == "before":
            cjk = f.east_asia or FALLBACK_CJK
            latin = f.ascii_font or FALLBACK_LATIN
            size = f.size_pt or 12.0
            bold = bool(f.bold)
            natural = _natural_pt(latin, size)
            line_pt = _line_pt_before(f.line_rule, f.line_value, natural)
            align = f.align
            indent_pt = float(f.first_indent_pt or 0.0)
            left_margin = float(f.left_indent_pt or 0.0)
            color = ""
            text = f.text
            label_bold = False

        else:
            d = analysis.detections[i]
            role = effective_role(d)
            ov = (overrides or {}).get(i)
            if role == "toc_item" and not ov:
                # 目录条目保留各自层级的字体/字号/加粗/缩进，不压成同一个字号
                cjk = f.east_asia or FALLBACK_CJK
                latin = f.ascii_font or FALLBACK_LATIN
                size = f.size_pt or 12.0
                bold = bool(f.bold)
                natural = _natural_pt(latin, size)
                line_pt = _line_pt_before(f.line_rule, f.line_value, natural)
                align = "left"
                indent_pt = 0.0
                lv = d.level or 1
                left_margin = float(f.left_indent_pt) if f.left_indent_pt else max(0.0, (lv - 1) * 14.0)
                color = ""
                text = f.text
                label_bold = False
            else:
                p = dict(settings.roles.get(role) or settings.roles.get("body") or {})
                if ov:
                    p.update({k: v for k, v in ov.items() if v is not None})
                cjk = p.get("font") or FALLBACK_CJK
                latin = p.get("ascii") or FALLBACK_LATIN
                size = float(p.get("size", 12))
                bold = bool(p.get("bold"))
                natural = _natural_pt(latin, size)
                line_pt = _line_pt(p.get("line_rule") or "exact",
                                   float(p.get("line", 22)), natural)
                color = p.get("color") or ""
                align = p.get("align", "left")
                from ..reflow import indent_chars_of
                indent_pt = (indent_chars_of(settings) * size
                             if role == "body" else 0.0)
                label_bold = bool(getattr(settings, "bold_label", False)) \
                    and role in LABEL_ROLES
                text = f.text
                if i in head_plans:
                    text = head_plans[i] + clean_heading(text)
                elif i in fig_plans:
                    text = fig_plans[i] + clean_caption(text, aggressive=True)
                elif i in tab_plans:
                    text = tab_plans[i] + clean_caption(text, aggressive=True)

        # ---- 公式：OMML 自绘 / MathType 走 EMF，按文档顺序全部渲染成图 ----
        formula_ims = (_formula_images(para_obj, content_w, size,
                                       latin, bold)
                       if (f.has_math or f.has_ole) else [])
        standalone = bool(formula_ims) and _standalone_formula(
            para_obj, f, len(formula_ims))
        if not formula_ims and f.has_math and side != "before":
            from ..features import _math_text
            text = _math_text(para_obj) or "（公式）"

        pb_before = (pending_break or br_before) and started
        if formula_ims and not standalone:
            # 行内公式：文字和公式图按原文顺序铺在同一段里。
            # 图比行距高时把这一段的固定行距抬到图高，否则公式会压到上一行
            im_h = max(float(im.deviceIndependentSize().height())
                       for im in formula_ims)
            line_pt = max(line_pt, im_h + 2.0)
            _insert_paragraph(cur, "", cjk, latin, size, bold, line_pt, align,
                              indent_pt, color, False, False, left_margin,
                              page_break_before=pb_before,
                              reuse_first=not started)
            block_of[i] = cur.block()
            _insert_inline_formula(cur, para_obj, formula_ims,
                                   cjk, latin, size, bold, color)
        else:
            _insert_paragraph(cur, text, cjk, latin, size, bold, line_pt, align,
                              indent_pt, color, marks_on, label_bold, left_margin,
                              page_break_before=pb_before,
                              reuse_first=not started)
            block_of[i] = cur.block()
        started = True
        pending_break = br_after

        if f.has_image:
            img_align = align if align in ("left", "center", "right") else "center"
            for blob, nat in paragraph_images(doc, para_obj):
                _insert_image(cur, blob, content_w, nat, img_align)
        elif formula_ims and standalone:
            # 显示公式 + 编号同一行（编号右对齐）
            from ..features import equation_label
            _insert_equation(cur, formula_ims,
                             equation_label(para_obj),
                             content_w, latin, size, bold, color)
        elif f.has_ole and not formula_ims and not text.strip():
            _insert_paragraph(cur, "（公式）", FALLBACK_CJK, FALLBACK_LATIN,
                              12.0, False, 15.0, "center", 0.0, "", False)
        else:
            shown += 1

    td.setMetaInformation(QTextDocument.DocumentTitle, "preview")
    return td, block_of


# ------------------------------------------------------------------ #
# 页码
# ------------------------------------------------------------------ #

_ROMAN = {1: "I", 2: "II", 3: "III", 4: "IV", 5: "V", 6: "VI", 7: "VII",
          8: "VIII", 9: "IX", 10: "X", 11: "XI", 12: "XII"}


def page_count(td: QTextDocument) -> int:
    """触发一次整篇布局并返回页数。很贵（实测 60ms/次），所以只调一次。"""
    return max(1, td.pageCount())


def page_labels(td: QTextDocument, block_of: dict[int, QTextBlock],
                body_index: int | None, content_h: float, n: int,
                profile: str = "thesis", show_numbers: bool = True) -> list[str]:
    """每页的页码文字。

    - 关掉页码 → 全空
    - 期刊论文 → 从第 1 页就是阿拉伯数字（没有封面/前置）
    - 学位论文 → 封面空，前置罗马数字，正文阿拉伯从 1 起
    """
    if not show_numbers:
        return [""] * n
    if profile == "journal":
        return [str(p + 1) for p in range(n)]

    # 找不到前置/正文边界（期刊论文常见）时，全部按正文阿拉伯数字，
    # 不要想当然当成前置部分编罗马数字
    front_pages = 0
    if body_index is not None and body_index in block_of and content_h > 1:
        try:
            y = td.documentLayout().blockBoundingRect(block_of[body_index]).y()
            front_pages = max(1, min(n, int(y / content_h)))
        except Exception:
            pass

    out = []
    for p in range(n):
        if p == 0:
            out.append("")                       # 封面：无页码
        elif p < front_pages:
            out.append(_ROMAN.get(p, str(p)))    # 第 1 页前置 = I
        else:
            out.append(str(p - front_pages + 1))
    return out
