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

from PySide6.QtCore import QSize, QSizeF, Qt
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
MAX_TABLES = 30
MAX_TABLE_ROWS = 10
MAX_TABLE_CELL = 60

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


def paragraph_image_blob(doc, paragraph) -> bytes | None:
    for blip in paragraph._p.findall(".//" + qn("a:blip")):
        rid = blip.get(qn("r:embed"))
        if not rid:
            continue
        try:
            part = doc.part.related_parts.get(rid)
        except Exception:
            part = None
        if part is not None:
            try:
                return part.blob
            except Exception:
                return None
    return None


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


def _insert_paragraph(cur: QTextCursor, text: str, cjk: str, latin: str,
                      size: float, bold: bool, line_pt: float,
                      align: str | None, indent_pt: float = 0.0,
                      color: str = "", marks: bool = False,
                      label_bold: bool = False) -> None:
    bf = QTextBlockFormat()
    bf.setAlignment(_align_flag(align))
    # 直接给磅值：FixedHeight 的单位就是布局单位（1 单位 = 1 磅）
    bf.setLineHeight(max(1.0, line_pt), _FIXED)
    bf.setLeftMargin(0)
    bf.setRightMargin(0)
    if indent_pt:
        bf.setTextIndent(indent_pt)
    cur.insertBlock(bf)
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
                  natural: tuple[float, float] | None = None) -> None:
    """插入图片（居中，单独一行）。"""
    try:
        im = _load_image(blob, content_w, natural)
        if im is None:
            return
        bf = QTextBlockFormat()
        bf.setAlignment(Qt.AlignHCenter)
        cur.insertBlock(bf)
        cur.insertImage(im)
    except Exception:
        pass


def _insert_equation(cur: QTextCursor, im: QImage, number: str,
                     content_w: float, latin: str, size: float, bold: bool,
                     color: str) -> None:
    """公式 + 编号：编号右对齐，和公式同一行（Word 的排法）。

    用 1 行 2 列的无边框表格：左列居中放公式，右列右对齐放编号。
    没有编号就只画公式。
    """
    if not number:
        bf = QTextBlockFormat()
        bf.setAlignment(Qt.AlignHCenter)
        cur.insertBlock(bf)
        cur.insertImage(im)
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
        cur0.insertImage(im)

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
        cur.insertImage(im)


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


def _rule_block(cur: QTextCursor, content_w: float, width_pt: float,
                w: float | None = None, off: float = 0.0) -> None:
    """画一条横线：一个细长的纯色图片（Qt 的块/单元格没有可用的分边边框）。"""
    w = float(w) if w else content_w
    h = max(1, int(round(width_pt)))
    im = QImage(QSize(max(2, int(round(w))), h), QImage.Format.Format_ARGB32)
    im.fill(INK_COLOR)
    bf = QTextBlockFormat()
    bf.setAlignment(Qt.AlignLeft)
    bf.setLeftMargin(off)
    bf.setRightMargin(max(0.0, content_w - off - w))
    bf.setLineHeight(float(h), _FIXED)
    # 表格/frame 后面通常已经留了一个空块，直接复用——否则线会和
    # 表格之间空出一行（实测约 16 磅）
    blk = cur.block()
    if blk.isValid() and not blk.text().strip():
        cur.setBlockFormat(bf)
        cur.insertImage(im)
    else:
        cur.insertBlock(bf)
        cur.insertImage(im)


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
                 color: str, align: str, w: float, off: float) -> None:
    """把一个表格包进无边框 frame 里，用 frame 的左右边距控制表格位置。"""
    tf = QTextTableFormat()
    tf.setBorder(0)
    tf.setCellPadding(3)
    tf.setCellSpacing(0)
    tf.setWidth(QTextLength(QTextLength.FixedLength, w))
    tf.setColumnWidthConstraints(
        [QTextLength(QTextLength.FixedLength, w / max(1, cols))] * cols
    )
    ff = QTextFrameFormat()
    ff.setBorder(0)
    ff.setPadding(0)
    ff.setLeftMargin(off)
    ff.setRightMargin(max(0.0, content_w - off - w))
    fr = cur.insertFrame(ff)
    cur.setPosition(fr.firstPosition())
    t = cur.insertTable(len(rows), cols, tf)
    for r, texts in enumerate(rows):
        for c in range(cols):
            cell = t.cellAt(r, c)
            cfmt = QTextTableCellFormat()
            cfmt.setFont(_font(cjk, size, bold))
            cell.setFormat(cfmt)
            cc = cell.firstCursorPosition()
            bf = QTextBlockFormat()
            bf.setAlignment(_align_flag(align))
            cc.setBlockFormat(bf)
            _insert_cell_text(cc, texts[c] if c < len(texts) else "",
                              cjk, latin, size, bold, color)
    _leave_frame(cur, fr)


def _one_row_table(cur: QTextCursor, texts: list, content_w: float,
                   line_pt: float, cjk: str = FALLBACK_CJK,
                   latin: str = FALLBACK_LATIN, size: float = 10.5,
                   bold: bool = False, color: str = "",
                   align: str = "left", w: float | None = None,
                   off: float = 0.0) -> None:
    w = float(w) if w else content_w
    _frame_table(cur, [list(texts)], len(texts), content_w, line_pt,
                 cjk, latin, size, bold, color, align, w, off)


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
                  line_pt: float = 1.0, cjk: str = FALLBACK_CJK,
                  latin: str = FALLBACK_LATIN, size: float = 10.5,
                  bold: bool = False, color: str = "", align: str = "left",
                  pos: str = "center", width_pt: float | None = None) -> None:
    """自绘三线表。

    为什么不用 QTextTable 的边框：实测 PySide6 的 QTextTableCellFormat 上
    TableCell*Border 属性根本不会被画出来；而单个表格内部又没法插块画线。
    所以拆成「表头一个表格 + 数据一个表格」，中间和两端用极矮块的下/上边框
    画三条线，列宽用 setColumnWidthConstraints 固定住，保证上下对齐。

    表格位置（pos）靠把表格包进 QTextFrame、用 frame 的左右边距实现
    （QTextTable 本身没有对齐属性）。
    """
    try:
        table = doc.tables[t_i]
    except IndexError:
        return
    rows = table.rows[:MAX_TABLE_ROWS]
    if not rows:
        return
    cols = max(1, len(rows[0].cells))
    three = line_pt > 0
    w = max(20.0, min(float(width_pt or content_w), content_w))
    off = {"left": 0.0,
           "center": (content_w - w) / 2.0,
           "right": content_w - w}.get(pos, 0.0)

    def row_text(r, c):
        try:
            return rows[r].cells[c].text.strip().replace("\n", " ")[:MAX_TABLE_CELL]
        except Exception:
            return ""

    if three:
        _rule_block(cur, content_w, line_pt, w, off)      # 顶线
    _one_row_table(cur, [row_text(0, c) for c in range(cols)], content_w,
                   line_pt, cjk, latin, size, bold, color, align, w, off)
    if three:
        _rule_block(cur, content_w, line_pt, w, off)      # 表头线
    body = [[row_text(r, c) for c in range(cols)]
            for r in range(1, len(rows))]
    if body:
        _frame_table(cur, body, cols, content_w, line_pt, cjk, latin, size,
                     bold, color, align, w, off)
    if three:
        _rule_block(cur, content_w, line_pt, w, off)      # 底线
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
        head_plans = {p.index: p.text for p in plan_heading_numbers(
            analysis.detections, body_index, settings)}
        fig_plans = {p.index: p.text for p in plan_caption_numbers(
            analysis.detections, "figure", settings, feats)}
        tab_plans = {p.index: p.text for p in plan_caption_numbers(
            analysis.detections, "table", settings, feats)}
    marks_on = bool(getattr(settings, "enable_marks", True))
    table_line = float(getattr(settings, "table_line", 1.0) or 0) \
        if getattr(settings, "table_style", "") == "three_line" else 0.0
    tbl_cjk = getattr(settings, "table_font", "") or FALLBACK_CJK
    tbl_latin = getattr(settings, "table_ascii", "") or FALLBACK_LATIN
    tbl_size = float(getattr(settings, "table_font_size", 10.5) or 10.5)
    tbl_bold = bool(getattr(settings, "table_bold", False))
    tbl_color = getattr(settings, "table_color", "") or ""
    tbl_align = getattr(settings, "table_align", "left") or "left"
    tbl_pos = getattr(settings, "table_pos", "center") or "center"

    td = QTextDocument()
    td.setPageSize(QSizeF(content_w, content_h))
    td.setDocumentMargin(0)
    td.setDefaultFont(_font(FALLBACK_LATIN, 12, False))

    cur = QTextCursor(td)
    block_of: dict[int, QTextBlock] = {}
    shown = 0
    shown_tbl = 0

    for kind, i in blocks_src:
        if kind == "tbl":
            if shown_tbl < MAX_TABLES:
                _insert_table(doc, cur, i, content_w, table_line,
                              tbl_cjk, tbl_latin, tbl_size, tbl_bold, tbl_color,
                              tbl_align, tbl_pos,
                              table_width_pt(doc.tables[i], content_w))
                shown_tbl += 1
            continue

        f = feats[i]
        if (not f.text.strip() and not f.has_image
                and not f.has_math and not f.has_ole):
            continue
        if not f.has_image and shown >= MAX_PARAS:
            continue

        if side == "before":
            cjk = f.east_asia or FALLBACK_CJK
            latin = f.ascii_font or FALLBACK_LATIN
            size = f.size_pt or 12.0
            bold = bool(f.bold)
            natural = _natural_pt(latin, size)
            line_pt = _line_pt_before(f.line_rule, f.line_value, natural)
            align = f.align
            indent_pt = float(f.first_indent_pt or 0.0)
            color = ""
            text = f.text
            label_bold = False

        else:
            d = analysis.detections[i]
            role = effective_role(d)
            p = dict(settings.roles.get(role) or settings.roles.get("body") or {})
            ov = (overrides or {}).get(i)
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
            indent_pt = (
                float(settings.first_line_indent_chars) * size
                if role == "body" and settings.apply_indent and not settings.apply_not_indent
                else 0.0
            )
            label_bold = bool(getattr(settings, "bold_label", False)) \
                and role in LABEL_ROLES
            text = f.text
            if f.has_math:
                from ..features import _math_text
                text = _math_text(doc.paragraphs[i]) or "（公式）"
            if i in head_plans:
                text = head_plans[i] + clean_heading(text)
            elif i in fig_plans:
                text = fig_plans[i] + clean_caption(text, aggressive=True)
            elif i in tab_plans:
                text = tab_plans[i] + clean_caption(text, aggressive=True)

        _insert_paragraph(cur, text, cjk, latin, size, bold, line_pt, align, indent_pt,
                          color, marks_on, label_bold)
        block_of[i] = cur.block()

        if f.has_image:
            from ..features import image_size
            blob = paragraph_image_blob(doc, doc.paragraphs[i])
            if blob:
                _insert_image(cur, blob, content_w,
                              image_size(doc.paragraphs[i]))
        elif f.has_ole:
            from ..features import ole_image_blob, ole_size, equation_label
            from ..emf import emf_to_png
            blob = ole_image_blob(doc.paragraphs[i])
            osize = ole_size(doc.paragraphs[i])
            nw = osize[0] if osize else None
            png = emf_to_png(blob, int(nw * 6) if nw else int(content_w * 6)) \
                if blob else None
            if png is not None:
                im = _load_image(Path(png).read_bytes(), content_w, osize)
                if im is not None:
                    # 公式 + 编号同一行（编号右对齐）
                    _insert_equation(cur, im,
                                     equation_label(doc.paragraphs[i]),
                                     content_w, latin, size, bold, color)
                else:
                    _insert_paragraph(cur, "（公式）", FALLBACK_CJK,
                                      FALLBACK_LATIN, 12.0, False, 15.0,
                                      "center", 0.0, "", False)
            else:
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
