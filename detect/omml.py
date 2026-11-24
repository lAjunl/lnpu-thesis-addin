"""把 Word 原生公式（OMML）渲染成图片，给预览用。

背景：Qt 画不出 OMML，MathType 的 EMF 是图片（走 `detect/emf.py`），
而 Word 公式编辑器插入的是 OMML（XML）——以前预览只把 `m:t` 里的文字抽出来
拼成一行，分式、上下标、根号、积分、矩阵的结构全丢（`a/b` 显示成 "ab"）。

这里按 OMML 的结构递归布局，再用 QPainter 画到高分辨率 QImage 上，
按 devicePixelRatio 折算成"磅"插进预览。支持：

    文字 / 分式 / 上标 / 下标 / 上下标 / 根号 / n 元算子(∫∑∏∪∩) /
    定界符(圆括号/方括号/花括号/绝对值…) / 矩阵 / 上下极限 /
    重音符号 / 上下划线 / 函数名 / 公式数组 / 分组字符

不支持的节点会退化成"取里面的文字"，不会崩。
"""
from __future__ import annotations

import math

from docx.oxml.ns import qn
from PySide6.QtCore import QPointF, QSize, Qt
from PySide6.QtGui import QColor, QFont, QFontMetricsF, QImage, QPainter, QPen

INK = QColor("#1d1d1f")
SCRIPT = 0.72          # 上下标字号比例
GAPX = 0.06            # 水平小间距（相对字号）


class Box:
    """一个排版盒：宽、高、基线到顶部的距离。

    结构化节点（分式/根号/矩阵…）会把子盒和尺寸挂在 `Box` 上给 draw 用，
    所以不能加 __slots__。
    """

    def __init__(self, w: float, h: float, base: float):
        self.w, self.h, self.base = w, h, base


def _font(size: float, latin: str, bold: bool = False, italic: bool = False) -> QFont:
    f = QFont()
    f.setFamily(latin)
    f.setPixelSize(max(4.0, float(size)))
    f.setBold(bool(bold))
    f.setItalic(bool(italic))
    return f


def _text_box(text: str, size: float, latin: str,
              bold: bool = False, italic: bool = False) -> Box:
    fm = QFontMetricsF(_font(size, latin, bold, italic))
    asc, desc = fm.ascent(), fm.descent()
    return Box(fm.horizontalAdvance(text), asc + desc, asc)


def _children(el):
    # 属性节点（m:oMathParaPr / m:ctrlPr / m:argPr…）不是排版内容，跳过
    return [c for c in el if isinstance(c.tag, str)
            and c.tag.startswith(qn("m:")) and not c.tag.endswith("Pr")]


def _kid(el, tag: str):
    return el.find(qn("m:" + tag))


def _kids(el, tag: str):
    return el.findall(qn("m:" + tag))


def _run_text(r_el) -> str:
    return "".join(t.text or "" for t in r_el.iter(qn("m:t")))


def _chr_of(el, default: str) -> str:
    """m:dPr/m:chr 或 m:accPr/m:chr 等属性节点里的字符。"""
    for pr_tag in ("dPr", "accPr", "naryPr", "groupChrPr", "limLowPr"):
        pr = _kid(el, pr_tag)
        if pr is None:
            continue
        ch = _kid(pr, "chr")
        if ch is not None:
            v = ch.get(qn("m:val"))
            if v:
                return v
    return default


# ------------------------------------------------------------------ #
# 测量
# ------------------------------------------------------------------ #
# 这些节点本身没有排版结构，只是"一串子节点横排"（分子/分母/上下标/矩阵单元…）
_SEQ_TAGS = {qn(t) for t in (
    "m:oMath", "m:oMathPara", "m:e", "m:num", "m:den", "m:sub", "m:sup",
    "m:lim", "m:deg", "m:mr", "m:fName", "m:funcName")}


def measure(el, size: float, latin: str, bold: bool = False) -> Box:
    if el is None:
        return Box(0.0, size, size * 0.8)
    tag = el.tag
    if tag in (qn("m:r"), qn("m:t")):
        return _text_box(_run_text(el), size, latin, bold)
    if tag in _SEQ_TAGS:
        return _seq_box(el, size, latin, bold)
    if tag == qn("m:f"):
        return _frac_box(el, size, latin, bold)
    if tag == qn("m:sSup"):
        return _ssup_box(el, size, latin, bold, sup=True, sub=False)
    if tag == qn("m:sSub"):
        return _ssup_box(el, size, latin, bold, sup=False, sub=True)
    if tag == qn("m:sSubSup"):
        return _ssup_box(el, size, latin, bold, sup=True, sub=True)
    if tag == qn("m:rad"):
        return _rad_box(el, size, latin, bold)
    if tag == qn("m:nary"):
        return _nary_box(el, size, latin, bold)
    if tag == qn("m:d"):
        return _delim_box(el, size, latin, bold)
    if tag == qn("m:m"):
        return _matrix_box(el, size, latin, bold)
    if tag == qn("m:limLow"):
        return _lim_box(el, size, latin, bold, below=True)
    if tag == qn("m:limUpp"):
        return _lim_box(el, size, latin, bold, below=False)
    if tag == qn("m:acc"):
        return _acc_box(el, size, latin, bold)
    if tag == qn("m:bar"):
        return _bar_box(el, size, latin, bold)
    if tag == qn("m:func"):
        return _func_box(el, size, latin, bold)
    if tag == qn("m:eqArr"):
        return _eqarr_box(el, size, latin, bold)
    if tag == qn("m:groupChr"):
        return _groupchr_box(el, size, latin, bold)
    # 未知节点：把里面的文字取出来当普通文字
    txt = "".join(t.text or "" for t in el.iter(qn("m:t")))
    if txt.strip():
        return _text_box(txt, size, latin, bold)
    return Box(0.0, size, size * 0.8)


def _seq_box(el, size, latin, bold) -> Box:
    x = 0.0
    base = 0.0
    h = 0.0
    first = True
    for ch in _children(el):
        b = measure(ch, size, latin, bold)
        if b.w <= 0 and b.h <= 0:
            continue
        if not first:
            x += size * GAPX
        top = 0.0  # 先累加，后面统一按最大基线对齐
        x += b.w
        base = max(base, b.base)
        h = max(h, b.h)
        first = False
    if first:
        return Box(0.0, size, size * 0.8)
    # 同一行内按基线对齐：最高的那个决定整行高
    h = 0.0
    x = 0.0
    first = True
    for ch in _children(el):
        b = measure(ch, size, latin, bold)
        if b.w <= 0 and b.h <= 0:
            continue
        if not first:
            x += size * GAPX
        x += b.w
        h = max(h, base - b.base + b.h)
        first = False
    return Box(x, h, base)


def _frac_box(el, size, latin, bold) -> Box:
    nb = measure(_kid(el, "num"), size, latin, bold)
    db = measure(_kid(el, "den"), size, latin, bold)
    inner = max(nb.w, db.w)
    pad = size * 0.16
    w = inner + 2 * pad
    rule = max(1.0, size * 0.055)
    gap = size * 0.10
    nb_base = nb.base
    axis = nb_base + gap + rule / 2.0
    den_top = axis + rule / 2.0 + gap
    den_base = den_top + db.base
    h = den_base + (db.h - db.base)
    b = Box(w, h, axis)
    b._frac = (nb, db, inner, pad, rule, nb_base, den_base)   # type: ignore
    return b


def _ssup_box(el, size, latin, bold, sup: bool, sub: bool) -> Box:
    eb = measure(_kid(el, "e"), size, latin, bold)
    gap = size * 0.06
    sb = measure(_kid(el, "sup"), size * SCRIPT, latin, bold) if sup else None
    db = measure(_kid(el, "sub"), size * SCRIPT, latin, bold) if sub else None
    w = eb.w
    if sb is not None:
        w += gap + sb.w
    if db is not None:
        w += gap + db.w
    base = eb.base
    h = eb.h
    if sb is not None:
        sup_base = eb.base - size * 0.34
        h = max(h, sup_base - sb.base + sb.h)
    if db is not None:
        sub_base = eb.base + size * 0.24
        h = max(h, sub_base + (db.h - db.base))
    b = Box(w, h, base)
    b._ss = (eb, sb, db, gap)      # type: ignore
    return b


def _rad_box(el, size, latin, bold) -> Box:
    eb = measure(_kid(el, "e"), size, latin, bold)
    deg = measure(_kid(el, "deg"), size * 0.55, latin, bold)
    dh = eb.h * 1.02
    gf = _font(dh, latin)
    gw = QFontMetricsF(gf).horizontalAdvance("√")
    gap = size * 0.05
    w = gw + gap + eb.w
    over = max(1.0, size * 0.055)
    h = max(eb.h, dh) + over
    b = Box(w, h, eb.base)
    b._rad = (eb, deg, gw, gap, over)   # type: ignore
    return b


def _nary_box(el, size, latin, bold) -> Box:
    ch = _chr_of(el, "∫")
    eb = measure(_kid(el, "e"), size, latin, bold)
    sb = measure(_kid(el, "sub"), size * SCRIPT, latin, bold)
    tb = measure(_kid(el, "sup"), size * SCRIPT, latin, bold)
    gh = size * (1.9 if ch == "∫" else 1.25)
    gf = _font(gh, latin)
    gw = QFontMetricsF(gf).horizontalAdvance(ch)
    g_base = QFontMetricsF(gf).ascent()
    gapx = size * 0.08
    w = gw + gapx + eb.w
    gapy = size * 0.04
    h = sb.h + gh + tb.h + 2 * gapy
    base = sb.h + gapy + g_base
    b = Box(w, h, base)
    b._nary = (ch, eb, sb, tb, gw, gapx, gapy, gh, g_base)   # type: ignore
    return b


def _delim_box(el, size, latin, bold) -> Box:
    eb = measure(_kid(el, "e"), size, latin, bold)
    beg = _chr_of(el, "(")
    # begChr/endChr 存在 m:dPr 里，_chr_of 只看 chr；单独取
    dpr = _kid(el, "dPr")
    if dpr is not None:
        v = _kid(dpr, "begChr")
        if v is not None and v.get(qn("m:val")):
            beg = v.get(qn("m:val"))
        v = _kid(dpr, "endChr")
        end = v.get(qn("m:val")) if v is not None else ")"
    else:
        end = ")"
    dh = eb.h * 1.18
    gap = size * 0.08
    w = size * 0.55 + gap + eb.w + gap + size * 0.55
    h = max(eb.h, dh)
    b = Box(w, h, eb.base)
    b._delim = (eb, beg or "(", end or ")", gap)   # type: ignore
    return b


def _matrix_box(el, size, latin, bold) -> Box:
    rows = _kids(el, "mr")
    cells = []                      # 每格存 (元素, 尺寸)，绘制时要按元素递归
    for r in rows:
        cells.append([(e, measure(e, size, latin, bold))
                      for e in _kids(r, "e")])
    if not cells:
        return Box(0.0, size, size * 0.8)
    ncols = max(len(r) for r in cells)
    colw = [0.0] * ncols
    for r in cells:
        for c, (_, b) in enumerate(r):
            colw[c] = max(colw[c], b.w)
    gapx, gapy = size * 0.34, size * 0.12
    rowb = [max(b.base for _, b in r) for r in cells]
    w = sum(colw) + gapx * (ncols - 1)
    h = sum(max(b.h for _, b in r) for r in cells) + gapy * (len(cells) - 1)
    base = max(b.base for r in cells for _, b in r)
    b = Box(w, h, base)
    b._mat = (cells, colw, gapx, gapy, rowb)   # type: ignore
    return b


def _lim_box(el, size, latin, bold, below: bool) -> Box:
    eb = measure(_kid(el, "e"), size, latin, bold)
    lb = measure(_kid(el, "lim"), size * 0.72, latin, bold)
    gap = size * 0.04
    w = max(eb.w, lb.w)
    if below:
        h = eb.h + gap + lb.h
        base = eb.base
    else:
        h = lb.h + gap + eb.h
        base = lb.h + gap + eb.base
    b = Box(w, h, base)
    b._lim = (eb, lb, gap, below)   # type: ignore
    return b


def _acc_box(el, size, latin, bold) -> Box:
    eb = measure(_kid(el, "e"), size, latin, bold)
    ch = _chr_of(el, "̇")
    ab = _text_box(ch, size * 0.9, latin)
    gap = size * 0.04
    b = Box(max(eb.w, ab.w), eb.h + gap + ab.h * 0.6, eb.base)
    b._acc = (eb, ch, ab, gap)   # type: ignore
    return b


def _bar_box(el, size, latin, bold) -> Box:
    eb = measure(_kid(el, "e"), size, latin, bold)
    t = max(1.0, size * 0.05)
    dpr = _kid(el, "dPr")
    val = "top"
    if dpr is not None:
        v = _kid(dpr, "pos")
        if v is not None and v.get(qn("m:val")):
            val = v.get(qn("m:val"))
    gap = size * 0.05
    h = eb.h + (gap + t if val == "top" else 0) + (gap + t if val == "bot" else 0)
    b = Box(eb.w, h, eb.base + (gap + t if val == "top" else 0))
    b._bar = (eb, val, gap, t)   # type: ignore
    return b


def _func_box(el, size, latin, bold) -> Box:
    eb = measure(_kid(el, "e"), size, latin, bold)
    fn = measure(_kid(el, "funcName"), size, latin, bold)
    gap = size * 0.10
    b = Box(fn.w + gap + eb.w, max(eb.h, fn.h), max(eb.base, fn.base))
    b._func = (eb, fn, gap)   # type: ignore
    return b


def _eqarr_box(el, size, latin, bold) -> Box:
    rows = []
    for mr in _kids(el, "mr"):
        rows.append(_seq_box(mr, size, latin, bold))
    if not rows:
        return Box(0.0, size, size * 0.8)
    gap = size * 0.14
    w = max(r.w for r in rows)
    h = sum(r.h for r in rows) + gap * (len(rows) - 1)
    base = rows[0].base
    b = Box(w, h, base)
    b._eqarr = (rows, gap)   # type: ignore
    return b


def _groupchr_box(el, size, latin, bold) -> Box:
    eb = measure(_kid(el, "e"), size, latin, bold)
    ch = _chr_of(el, "⏟")
    gb = _text_box(ch, size * 1.1, latin)
    gap = size * 0.04
    b = Box(max(eb.w, gb.w), eb.h + gap + gb.h * 0.5, eb.base)
    b._gc = (eb, ch, gb, gap)   # type: ignore
    return b


# ------------------------------------------------------------------ #
# 绘制
# ------------------------------------------------------------------ #
def draw(p: QPainter, el, x: float, base: float, size: float,
         latin: str, bold: bool = False) -> None:
    if el is None:
        return
    tag = el.tag
    if tag in (qn("m:r"), qn("m:t")):
        _draw_text(p, _run_text(el), x, base, size, latin, bold)
        return
    if tag in _SEQ_TAGS:
        _draw_seq(p, el, x, base, size, latin, bold)
        return
    if tag == qn("m:f"):
        _draw_frac(p, el, x, base, size, latin, bold)
        return
    if tag in (qn("m:sSup"), qn("m:sSub"), qn("m:sSubSup")):
        _draw_ss(p, el, x, base, size, latin, bold)
        return
    if tag == qn("m:rad"):
        _draw_rad(p, el, x, base, size, latin, bold)
        return
    if tag == qn("m:nary"):
        _draw_nary(p, el, x, base, size, latin, bold)
        return
    if tag == qn("m:d"):
        _draw_delim(p, el, x, base, size, latin, bold)
        return
    if tag == qn("m:m"):
        _draw_matrix(p, el, x, base, size, latin, bold)
        return
    if tag in (qn("m:limLow"), qn("m:limUpp")):
        _draw_lim(p, el, x, base, size, latin, bold)
        return
    if tag == qn("m:acc"):
        _draw_acc(p, el, x, base, size, latin, bold)
        return
    if tag == qn("m:bar"):
        _draw_bar(p, el, x, base, size, latin, bold)
        return
    if tag == qn("m:func"):
        _draw_func(p, el, x, base, size, latin, bold)
        return
    if tag == qn("m:eqArr"):
        _draw_eqarr(p, el, x, base, size, latin, bold)
        return
    if tag == qn("m:groupChr"):
        _draw_groupchr(p, el, x, base, size, latin, bold)
        return
    txt = "".join(t.text or "" for t in el.iter(qn("m:t")))
    if txt.strip():
        _draw_text(p, txt, x, base, size, latin, bold)


def _draw_text(p, text, x, base, size, latin, bold=False, italic=False):
    if not text:
        return
    p.setFont(_font(size, latin, bold, italic))
    p.setPen(INK)
    p.drawText(QPointF(x, base), text)


def _draw_seq(p, el, x, base, size, latin, bold):
    cx = x
    for ch in _children(el):
        b = measure(ch, size, latin, bold)
        if b.w <= 0 and b.h <= 0:
            continue
        draw(p, ch, cx, base - (b.base - measure(ch, size, latin, bold).base),
             size, latin, bold)
        cx += b.w + size * GAPX


def _draw_frac(p, el, x, base, size, latin, bold):
    b = measure(el, size, latin, bold)
    nb, db, inner, pad, rule, nb_base, den_base = b._frac   # type: ignore
    axis = base
    _draw_text(p, "", x, base, size, latin)  # no-op，保持结构清晰
    draw(p, _kid(el, "num"), x + pad + (inner - nb.w) / 2.0, nb_base, size, latin, bold)
    p.setPen(QPen(INK, rule))
    p.drawLine(QPointF(x + pad, axis), QPointF(x + pad + inner, axis))
    draw(p, _kid(el, "den"), x + pad + (inner - db.w) / 2.0, den_base, size, latin, bold)


def _draw_ss(p, el, x, base, size, latin, bold):
    b = measure(el, size, latin, bold)
    eb, sb, db, gap = b._ss   # type: ignore
    draw(p, _kid(el, "e"), x, base, size, latin, bold)
    cx = x + eb.w + gap
    if sb is not None:
        _draw_text(p, "", x, base, size, latin)
        draw(p, _kid(el, "sup"), cx, base - size * 0.34, size * SCRIPT, latin, bold)
        cx += sb.w + gap
    if db is not None:
        draw(p, _kid(el, "sub"), cx, base + size * 0.24, size * SCRIPT, latin, bold)


def _draw_rad(p, el, x, base, size, latin, bold):
    b = measure(el, size, latin, bold)
    eb, deg, gw, gap, over = b._rad   # type: ignore
    p.setFont(_font(eb.h * 1.02, latin))
    p.setPen(INK)
    p.drawText(QPointF(x, base), "√")
    cx = x + gw + gap
    top = base - eb.base
    p.setPen(QPen(INK, over))
    p.drawLine(QPointF(cx, top), QPointF(cx + eb.w, top))
    draw(p, _kid(el, "e"), cx, base, size, latin, bold)
    if _kid(el, "deg") is not None:
        draw(p, _kid(el, "deg"), x, base - eb.h - size * 0.02, size * 0.55, latin, bold)


def _draw_nary(p, el, x, base, size, latin, bold):
    b = measure(el, size, latin, bold)
    ch, eb, sb, tb, gw, gapx, gapy, gh, g_base = b._nary   # type: ignore
    p.setFont(_font(gh, latin))
    p.setPen(INK)
    p.drawText(QPointF(x, base), ch)
    cx = x + gw + gapx
    draw(p, _kid(el, "e"), cx, base, size, latin, bold)
    if sb is not None:
        draw(p, _kid(el, "sub"), x + (gw - sb.w) / 2.0,
             base + (gh - g_base) + gapy + sb.base, size * SCRIPT, latin, bold)
    if tb is not None:
        draw(p, _kid(el, "sup"), x + (gw - tb.w) / 2.0,
             base - g_base - gapy - (tb.h - tb.base), size * SCRIPT, latin, bold)


def _draw_delim(p, el, x, base, size, latin, bold):
    b = measure(el, size, latin, bold)
    eb, beg, end, gap = b._delim   # type: ignore
    dh = eb.h * 1.18
    p.setFont(_font(dh, latin))
    p.setPen(INK)
    bw = QFontMetricsF(_font(dh, latin)).horizontalAdvance(beg or "(")
    ew = QFontMetricsF(_font(dh, latin)).horizontalAdvance(end or ")")
    p.drawText(QPointF(x, base), beg or "(")
    draw(p, _kid(el, "e"), x + bw + gap, base, size, latin, bold)
    p.setFont(_font(dh, latin))
    p.drawText(QPointF(x + bw + gap + eb.w + gap, base), end or ")")


def _draw_matrix(p, el, x, base, size, latin, bold):
    b = measure(el, size, latin, bold)
    cells, colw, gapx, gapy, rowb = b._mat   # type: ignore
    cy = base - rowb[0]
    for r, row in enumerate(cells):
        if r:
            cy += max(bb.h for _, bb in cells[r - 1]) + gapy
        cx = x
        for c, (cell_el, cell) in enumerate(row):
            draw(p, cell_el, cx + (colw[c] - cell.w) / 2.0, cy + rowb[r],
                 size, latin, bold)
            cx += colw[c] + gapx


def _draw_lim(p, el, x, base, size, latin, bold):
    b = measure(el, size, latin, bold)
    eb, lb, gap, below = b._lim   # type: ignore
    draw(p, _kid(el, "e"), x + (max(eb.w, lb.w) - eb.w) / 2.0, base, size, latin, bold)
    ly = base + (eb.h - eb.base) + gap + lb.base if below else \
        base - eb.base - gap - (lb.h - lb.base)
    draw(p, _kid(el, "lim"), x + (max(eb.w, lb.w) - lb.w) / 2.0, ly,
         size * 0.72, latin, bold)


def _draw_acc(p, el, x, base, size, latin, bold):
    b = measure(el, size, latin, bold)
    eb, ch, ab, gap = b._acc   # type: ignore
    draw(p, _kid(el, "e"), x, base, size, latin, bold)
    _draw_text(p, ch, x + (eb.w - ab.w) / 2.0, base - eb.base - gap - ab.base * 0.4,
               size * 0.9, latin, bold)


def _draw_bar(p, el, x, base, size, latin, bold):
    b = measure(el, size, latin, bold)
    eb, val, gap, t = b._bar   # type: ignore
    off = gap + t if val == "top" else 0.0
    draw(p, _kid(el, "e"), x, base, size, latin, bold)
    y = base - eb.base - gap - t if val == "top" else base + (eb.h - eb.base) + gap
    p.setPen(QPen(INK, t))
    p.drawLine(QPointF(x, y), QPointF(x + eb.w, y))


def _draw_func(p, el, x, base, size, latin, bold):
    b = measure(el, size, latin, bold)
    eb, fn, gap = b._func   # type: ignore
    draw(p, _kid(el, "funcName"), x, base, size, latin, bold)
    draw(p, _kid(el, "e"), x + fn.w + gap, base, size, latin, bold)


def _draw_eqarr(p, el, x, base, size, latin, bold):
    b = measure(el, size, latin, bold)
    rows, gap = b._eqarr   # type: ignore
    top = base - rows[0].base
    for r, rb in enumerate(rows):
        cy = top + sum(rr.h for rr in rows[:r]) + gap * r
        draw(p, _kids(el, "mr")[r], x, cy + rb.base, size, latin, bold)


def _draw_groupchr(p, el, x, base, size, latin, bold):
    b = measure(el, size, latin, bold)
    eb, ch, gb, gap = b._gc   # type: ignore
    draw(p, _kid(el, "e"), x, base, size, latin, bold)
    _draw_text(p, ch, x + (eb.w - gb.w) / 2.0, base + (eb.h - eb.base) + gap,
               size * 1.1, latin, bold)


# ------------------------------------------------------------------ #
# 入口
# ------------------------------------------------------------------ #
def math_roots(p) -> list:
    """段落里所有顶层公式节点（m:oMath / m:oMathPara），按文档顺序。

    m:oMathPara 是容器，它里面的 m:oMath 不重复返回。
    """
    roots = []
    for el in p._p.iter():
        if el.tag == qn("m:oMathPara"):
            roots.append(el)
        elif el.tag == qn("m:oMath"):
            par = el.getparent()
            if par is not None and par.tag == qn("m:oMathPara"):
                continue
            roots.append(el)
    return roots


def render_element(el, size: float, latin: str = "Times New Roman",
                   bold: bool = False, scale: float = 4.0):
    """把一个公式节点（m:oMath / m:oMathPara）画成 QImage。"""
    box = measure(el, size, latin, bold)
    if box.w <= 0 or box.h <= 0:
        return None
    w = max(1, int(math.ceil(box.w * scale)))
    h = max(1, int(math.ceil(box.h * scale)))
    im = QImage(w, h, QImage.Format.Format_ARGB32)
    im.fill(0)                      # 透明底
    pt = QPainter(im)
    pt.setRenderHint(QPainter.RenderHint.Antialiasing)
    pt.scale(scale, scale)
    draw(pt, el, 0.0, box.base, size, latin, bold)
    pt.end()
    im.setDevicePixelRatio(scale)
    return im


def render_omml(p, size: float, latin: str = "Times New Roman",
                cjk: str = "宋体", bold: bool = False,
                scale: float = 4.0):
    """把段落里的 OMML 公式画成 QImage（高分辨率，devicePixelRatio=scale）。

    返回 None 表示这段里没有可渲染的公式。
    """
    roots = math_roots(p)
    if not roots:
        return None
    return render_element(roots[0], size, latin, bold, scale)
