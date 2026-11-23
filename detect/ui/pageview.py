"""逐页预览控件：把 QTextDocument 画成一页一页的 A4 纸。

旧版用 QTextBrowser，是一条无限滚动的文本流——没有纸张、页边距、页眉、
页码、分页，所以看不出版式。这里改成真正的页面视图：
白底纸张 + 页眉 + 页脚页码 + Qt 自己算出来的分页位置。

注意天花板：Qt 用自己的字体度量分页，分页位置和 Word 不会完全一致，
只能当参考。要 100% 真实需要 Word 自己渲染。
"""
from __future__ import annotations

from PySide6.QtCore import QRectF, Qt, QPointF
from PySide6.QtGui import QColor, QFont, QFontMetricsF, QPainter, QPen
from PySide6.QtWidgets import QScrollArea, QWidget


class PageView(QWidget):
    PAD = 12          # 栏内边距（收紧，别让纸页飘着）
    GAP = 16          # 页与页之间的间距

    ZOOM_MIN, ZOOM_MAX = 50, 200

    def __init__(self, parent=None):
        super().__init__(parent)
        self._doc = None
        self._paper = (595.276, 841.890)          # A4，单位 = 点
        self._margins = (70.87, 85.04, 70.87, 85.04)
        self._content_h = 700.0
        self._labels: list[str] = []
        self._header = ""
        self._show_header = True
        self._show_pagenum = True
        self._zoom = 1.0
        self._pages = 1
        self._fit_mode = "window"       # window 适应窗口 / page 适应整页 / pct 固定比例
        self._zoom_pct = 100
        self._sel_block = None
        self._header_line = 0.0
        self._header_line_style = "single"
        self._header_line_color = QColor("#000000")
        self._header_font_cjk = "宋体"
        self._header_font_latin = "Times New Roman"
        self._header_size = 10.5
        self._header_bold = False
        self._header_color = QColor("#555557")
        self._header_gap = 4.0
        self.on_paragraph = None
        self.zoom_changed = None        # 回调：缩放变化时通知面板标题

    # ---------------------------------------------------------------- #
    def set_paper(self, paper, margins) -> None:
        self._paper = (float(paper[0]), float(paper[1]))
        t, r, b, l = margins
        self._margins = (float(t), float(r), float(b), float(l))
        self.refit()

    def set_header_line(self, width_pt: float, color: str,
                        style: str = "single") -> None:
        self._header_line = float(width_pt or 0)
        self._header_line_style = style or "single"
        try:
            self._header_line_color = QColor("#" + (color or "000000").lstrip("#"))
        except Exception:
            self._header_line_color = QColor("#000000")
        self.update()

    def set_header_style(self, cjk: str = "宋体", latin: str = "Times New Roman",
                         size: float = 10.5, bold: bool = False,
                         color: str = "", gap: float = 4.0) -> None:
        """页眉文字的字体/字号/加粗/字色，以及文字与横线的距离（磅）。"""
        self._header_font_cjk = cjk or "宋体"
        self._header_font_latin = latin or "Times New Roman"
        self._header_size = float(size or 10.5)
        self._header_bold = bool(bold)
        self._header_gap = float(gap or 0)
        try:
            self._header_color = QColor("#" + (color or "555557").lstrip("#"))
        except Exception:
            self._header_color = QColor("#555557")
        self.update()

    def set_document(self, doc, labels, header: str = "",
                     show_header: bool = True, show_pagenum: bool = True,
                     page_count: int | None = None,
                     src_of_block: dict | None = None) -> None:
        self._doc = doc
        self._labels = list(labels or [])
        self._header = header
        self._show_header = show_header
        self._show_pagenum = show_pagenum
        self._src_of_block = src_of_block or {}
        self._sel_block = None
        if page_count is None:
            page_count = doc.pageCount() if doc is not None else 1
        self._pages = max(1, page_count)
        self.refit()

    # ---------------------------------------------------------------- #
    def set_fit_mode(self, mode: str, pct: int | None = None) -> None:
        if mode not in ("window", "page", "pct"):
            return
        self._fit_mode = mode
        if pct is not None:
            self._zoom_pct = max(self.ZOOM_MIN, min(self.ZOOM_MAX, int(pct)))
        self.refit()

    def set_selected_block(self, block) -> None:
        self._sel_block = block
        self.update()

    def zoom_percent(self) -> int:
        return int(round(self._zoom * 100))

    def refit(self, avail_width: float | None = None, avail_height: float | None = None) -> None:
        if avail_width is None:
            avail_width = float(self.width())
        pw, ph0 = max(1.0, self._paper[0]), max(1.0, self._paper[1])
        if self._fit_mode == "window":
            z = (avail_width - 2 * self.PAD) / pw
        elif self._fit_mode == "page" and avail_height:
            z = min((avail_width - 2 * self.PAD) / pw,
                    (avail_height - 2 * self.PAD) / ph0)
        else:
            z = self._zoom_pct / 100.0
        self._zoom = max(0.20, min(3.0, z))
        ph = ph0 * self._zoom
        self._content_h = ph0 - self._margins[0] - self._margins[2]
        self._content_w = self._paper[0] - self._margins[3] - self._margins[1]
        h = self.PAD * 2 + self._pages * (ph + self.GAP) - self.GAP
        # 宽度始终等于可用宽度，纸页在中间居中（窗口放大后不再靠左）
        self.setFixedSize(max(80, int(avail_width)), max(80, int(h)))
        self.update()

    def page_x(self) -> float:
        pw = self._paper[0] * self._zoom
        return max(self.PAD, (self.width() - pw) / 2)

    def document(self):
        return self._doc

    def page_count(self) -> int:
        return self._pages

    def y_for_page(self, i: int) -> float:
        """第 i 页纸面在控件里的 y 坐标。"""
        return self.PAD + i * (self._paper[1] * self._zoom + self.GAP)

    def page_for_y(self, y: int) -> int:
        """控件 y → 页号。必须用 floor，不能用 round。

        页步长是「纸高 + 页间距」，round 会让每一页下半部分的点击
        被算到下一页（实测 258 次点击只命中 139 次），
        表现为「点某一段，弹出的却是上面几段」。
        """
        step = self._paper[1] * self._zoom + self.GAP
        return max(0, min(self._pages - 1,
                          int((y - self.PAD) // max(1.0, step))))

    # ---------------------------------------------------------------- #
    # 段落级手动格式：点击段落 → 通知外部
    # ---------------------------------------------------------------- #
    def set_paragraph_handler(self, fn) -> None:
        self.on_paragraph = fn

    def _zoom_geom(self):
        z = self._zoom
        ph = self._paper[1] * z
        l = self._margins[3] * z
        t = self._margins[0] * z
        return z, ph, l, t

    def mousePressEvent(self, event) -> None:
        fn = getattr(self, "on_paragraph", None)
        if fn is None or self._doc is None:
            super().mousePressEvent(event)
            return
        z, ph, l, t = self._zoom_geom()
        y = event.pos().y()
        x = event.pos().x()
        page = self.page_for_y(y)
        # 控件坐标 → 文档坐标（见 paintEvent 的变换，反解）
        qx = (x - self.page_x() - l) / z
        qy = (y - self.PAD - page * (ph + self.GAP) - t) / z + page * self._content_h
        if qx < 0 or qy < 0:
            return
        try:
            pos = self._doc.documentLayout().hitTest(QPointF(qx, qy), Qt.FuzzyHit)
        except Exception:
            return
        if pos < 0:
            return
        block = self._doc.findBlock(pos)
        src = self._src_of_block.get(block.position())
        if src is None:
            return
        fn(src, self.mapToGlobal(event.pos()), block)

    # ---------------------------------------------------------------- #
    def wheelEvent(self, event) -> None:
        """Ctrl + 滚轮 = 缩放（Word 的操作习惯）；没按 Ctrl 就照常滚动。"""
        if event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            dy = event.angleDelta().y()
            if dy:
                base = self._zoom_pct if self._fit_mode == "pct" else self.zoom_percent()
                self.set_fit_mode("pct", base + (10 if dy > 0 else -10))
                if self.zoom_changed is not None:
                    self.zoom_changed()
            event.accept()
            return
        super().wheelEvent(event)

    # ---------------------------------------------------------------- #
    def paintEvent(self, event) -> None:
        p = QPainter(self)
        p.fillRect(event.rect(), QColor("#ececf0"))
        if self._doc is None:
            p.setPen(QColor("#86868b"))
            p.drawText(self.rect(), Qt.AlignCenter, "先打开一个 docx")
            p.end()
            return

        z = self._zoom
        pw, ph = self._paper[0] * z, self._paper[1] * z
        t = self._margins[0] * z
        b = self._margins[2] * z
        l = self._margins[3] * z
        r = self._margins[1] * z
        cw, ch = pw - l - r, ph - t - b
        px = self.page_x()
        exposed = event.rect()

        for i in range(self._pages):
            y = self.PAD + i * (ph + self.GAP)
            page = QRectF(px, y, pw, ph)
            if not exposed.intersects(page.toRect()):
                continue

            p.setPen(QPen(QColor("#dcdce2"), 1))
            p.setBrush(QColor("#ffffff"))
            p.drawRect(page.adjusted(0.5, 0.5, -0.5, -0.5))

            cx, cy = page.x() + l, page.y() + t

            # 选中段落的高亮（画在文字下面）
            if self._sel_block is not None and self._doc is not None:
                try:
                    br = self._doc.documentLayout().blockBoundingRect(self._sel_block)
                except Exception:
                    br = None
                if br is not None and br.height() > 0:
                    top = br.y() - i * self._content_h
                    if top + br.height() > 0 and top < self._content_h:
                        p.save()
                        p.setClipRect(QRectF(cx, cy, cw, ch))
                        p.translate(cx, cy)
                        p.scale(z, z)
                        p.translate(0, -i * self._content_h)
                        p.fillRect(QRectF(br.x(), br.y(), br.width(), br.height()),
                                   QColor(0, 122, 255, 34))
                        p.restore()

            p.save()
            p.setClipRect(QRectF(cx, cy, cw, ch))
            p.translate(cx, cy)
            p.scale(z, z)
            p.translate(0, -i * self._content_h)
            # 必须把裁剪矩形一起给 drawContents：不给的话 Qt 会走整篇文档
            # （实测 47 页时每次绘制 67ms，滚动一卡一卡；给了之后只画这一页）
            self._doc.drawContents(
                p, QRectF(0.0, i * self._content_h,
                          self._content_w, self._content_h))
            p.restore()

            self._draw_header(p, QRectF(page.x() + l, page.y(), cw, t), i, z)
            self._draw_pagenum(p, QRectF(page.x() + l, page.y() + ph - b, cw, b), i, z)

        p.end()

    def _draw_header(self, p: QPainter, rect: QRectF, page_i: int, z: float) -> None:
        if not self._show_header or page_i == 0 or not self._header:
            return
        from .render import _font, _segments

        # 文字底部抬到横线以上 header_gap 磅，否则文字贴着线（旧版写死 1 磅）
        text_rect = QRectF(rect)
        text_rect.setBottom(rect.bottom() - self._header_gap * z)
        if text_rect.height() <= 2:
            text_rect = QRectF(rect)

        segs = []
        for is_cjk, s in _segments(self._header):
            segs.append((s, _font(self._header_font_cjk if is_cjk
                                  else self._header_font_latin,
                                  self._header_size, self._header_bold)))
        total = sum(QFontMetricsF(f).horizontalAdvance(s) for s, f in segs)
        x = rect.center().x() - total / 2.0
        p.setPen(self._header_color)
        for s, f in segs:
            fm = QFontMetricsF(f)
            p.setFont(f)
            p.drawText(QPointF(x, text_rect.bottom() - fm.descent()), s)
            x += fm.horizontalAdvance(s)

        if self._header_line > 0:
            st = self._header_line_style
            pen = QPen(self._header_line_color, max(0.5, self._header_line * z))
            if st == "dashed":
                pen.setStyle(Qt.PenStyle.DashLine)
            elif st == "dotted":
                pen.setStyle(Qt.PenStyle.DotLine)
            elif st == "dashDot":
                pen.setStyle(Qt.PenStyle.DashDotLine)
            p.setPen(pen)
            y = rect.bottom() - 1
            if st == "double":          # 双线：画两条紧挨着的线
                p.drawLine(QPointF(rect.left(), y - 1), QPointF(rect.right(), y - 1))
                p.drawLine(QPointF(rect.left(), y + 1.5), QPointF(rect.right(), y + 1.5))
            else:
                p.drawLine(QPointF(rect.left(), y), QPointF(rect.right(), y))

    def _draw_pagenum(self, p: QPainter, rect: QRectF, page_i: int, z: float) -> None:
        if not self._show_pagenum or page_i >= len(self._labels) or not self._labels[page_i]:
            return
        f = QFont()
        f.setFamily("Times New Roman")
        f.setPixelSize(max(6, int(round(10.5 * z))))
        p.setFont(f)
        p.setPen(QColor("#555557"))
        p.drawText(rect, Qt.AlignHCenter | Qt.AlignTop, self._labels[page_i])


class PreviewScroll(QScrollArea):
    """把 PageView 的宽度变化同步给它，让它按可用宽度缩放/居中。"""

    def __init__(self, widget: PageView, parent=None):
        super().__init__(parent)
        self.setWidget(widget)
        self.setWidgetResizable(False)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        # viewport 默认是白底，内容比视口矮时会露白
        self.viewport().setStyleSheet("background:#ececf0;")
        self.setStyleSheet("QScrollArea{border:none;background:#ececf0;}")

    def refit_child(self) -> None:
        w = self.widget()
        if isinstance(w, PageView):
            w.refit(float(self.viewport().width()), float(self.viewport().height()))

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self.refit_child()
