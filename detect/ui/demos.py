"""设置卡片的"效果示意"：用 QPainter 绘图，不用字符画。"""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QFont, QFontMetrics, QPainter, QPen
from PySide6.QtWidgets import QSizePolicy, QWidget

INK = QColor("#1d1d1f")
SUBTLE = QColor("#86868b")
LINE = QColor("#c7c7cc")
ACCENT = QColor("#007aff")
PAPER = QColor("#ffffff")


class DemoWidget(QWidget):
    """根据 kind 画出对应设置的效果示意。

    kind ∈ page / indent / topalign / number / caption_fig / caption_tab /
            section / pagenum / header
    """

    def __init__(self, kind: str, parent=None):
        super().__init__(parent)
        self.kind = kind
        heights = {
            "page": 190, "indent": 118, "topalign": 108, "number": 150,
            "caption_fig": 110, "caption_tab": 132, "table": 124,
            "section": 132, "pagenum": 96, "header": 96,
        }
        self.setMinimumHeight(heights.get(kind, 120))
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.data: dict = {}

    # ---------------------------------------------------------------- #
    def paintEvent(self, _e) -> None:
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.fillRect(self.rect(), PAPER)
        try:
            {
                "page": self._page,
                "indent": self._indent,
                "topalign": self._topalign,
                "number": self._number,
                "caption_fig": self._caption_fig,
                "caption_tab": self._caption_tab,
                "table": self._table,
                "section": self._section,
                "pagenum": self._pagenum,
                "header": self._header,
            }[self.kind](p)
        except KeyError:
            pass
        p.end()

    # ---------------------------------------------------------------- #
    def _font(self, size: float = 9, bold: bool = False) -> QFont:
        f = QFont("Microsoft YaHei")
        f.setPointSizeF(size)
        if bold:
            f.setWeight(QFont.Weight.Bold)
        return f

    def _label(self, p: QPainter, x, y, text: str, color: QColor = SUBTLE,
               size: float = 8.5, bold: bool = False) -> None:
        p.setPen(color)
        p.setFont(self._font(size, bold))
        p.drawText(int(x), int(y), text)

    def _page_box(self, p: QPainter, x, y, w, h) -> None:
        p.setPen(QPen(LINE, 1))
        p.setBrush(QColor("#fafafc"))
        p.drawRect(int(x), int(y), int(w), int(h))

    # ---------------------------------------------------------------- #
    # 页面设置
    # ---------------------------------------------------------------- #
    def _page(self, p: QPainter) -> None:
        d = self.data
        top = d.get("top", 25); bottom = d.get("bottom", 25)
        left = d.get("left", 30); right = d.get("right", 30)
        header = d.get("header", 15); footer = d.get("footer", 20)

        # 左边画纸，右边一列数值——不把字压到纸上（旧版在小尺寸下全叠在一起）
        pad = 12
        col_x = self.width() - 108
        area_w = max(60, col_x - pad * 2)
        area_h = max(60, self.height() - pad * 2)
        scale = min(area_w / 150.0, area_h / 210.0)
        pw = 150 * scale
        ph = 210 * scale
        x = pad + max(0.0, (area_w - pw) / 2)
        y = pad + max(0.0, (area_h - ph) / 2)
        self._page_box(p, x, y, pw, ph)

        p.setPen(Qt.NoPen)
        p.setBrush(QColor(ACCENT.red(), ACCENT.green(), ACCENT.blue(), 30))
        p.drawRect(int(x), int(y), int(pw), max(1.0, top * scale))
        p.drawRect(int(x), int(y + ph - bottom * scale), int(pw), max(1.0, bottom * scale))
        p.drawRect(int(x), int(y), max(1.0, left * scale), int(ph))
        p.drawRect(int(x + pw - right * scale), int(y), max(1.0, right * scale), int(ph))

        body_x = x + left * scale
        body_y = y + top * scale
        body_w = max(1.0, pw - (left + right) * scale)
        body_h = max(1.0, ph - (top + bottom) * scale)
        p.setPen(QPen(ACCENT, 1, Qt.PenStyle.DashLine))
        p.setBrush(Qt.NoBrush)
        p.drawRect(int(body_x), int(body_y), int(body_w), int(body_h))

        p.setPen(QPen(LINE, 1, Qt.PenStyle.DotLine))
        if header > 0.5:
            p.drawLine(int(body_x), int(y + header * scale),
                       int(body_x + body_w), int(y + header * scale))
        if footer > 0.5:
            p.drawLine(int(body_x), int(y + ph - footer * scale),
                       int(body_x + body_w), int(y + ph - footer * scale))

        rows = [
            ("上", f"{top:g} mm", ACCENT),
            ("下", f"{bottom:g} mm", ACCENT),
            ("左", f"{left:g} mm", ACCENT),
            ("右", f"{right:g} mm", ACCENT),
            ("页眉", f"{header:g} mm", SUBTLE),
            ("页脚", f"{footer:g} mm", SUBTLE),
        ]
        n = len(rows)
        for i, (name, val, col) in enumerate(rows):
            ry = y + 10 + i * (ph - 20) / max(1, n - 1)
            self._label(p, col_x, ry, name, col, 9, True)
            self._label(p, col_x + 38, ry, val, INK, 9)

    # ---------------------------------------------------------------- #
    # 缩进 / 顶格
    # ---------------------------------------------------------------- #
    def _indent(self, p: QPainter) -> None:
        chars = self.data.get("chars", 2)
        y = 30
        line_h = 22
        for i in range(2):
            indent = int(chars * 12) if i == 0 else 0
            p.setPen(QPen(LINE, 1))
            p.drawLine(20 + indent, y + i * line_h, self.width() - 20, y + i * line_h)
        # 缩进标记
        if chars:
            x0 = 20
            x1 = 20 + int(chars * 12)
            p.setPen(QPen(ACCENT, 1.4))
            p.drawLine(x0, y - 6, x1, y - 6)
            p.drawLine(x0, y - 9, x0, y - 3)
            p.drawLine(x1, y - 9, x1, y - 3)
            self._label(p, x0 + 4, y - 12, f"{chars:g} 字", ACCENT, 8.5, True)
        self._label(p, 20, y + 2 * line_h + 8, "首行缩进效果（每段开头）")

    def _topalign(self, p: QPainter) -> None:
        y = 26
        for i in range(3):
            p.setPen(QPen(LINE, 1))
            p.drawLine(20, y + i * 20, self.width() - 20, y + i * 20)
        p.setPen(QPen(ACCENT, 1.2))
        p.drawRect(20, y - 6, 8, 8)
        self._label(p, 34, y, "顶格：所有段落从最左侧开始，不缩进", ACCENT, 9, True)

    # ---------------------------------------------------------------- #
    # 编号
    # ---------------------------------------------------------------- #
    def _number(self, p: QPainter) -> None:
        x = 24
        y = 26
        items = self.data.get("items") or [
            ("1  一级标题", 0), ("1.1  二级标题", 1), ("1.1.1  三级标题", 2),
        ]
        line_h = 24
        for i, item in enumerate(items):
            text, lvl = item if isinstance(item, (tuple, list)) else (item, 0)
            self._label(p, x + lvl * 20, y + i * line_h, text,
                        ACCENT if lvl == 0 else INK, 10, lvl == 0)
            if lvl > 0:
                p.setPen(QPen(LINE, 1, Qt.PenStyle.DotLine))
                p.drawLine(x + lvl * 20 - 8, y + i * line_h - 4,
                           x + lvl * 20 - 8, y + i * line_h + 8)

    # ---------------------------------------------------------------- #
    # 图注 / 表注
    # ---------------------------------------------------------------- #
    def _caption_fig(self, p: QPainter) -> None:
        w = min(self.width() - 40, 260)
        x = (self.width() - w) / 2
        y = 14
        p.setPen(QPen(LINE, 1))
        p.drawRect(int(x), int(y), int(w), int(w * 0.62))
        p.setPen(QPen(SUBTLE, 1))
        p.drawLine(int(x), int(y + w * 0.4), int(x + w), int(y + w * 0.4))
        self._label(p, x + w / 2 - 30, y + w * 0.4 + 20, "图1.1  图题示例",
                    INK, 10, True)
        self._label(p, x + w / 2 - 40, y + w * 0.4 + 34, "图题在图的下方，居中加粗",
                    SUBTLE, 8.5)

    def _caption_tab(self, p: QPainter) -> None:
        d = self.data
        size = float(d.get("size", 10.5) or 10.5)
        bold = bool(d.get("bold", False))
        text = d.get("text") or "表1.1  表题示例"
        w = min(self.width() - 40, 260)
        x = (self.width() - w) / 2
        y = 10
        f = self._font(size, bold)
        fm = QFontMetrics(f)
        p.setFont(f)
        p.setPen(INK)
        p.drawText(int(x + w / 2 - fm.horizontalAdvance(text) / 2),
                   int(y + fm.ascent()), text)
        self._label(p, x + w / 2 - 44, y + size + 20,
                    "表题在表的上方，居中加粗", SUBTLE, 8.5)
        top = y + size + 38
        p.setPen(QPen(INK, 1.2))
        p.drawLine(int(x), int(top), int(x + w), int(top))
        p.setPen(QPen(LINE, 1))
        for i in range(1, 3):
            p.drawLine(int(x), int(top + i * 14), int(x + w), int(top + i * 14))
        p.drawLine(int(x + w / 2), int(top), int(x + w / 2), int(top + 3 * 14))

    def _table(self, p: QPainter) -> None:
        d = self.data
        style = d.get("style") or "three_line"
        outer = max(0.0, float(d.get("outer", 1.5) or 0))
        inner = max(0.0, float(d.get("inner", 0.75) or 0))
        w = min(self.width() - 48, 280)
        x = (self.width() - w) / 2
        y0 = 18
        rh = 22
        rows = 3
        y_end = y0 + rows * rh
        cols = ["项目", "指标 A", "指标 B"]
        vals = [["样本 1", "12.4", "0.85"], ["样本 2", "15.8", "0.92"]]

        # 画表线
        if style == "three_line":
            if outer > 0:
                p.setPen(QPen(INK, outer))
                p.drawLine(int(x), int(y0), int(x + w), int(y0))
                p.drawLine(int(x), int(y_end), int(x + w), int(y_end))
            hw = inner if inner > 0 else outer
            if hw > 0:
                p.setPen(QPen(INK, hw))
                p.drawLine(int(x), int(y0 + rh), int(x + w), int(y0 + rh))
        elif style == "grid":
            iw = inner if inner > 0 else outer
            if iw > 0:
                p.setPen(QPen(INK, iw))
                for r in range(1, rows):
                    p.drawLine(int(x), int(y0 + r * rh), int(x + w), int(y0 + r * rh))
                for c in range(1, 3):
                    cx = x + w * c / 3
                    p.drawLine(int(cx), int(y0), int(cx), int(y_end))
            if outer > 0:
                p.setPen(QPen(INK, outer))
                p.drawRect(int(x), int(y0), int(w), int(rows * rh))
        else:
            p.setPen(QPen(LINE, 1, Qt.PenStyle.DashLine))
            p.drawRect(int(x), int(y0), int(w), int(rows * rh))

        # 单元格示意文字
        all_rows = [cols] + vals
        for r, row_txts in enumerate(all_rows):
            for c, txt in enumerate(row_txts):
                cx = x + w * c / 3 + 10
                cy = y0 + r * rh + 15
                self._label(p, cx, cy, txt, INK, 9.5, r == 0)

        note = {
            "three_line": f"三线表：顶/底线 {outer:g} 磅（粗）· 栏目线 {inner:g} 磅（细）",
            "grid": f"网格表：外框 {outer:g} 磅 · 内线 {inner:g} 磅",
            "none": "无线表：不显示框线",
        }.get(style, "")
        self._label(p, x, y_end + 18, note, SUBTLE, 8.5)

    # ---------------------------------------------------------------- #
    # 分节 / 页码 / 页眉
    # ---------------------------------------------------------------- #
    def _section(self, p: QPainter) -> None:
        labels = self.data.get("labels") or [
            ("封面", "无页眉 · 无页码"),
            ("前置", "页眉 · Ⅰ Ⅱ Ⅲ"),
            ("正文", "页眉 · 1 2 3"),
        ]
        w = min((self.width() - 60) / 3, 110)
        gap = 12
        total = w * 3 + gap * 2
        x0 = (self.width() - total) / 2
        y = 16
        h = self.height() - 54
        for i, (title, sub) in enumerate(labels):
            x = x0 + i * (w + gap)
            self._page_box(p, x, y, w, h)
            if i > 0:
                p.setPen(QPen(ACCENT, 1))
                p.drawLine(int(x + 8), int(y + h * 0.09), int(x + w - 8), int(y + h * 0.09))
                self._label(p, x + w / 2 - 8, y + h * 0.09 + 12, "页眉", SUBTLE, 7.5)
            if i == 0:
                self._label(p, x + w / 2 - 12, y + h - 12, "无页码", SUBTLE, 7.5)
            else:
                roman = "Ⅰ" if i == 1 else "1"
                p.setPen(QPen(ACCENT, 1))
                p.drawEllipse(int(x + w / 2 - 9), int(y + h - 24), 18, 14)
                self._label(p, x + w / 2 - 3, y + h - 20, roman, ACCENT, 8, True)
            self._label(p, x, y + h + 16, title, INK, 9, True)
            self._label(p, x, y + h + 28, sub, SUBTLE, 7.5)
            if i < 2:
                p.setPen(QPen(LINE, 1))
                p.drawLine(int(x + w + 2), int(y + h / 2), int(x + w + gap - 2), int(y + h / 2))

    def _pagenum(self, p: QPainter) -> None:
        roman = self.data.get("roman", True)
        w = min(self.width() - 40, 240)
        x = (self.width() - w) / 2
        y = 20
        h = 54
        self._page_box(p, x, y, w, h)
        p.setPen(QPen(LINE, 1))
        p.drawLine(int(x), int(y + h - 16), int(x + w), int(y + h - 16))
        text = "Ⅰ  Ⅱ  Ⅲ" if roman else "1  2  3"
        p.setPen(ACCENT)
        p.setFont(self._font(11, True))
        p.drawText(int(x + w / 2 - 24), int(y + h - 20), text)
        self._label(p, x, y + h + 14, "前置用罗马数字，正文用阿拉伯数字", SUBTLE, 8)

    def _header(self, p: QPainter) -> None:
        d = self.data
        text = d.get("header_text") or "硕士学位论文"
        line = float(d.get("line", 0.75) or 0)
        gap = float(d.get("gap", 4.0) or 0)
        size = float(d.get("size", 10.5) or 10.5)
        bold = bool(d.get("bold", False))
        w = min(self.width() - 40, 240)
        x = (self.width() - w) / 2
        y = 18
        h = 58
        self._page_box(p, x, y, w, h)
        line_y = y + 32
        if line > 0 and d.get("show_line", True):
            p.setPen(QPen(INK, max(1.0, line)))
            p.drawLine(int(x), int(line_y), int(x + w), int(line_y))
        f = self._font(size, bold)
        p.setFont(f)
        p.setPen(INK)
        fm = QFontMetrics(f)
        p.drawText(int(x + w / 2 - fm.horizontalAdvance(text) / 2),
                   int(line_y - gap - fm.descent()), text)
        self._label(p, x, y + h + 14,
                    f"文字与横线间距 {gap:g} 磅 · 字号 {size:g}", SUBTLE, 8)
