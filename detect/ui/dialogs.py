"""功能设置卡片（点 Ribbon 按钮弹出）。"""
from __future__ import annotations

from PySide6.QtCore import QSizeF, Qt

from .demos import DemoWidget
from .widgets import MarkCheck
from .settings import COLOR_CHOICES, LINE_RULES
from PySide6.QtGui import QFont, QTextBlockFormat, QTextCharFormat, QTextCursor
from PySide6.QtWidgets import (
    QAbstractItemView,

    QComboBox,
    QDialog,
    QDoubleSpinBox,
    QSpinBox,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from ..classify import effective_role
from .. import __repo_url__, __version__


FONTS = ["宋体", "黑体", "楷体", "仿宋", "Times New Roman"]
LATIN_FONTS = ["Times New Roman", "Arial", "Calibri", "Courier New",
               "Verdana", "Helvetica"]
ALIGNS = [("left", "左对齐"), ("center", "居中"), ("right", "右对齐"), ("both", "两端对齐")]


class _BaseDialog(QDialog):
    def __init__(self, title: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(500, 0)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 20, 20, 20)
        lay.setSpacing(14)
        self.form = QFormLayout()
        self.form.setSpacing(10)
        lay.addLayout(self.form)
        btns = QHBoxLayout()
        btns.addStretch()
        cancel = QPushButton("取消")
        cancel.clicked.connect(self.reject)
        ok = QPushButton("确定")
        ok.setObjectName("primary")
        ok.clicked.connect(self.accept)
        btns.addWidget(cancel)
        btns.addWidget(ok)
        lay.addLayout(btns)

    def row(self, label: str, widget: QWidget) -> None:
        self.form.addRow(label, widget)


class RoleDialog(_BaseDialog):
    """角色格式设置：字号/中文字体/西文字体/加粗/行距/对齐。"""

    def __init__(self, title: str, value: dict, parent=None):
        super().__init__(title, parent)

        self.size = QDoubleSpinBox()
        self.size.setRange(4, 72)
        self.size.setValue(float(value.get("size") or 12))
        self.row("字号（磅）", self.size)

        self.font = QComboBox()
        self.font.addItems(FONTS)
        i = self.font.findText(value.get("font", "宋体"))
        self.font.setCurrentIndex(i if i >= 0 else 0)
        self.row("中文字体", self.font)

        self.latin = QComboBox()
        self.latin.addItems(LATIN_FONTS)
        i = self.latin.findText(value.get("ascii") or "Times New Roman")
        self.latin.setCurrentIndex(i if i >= 0 else 0)
        self.row("西文字体（英文/数字）", self.latin)

        self.bold = MarkCheck("加粗")
        self.bold.setChecked(bool(value.get("bold")))
        self.row("", self.bold)

        self.line_rule = QComboBox()
        for key, text in LINE_RULES:
            self.line_rule.addItem(text, key)
        rule0 = value.get("line_rule") or "exact"
        i = self.line_rule.findData(rule0)
        self.line_rule.setCurrentIndex(i if i >= 0 else 0)
        self.row("行距方式", self.line_rule)

        self.line = QDoubleSpinBox()
        self.line.setRange(1, 100)
        lv0 = float(value.get("line", 22) or 22)
        if rule0 in ("exact", "minimum") and lv0 < 6.0:
            lv0 = max(18.0, round(float(value.get("size") or 12) * 1.5, 1))
        self.line.setValue(lv0)
        self.line.setEnabled(rule0 in ("exact", "minimum"))
        self.row("行距（磅）", self.line)

        self.color = QComboBox()
        for key, text in COLOR_CHOICES:
            self.color.addItem(text, key)
        i = self.color.findData(value.get("color") or "")
        self.color.setCurrentIndex(i if i >= 0 else 0)
        self.row("字色", self.color)

        self.align = QComboBox()
        for key, text in ALIGNS:
            self.align.addItem(text, key)
        i = self.align.findData(value.get("align", "left"))
        self.align.setCurrentIndex(i if i >= 0 else 0)
        self.row("对齐", self.align)

        self.preview = QTextEdit()
        self.preview.setReadOnly(True)
        self.preview.setFixedHeight(112)
        self.row("效果预览", self.preview)

        for w in (self.size, self.line):
            w.valueChanged.connect(lambda *_: self._refresh_preview())
        for w in (self.font, self.latin, self.align, self.color):
            w.currentIndexChanged.connect(lambda *_: self._refresh_preview())
        self.line_rule.currentIndexChanged.connect(self._line_rule_changed)
        self.bold.toggled.connect(lambda *_: self._refresh_preview())
        self._refresh_preview()

    def _line_rule_changed(self) -> None:
        rule = self.line_rule.currentData() or "exact"
        need_pt = rule in ("exact", "minimum")
        self.line.setEnabled(need_pt)
        if need_pt and self.line.value() < 6.0:
            self.line.blockSignals(True)
            self.line.setValue(max(18.0, round(self.size.value() * 1.5, 1)))
            self.line.blockSignals(False)
        self._refresh_preview()

    def _refresh_preview(self) -> None:
        """在卡片里画出当前设置的样子（与主预览同一套渲染规则）。"""
        from .render import _FIXED, _char_format, _line_pt, _natural_pt, _segments

        size = self.size.value()
        bold = self.bold.isChecked()
        cjk = self.font.currentText()
        latin = self.latin.currentText()
        color = self.color.currentData() or ""

        align = {"left": Qt.AlignLeft, "center": Qt.AlignHCenter,
                 "right": Qt.AlignRight, "both": Qt.AlignJustify}.get(
            self.align.currentData() or "left", Qt.AlignLeft)

        natural = _natural_pt(latin, size)
        line_pt = _line_pt(self.line_rule.currentData() or "exact",
                           self.line.value(), natural)

        self.preview.setMinimumWidth(300)
        doc = self.preview.document()
        doc.clear()
        doc.setPageSize(QSizeF(max(200.0, self.preview.width()), 400))
        doc.setDocumentMargin(0)

        bf = QTextBlockFormat()
        bf.setAlignment(align)
        bf.setLineHeight(max(1.0, line_pt), _FIXED)

        cur = QTextCursor(doc)
        # 通用示例句：不出现任何专业词汇，避免用户以为工具只认某一学科
        for line in ("示例文字：中文字体效果 ABC 123",
                     "第二行：字号与行距效果 XYZ 456"):
            cur.insertBlock(bf)
            for is_cjk, seg in _segments(line):
                cur.insertText(seg, _char_format(cjk if is_cjk else latin, size, bold, color))

    def value(self) -> dict:
        return {
            "size": self.size.value(),
            "font": self.font.currentText(),
            "ascii": self.latin.currentText(),
            "bold": self.bold.isChecked(),
            "line": self.line.value(),
            "line_rule": self.line_rule.currentData(),
            "color": self.color.currentData() or "",
            "align": self.align.currentData(),
        }


class PageDialog(_BaseDialog):
    """页面设置：页边距 / 页眉页脚距离。"""

    def __init__(self, value: dict, parent=None):
        super().__init__("页面设置", parent)
        self.fields: dict[str, QDoubleSpinBox] = {}
        for key, label, vmax in (
            ("top", "上边距 (mm)", 80), ("bottom", "下边距 (mm)", 80),
            ("left", "左边距 (mm)", 80), ("right", "右边距 (mm)", 80),
            ("header", "页眉距离 (mm)", 60), ("footer", "页脚距离 (mm)", 60),
        ):
            sb = QDoubleSpinBox()
            sb.setRange(0, vmax)
            sb.setValue(float(value.get(key, 0)))
            self.fields[key] = sb
            self.row(label, sb)

        self.demo = DemoWidget("page", self)
        self.row("示意", self.demo)
        for sb in self.fields.values():
            sb.valueChanged.connect(lambda *_: self._refresh())
        self._refresh()

    def _refresh(self) -> None:
        self.demo.data = {k: sb.value() for k, sb in self.fields.items()}
        self.demo.update()

    def value(self) -> dict:
        return {k: sb.value() for k, sb in self.fields.items()}


class NumberingDialog(_BaseDialog):
    """单个编号种类的设置卡片（标题 / 图 / 表 各一张，不混在一起）。"""

    TITLES = {
        "heading": "标题编号",
        "figure": "图号",
        "table": "表号",
    }
    CHECKS = {
        "heading": ("标题编号（1 / 1.1 / 1.1.1）", "enable_number_heading"),
        "figure": ("图号（图1.1，按章编排）", "enable_number_figure"),
        "table": ("表号（表1.1，按章编排）", "enable_number_table"),
    }
    NOTES = {
        "heading": "层级由识别决定；推断有误请先在「识别明细」里改正。"
                   "默认保留原稿已有的编号，只给缺编号的补号。",
        "figure": "章号取一级标题。默认保留原稿已有的图号，只补缺的。",
        "table": "章号取一级标题。默认保留原稿已有的表号，只补缺的。",
    }
    DEMOS = {
        "heading": [("1  一级标题", 0), ("1.1  二级标题", 1), ("1.1.1  三级标题", 2)],
        "figure": [("图1.1  图题（在图下方）", 0)],
        "table": [("表1.1  表题（在表上方）", 0)],
    }

    def __init__(self, settings, parent=None, kind: str = "heading"):
        self.kind = kind
        super().__init__(self.TITLES.get(kind, "编号设置"), parent)
        label, attr = self.CHECKS[kind]
        self.chk = MarkCheck(label)
        self.chk.setChecked(getattr(settings, attr))
        self.row("", self.chk)

        self.keep = MarkCheck("保留原稿已有编号（只给缺编号的补号）")
        self.keep.setChecked(bool(getattr(settings, "number_keep_existing", True)))
        self.row("", self.keep)

        note = QLabel(self.NOTES[kind])
        note.setObjectName("card-note")
        note.setWordWrap(True)
        self.row("", note)

        if kind == "heading":
            self.sep = QComboBox()
            for key, text in ((".", "点（1.1）"), ("-", "短横（1-1）"),
                              ("", "无分隔（11）")):
                self.sep.addItem(text, key)
            i = self.sep.findData(getattr(settings, "heading_sep", "."))
            self.sep.setCurrentIndex(i if i >= 0 else 0)
            self.row("章号与节号分隔", self.sep)

        self.demo = DemoWidget("number", self)
        self.row("示意", self.demo)
        self.chk.toggled.connect(lambda *_: self._refresh())
        if kind == "heading":
            self.sep.currentIndexChanged.connect(lambda *_: self._refresh())
        self._refresh()

    def _refresh(self) -> None:
        on = self.chk.isChecked()
        self.keep.setEnabled(on)
        if self.kind == "heading":
            self.sep.setEnabled(on)
        if not on:
            items = [("（未启用）", 0)]
        elif self.kind == "heading":
            sep = self.sep.currentData() or ""
            items = [(f"1{sep}  一级标题", 0),
                     (f"1{sep}1{sep}  二级标题", 1),
                     (f"1{sep}1{sep}1{sep}  三级标题", 2)]
        else:
            items = self.DEMOS[self.kind]
        self.demo.data = {"items": items}
        self.demo.update()

    def apply_to(self, settings) -> None:
        setattr(settings, self.CHECKS[self.kind][1], self.chk.isChecked())
        settings.number_keep_existing = self.keep.isChecked()
        if self.kind == "heading":
            settings.heading_sep = self.sep.currentData() or ""


class LayoutDialog(_BaseDialog):
    """单个页面设置卡片（分节 / 页码 / 页眉 各一张，不混在一起）。"""

    TITLES = {
        "section": "分节",
        "pagenum": "页码",
        "header": "页眉",
    }

    def __init__(self, settings, parent=None, kind: str = "section"):
        self.kind = kind
        super().__init__(self.TITLES.get(kind, "页面与页码"), parent)
        if kind == "section":
            self.chk = MarkCheck("分节（封面 / 前置 / 正文）")
            self.chk.setChecked(settings.enable_sections)
            self.row("", self.chk)
            note = QLabel("分节之后才能分别设置页眉、页码和页码格式。")
            note.setObjectName("card-note")
            note.setWordWrap(True)
            self.row("", note)
        elif kind == "pagenum":
            self.chk = MarkCheck("加页码")
            self.chk.setChecked(settings.enable_page_numbers)
            self.row("", self.chk)
            note = QLabel("前置部分用罗马数字 Ⅰ Ⅱ Ⅲ，正文用阿拉伯数字从 1 起。")
            note.setObjectName("card-note")
            note.setWordWrap(True)
            self.row("", note)
        else:
            self.chk = MarkCheck("加页眉")
            self.chk.setChecked(getattr(settings, "enable_header", True))
            self.row("", self.chk)
            self.header = QLineEdit(settings.header_text)
            self.header.setEnabled(self.chk.isChecked())
            self.row("页眉文字", self.header)
            self.hline = QDoubleSpinBox()
            self.hline.setRange(0, 6)
            self.hline.setSingleStep(0.25)
            self.hline.setValue(float(getattr(settings, "header_line", 0) or 0))
            self.row("横线（磅，0=无线）", self.hline)
            self.hcolor = QComboBox()
            for key, text in COLOR_CHOICES:
                self.hcolor.addItem(text, key)
            i = self.hcolor.findData(getattr(settings, "header_line_color", "000000"))
            self.hcolor.setCurrentIndex(i if i >= 0 else 1)
            self.row("横线颜色", self.hcolor)
            self.hstyle = QComboBox()
            for _k, _t in (("single", "单实线"), ("double", "双线"),
                           ("dashed", "虚线"), ("dotted", "点线"),
                           ("dashDot", "点划线")):
                self.hstyle.addItem(_t, _k)
            _i = self.hstyle.findData(getattr(settings, "header_line_style", "single"))
            self.hstyle.setCurrentIndex(_i if _i >= 0 else 0)
            self.row("横线线型", self.hstyle)

            self.gap = QDoubleSpinBox()
            self.gap.setRange(0, 20)
            self.gap.setSingleStep(0.5)
            self.gap.setValue(float(getattr(settings, "header_gap", 4.0) or 0))
            self.row("文字与横线间距（磅）", self.gap)

            self.hfont = QComboBox()
            self.hfont.addItems(FONTS)
            _i = self.hfont.findText(getattr(settings, "header_font", "宋体"))
            self.hfont.setCurrentIndex(_i if _i >= 0 else 0)
            self.row("页眉中文字体", self.hfont)

            self.hlatin = QComboBox()
            self.hlatin.addItems(LATIN_FONTS)
            _i = self.hlatin.findText(getattr(settings, "header_ascii", "Times New Roman"))
            self.hlatin.setCurrentIndex(_i if _i >= 0 else 0)
            self.row("页眉西文字体", self.hlatin)

            self.hsize = QDoubleSpinBox()
            self.hsize.setRange(6, 32)
            self.hsize.setSingleStep(0.5)
            self.hsize.setValue(float(getattr(settings, "header_size", 10.5) or 10.5))
            self.row("页眉字号（磅）", self.hsize)

            self.hbold = MarkCheck("页眉加粗")
            self.hbold.setChecked(bool(getattr(settings, "header_bold", False)))
            self.row("", self.hbold)

            self.hfont_color = QComboBox()
            for key, text in COLOR_CHOICES:
                self.hfont_color.addItem(text, key)
            _i = self.hfont_color.findData(getattr(settings, "header_color", ""))
            self.hfont_color.setCurrentIndex(_i if _i >= 0 else 0)
            self.row("页眉字色", self.hfont_color)

            note = QLabel("封面不加页眉；其余各部分沿用同一页眉。")
            note.setObjectName("card-note")
            note.setWordWrap(True)
            self.row("", note)
            self.header.textChanged.connect(lambda *_: self._refresh())
            self.hline.valueChanged.connect(lambda *_: self._refresh())
            self.hcolor.currentIndexChanged.connect(lambda *_: self._refresh())
            self.gap.valueChanged.connect(lambda *_: self._refresh())
            self.hsize.valueChanged.connect(lambda *_: self._refresh())
            self.hfont.currentIndexChanged.connect(lambda *_: self._refresh())
            self.hlatin.currentIndexChanged.connect(lambda *_: self._refresh())
            self.hfont_color.currentIndexChanged.connect(lambda *_: self._refresh())
            self.hbold.toggled.connect(lambda *_: self._refresh())
            self.chk.toggled.connect(lambda *_: self._set_header_enabled(self.chk.isChecked()))
            self._set_header_enabled(self.chk.isChecked())

        self.demo = DemoWidget({"section": "section", "pagenum": "pagenum",
                                "header": "header"}[kind], self)
        self.row("示意", self.demo)
        self.chk.toggled.connect(lambda *_: self._refresh())
        self._refresh()

    def _set_header_enabled(self, on: bool) -> None:
        for w in (self.header, self.hline, self.hcolor, self.hstyle, self.gap,
                  self.hfont, self.hlatin, self.hsize, self.hbold,
                  self.hfont_color):
            w.setEnabled(on)

    def _refresh(self) -> None:
        on = self.chk.isChecked()
        if self.kind == "section":
            if on:
                self.demo.data = {
                    "labels": [("封面", "无页眉 · 无页码"),
                               ("前置", "页眉 · Ⅰ Ⅱ Ⅲ"),
                               ("正文", "页眉 · 1 2 3")],
                }
            else:
                self.demo.data = {"labels": [("全文", "一节 · 页码连续")]}
        elif self.kind == "pagenum":
            self.demo.data = {"roman": on}
        else:
            self.demo.data = {
                "header_text": self.header.text() or "（无）",
                "line": 0.0 if not on else self.hline.value(),
                "show_line": on,
                "line_style": self.hstyle.currentData() or "single",
                "gap": self.gap.value(),
                "size": self.hsize.value(),
                "bold": self.hbold.isChecked(),
            }
        self.demo.update()

    def apply_to(self, settings) -> None:
        if self.kind == "section":
            settings.enable_sections = self.chk.isChecked()
        elif self.kind == "pagenum":
            settings.enable_page_numbers = self.chk.isChecked()
        else:
            settings.enable_header = self.chk.isChecked()
            settings.header_text = self.header.text()
            settings.header_line = self.hline.value()
            settings.header_line_color = self.hcolor.currentData() or "000000"
            settings.header_line_style = self.hstyle.currentData() or "single"
            settings.header_gap = self.gap.value()
            settings.header_font = self.hfont.currentText()
            settings.header_ascii = self.hlatin.currentText()
            settings.header_size = self.hsize.value()
            settings.header_bold = self.hbold.isChecked()
            settings.header_color = self.hfont_color.currentData() or ""


class IndentDialog(_BaseDialog):
    """正文缩进 / 段落顶格。"""

    def __init__(self, settings, parent=None, focus_not_indent: bool = False):
        super().__init__("正文缩进 / 段落顶格", parent)
        self.chk = MarkCheck("正文首行缩进")
        self.chk.setChecked(settings.apply_indent)
        self.row("", self.chk)
        self.chars = QDoubleSpinBox()
        self.chars.setRange(0, 6)
        self.chars.setValue(settings.first_line_indent_chars)
        self.row("缩进字符数", self.chars)
        self.not_indent = MarkCheck("段落顶格（取消所有首行缩进）")
        self.not_indent.setChecked(settings.apply_not_indent or focus_not_indent)
        self.row("", self.not_indent)

        self.demo = DemoWidget("indent", self)
        self.row("示意", self.demo)
        self.chars.valueChanged.connect(lambda *_: self._refresh())
        self.chk.toggled.connect(lambda *_: self._refresh())
        self.not_indent.toggled.connect(lambda *_: self._refresh())
        self._refresh()

    def _refresh(self) -> None:
        self.demo.kind = "topalign" if self.not_indent.isChecked() else "indent"
        self.demo.data = {"chars": self.chars.value()}
        self.demo.update()

    def apply_to(self, settings) -> None:
        settings.apply_indent = self.chk.isChecked()
        settings.first_line_indent_chars = self.chars.value()
        settings.apply_not_indent = self.not_indent.isChecked()


HELP_LINES = [
    "怎么用",
    "────────────",
    "1. 打开文档：点左上角「打开文档」或直接把 .docx / .doc 拖进窗口，",
    "   双栏预览立刻生成（左=原稿、右=排版后；.doc 会自动转 .docx）。",
    "2. 选「格式方案」（右下角）：内置「期刊论文默认」和「学位论文默认」，",
    "   学位论文=三节+罗马/阿拉伯页码，期刊论文=单节+阿拉伯页码从 1 起。",
    "3. 按角色设格式：顶部 Ribbon 每个按钮一张卡片，卡片里的",
    "   「示意」会跟着设置实时变；三线表支持顶/底线（1.5 磅）与",
    "   栏目线（0.75 磅）两粗一细独立调节。",
    "4. 保存格式方案：调好参数后点右下角「格式方案 → 保存当前设置为方案…」，",
    "   无需经过格式识别即可存为可复用方案，支持选「期刊论文」或「学位论文」。",
    "5. 从模板学格式：点顶栏「从原稿识别 ▾」，可从当前文档或模板文件读取",
    "   各角色格式，核对后直接保存为方案。",
    "6. 单独/批量改段落：右栏单击某段弹浮动格式条（支持字体/字号/加粗/",
    "   行距方式/行距/对齐/字色，或「设为」某个角色）；Ctrl+点击可多选批量改。",
    "7. 一键排版：只生成结果并刷新预览，不落盘。",
    "8. 保存：顶栏「保存 ▾」自己选覆盖原文件（自动备份 .bak.docx）",
    "   还是另存为。",
    "",
    "几个常用开关",
    "────────────",
    "· 同步：两栏一起滑动（左栏标题行右侧）。",
    "· −/＋/适应窗口/适应整页 / Ctrl+滚轮：每栏单独缩放。",
    "· 识别明细：逐段查看自动识别结果，可手工改正角色/层级。",
    "",
    "自动上下标",
    "────────────",
    "CO2、H2S、CH4 的数字转下标，m2、cm3 转上标。",
    "Figure 2 / Table 3 / A4 / COVID-19 不会被动。",
    "",
    "已知限制",
    "────────────",
    "· 预览的分页位置按 Qt 的字体度量算，和 Word 不会完全一致。",
    "· 目录页码要在 Word 里按 F9 刷新（域缓存改不了）。",
    "· Word 公式（OMML）与 MathType 公式在预览里以位图绘制，排版结果保留原公式可继续编辑。",
]


class HelpDialog(QDialog):
    """单独的帮助文档窗口（从「关于」里点进来）。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("帮助文档")
        self.resize(560, 620)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 20, 20, 20)
        txt = QTextEdit()
        txt.setReadOnly(True)
        txt.setPlainText("\n".join(HELP_LINES))
        lay.addWidget(txt, 1)
        b = QPushButton("关闭")
        b.clicked.connect(self.accept)
        lay.addWidget(b)


class AboutDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("关于")
        self.resize(460, 0)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 24, 24, 24)
        lay.setSpacing(10)
        lay.addWidget(QLabel("论文格式助手"))
        lay.addWidget(QLabel("面向中文学位论文 / 期刊论文（科技论文体系）"))
        ver = QLabel(f"版本 v{__version__}")
        ver.setObjectName("about-version")
        lay.addWidget(ver)
        lay.addWidget(QLabel("本地运行，文档不上传。"))
        repo = QLabel(
            f'GitHub 技术文档与更新：<br>'
            f'<a href="{__repo_url__}" style="color:#0071e3;">{__repo_url__}</a>'
        )
        repo.setOpenExternalLinks(True)
        repo.setTextInteractionFlags(Qt.TextInteractionFlag.TextBrowserInteraction)
        repo.setWordWrap(True)
        self.repo_label = repo
        lay.addWidget(repo)
        link = QPushButton("帮助文档")
        link.setFlat(True)
        link.setCursor(Qt.CursorShape.PointingHandCursor)
        link.setObjectName("help-link")
        link.clicked.connect(self._open_help)
        lay.addWidget(link)
        b = QPushButton("关闭")
        b.clicked.connect(self.accept)
        lay.addWidget(b)

    def _open_help(self, _url: str = "") -> None:
        HelpDialog(self).exec()


# 角色选项：value = 写回 Detection.role 的值（heading1/2/3 已含层级）
ROLE_CHOICES = [
    ("title", "论文题目"),
    ("heading1", "一级标题"),
    ("heading2", "二级标题"),
    ("heading3", "三级标题"),
    ("body", "正文"),
    ("figure_caption", "图注"),
    ("table_caption", "表注"),
    ("figure", "图"),
    ("table", "表格"),
    ("equation", "公式"),
    ("reference", "参考文献"),
    ("abstract", "摘要"),
    ("keywords", "关键词"),
    ("toc", "目录标题"),
    ("toc_item", "目录条目"),
    ("meta", "封面信息"),
    ("other", "其他"),
]
ROLE_TEXT = dict(ROLE_CHOICES)


class TableDialog(_BaseDialog):
    """表格样式：三线表 / 网格 / 无线，线宽，表文字字号。"""

    def __init__(self, settings, parent=None):
        super().__init__("表格样式", parent)
        self.style = QComboBox()
        for key, text in (("three_line", "三线表（顶线+表头线+底线）"),
                          ("grid", "网格（全部框线）"),
                          ("none", "无线")):
            self.style.addItem(text, key)
        i = self.style.findData(getattr(settings, "table_style", "three_line"))
        self.style.setCurrentIndex(i if i >= 0 else 0)
        self.row("样式", self.style)

        self.line = QDoubleSpinBox()
        self.line.setRange(0, 6)
        self.line.setSingleStep(0.25)
        self.line.setDecimals(2)
        self.line.setValue(float(getattr(settings, "table_line", 1.5) or 0))
        self.row("顶线/底线（外框）粗细（磅）", self.line)

        self.line_inner = QDoubleSpinBox()
        self.line_inner.setRange(0, 6)
        self.line_inner.setSingleStep(0.25)
        self.line_inner.setDecimals(2)
        self.line_inner.setValue(float(getattr(settings, "table_line_inner", 0.75) or 0))
        self.row("栏目线（内线）粗细（磅）", self.line_inner)

        self.font_size = QDoubleSpinBox()
        self.font_size.setRange(6, 20)
        self.font_size.setSingleStep(0.5)
        self.font_size.setValue(float(getattr(settings, "table_font_size", 10.5)))
        self.row("表文字字号（磅）", self.font_size)

        self.font = QComboBox()
        self.font.addItems(FONTS)
        i = self.font.findText(getattr(settings, "table_font", "宋体"))
        self.font.setCurrentIndex(i if i >= 0 else 0)
        self.row("表内中文字体", self.font)

        self.latin = QComboBox()
        self.latin.addItems(LATIN_FONTS)
        i = self.latin.findText(getattr(settings, "table_ascii", "Times New Roman"))
        self.latin.setCurrentIndex(i if i >= 0 else 0)
        self.row("表内西文字体", self.latin)

        bold_wrap = QWidget(self)
        bold_lay = QHBoxLayout(bold_wrap)
        bold_lay.setContentsMargins(0, 0, 0, 0)
        bold_lay.setSpacing(16)
        self.header_bold = MarkCheck("表头加粗")
        self.header_bold.setChecked(
            bool(getattr(settings, "table_header_bold",
                         getattr(settings, "table_bold", False))))
        self.bold = MarkCheck("表体加粗")
        self.bold.setChecked(bool(getattr(settings, "table_bold", False)))
        bold_lay.addWidget(self.header_bold)
        bold_lay.addWidget(self.bold)
        bold_lay.addStretch(1)
        self.row("加粗", bold_wrap)

        self.color = QComboBox()
        for key, text in COLOR_CHOICES:
            self.color.addItem(text, key)
        i = self.color.findData(getattr(settings, "table_color", ""))
        self.color.setCurrentIndex(i if i >= 0 else 0)
        self.row("表内字色", self.color)

        self.align = QComboBox()
        for key, text in ALIGNS:
            self.align.addItem(text, key)
        i = self.align.findData(getattr(settings, "table_align", "left"))
        self.align.setCurrentIndex(i if i >= 0 else 0)
        self.row("表内文字对齐", self.align)

        self.pos = QComboBox()
        for key, text in (("left", "靠左"), ("center", "居中"),
                          ("right", "靠右")):
            self.pos.addItem(text, key)
        i = self.pos.findData(getattr(settings, "table_pos", "center"))
        self.pos.setCurrentIndex(i if i >= 0 else 0)
        self.row("表格位置", self.pos)

        self.clear_shading = MarkCheck("清除原表底纹（去掉表头/行背景填充色）")
        self.clear_shading.setChecked(
            bool(getattr(settings, "table_clear_shading", True)))
        self.row("底纹", self.clear_shading)

        note = QLabel("三线表常用「两粗一细」：顶线/底线 1.5 磅，表头下方的栏目线 0.75 磅。")
        note.setObjectName("card-note")
        note.setWordWrap(True)
        self.row("", note)

        self.demo = DemoWidget("table", self)
        self.row("示意", self.demo)
        for w in (self.style, self.font, self.latin, self.color):
            w.currentIndexChanged.connect(lambda *_: self._refresh())
        for w in (self.line, self.line_inner, self.font_size):
            w.valueChanged.connect(lambda *_: self._refresh())
        self.header_bold.toggled.connect(lambda *_: self._refresh())
        self.bold.toggled.connect(lambda *_: self._refresh())
        self.align.currentIndexChanged.connect(lambda *_: self._refresh())
        self.pos.currentIndexChanged.connect(lambda *_: self._refresh())
        self._refresh()

    def _refresh(self) -> None:
        st = self.style.currentData() or "three_line"
        self.line.setEnabled(st != "none")
        self.line_inner.setEnabled(st != "none")
        self.demo.data = {
            "style": st,
            "outer": self.line.value(),
            "inner": self.line_inner.value(),
            "size": self.font_size.value(),
            "header_bold": self.header_bold.isChecked(),
            "bold": self.bold.isChecked(),
        }
        self.demo.update()

    def apply_to(self, settings) -> None:
        settings.table_style = self.style.currentData()
        settings.table_line = self.line.value()
        settings.table_line_inner = self.line_inner.value()
        settings.table_font_size = self.font_size.value()
        settings.table_font = self.font.currentText()
        settings.table_ascii = self.latin.currentText()
        settings.table_header_bold = self.header_bold.isChecked()
        settings.table_bold = self.bold.isChecked()
        settings.table_color = self.color.currentData() or ""
        settings.table_align = self.align.currentData() or "left"
        settings.table_pos = self.pos.currentData() or "center"
        settings.table_clear_shading = self.clear_shading.isChecked()


class CaptionNumberDialog(_BaseDialog):
    """图号 / 表号的编号格式自定义。"""

    def __init__(self, settings, parent=None, kind: str = "figure"):
        self.kind = kind
        super().__init__("图号格式" if kind == "figure" else "表号格式", parent)
        self._en_attr = "enable_number_figure" if kind == "figure" else "enable_number_table"
        self.chk = MarkCheck("启用图号编排（图1.1）" if kind == "figure" else "启用表号编排（表1.1）")
        self.chk.setChecked(bool(getattr(settings, self._en_attr, False)))
        self.row("", self.chk)

        pfx = "caption_fig_prefix" if kind == "figure" else "caption_tab_prefix"
        self.prefix = QLineEdit(getattr(settings, pfx, "图" if kind == "figure" else "表"))
        self.row("前缀文字", self.prefix)

        self.auto_pfx = MarkCheck("自动沿用原稿前缀（Fig. / Table）")
        self.auto_pfx.setChecked(
            bool(getattr(settings, "caption_prefix_auto", True)))
        self.row("", self.auto_pfx)

        self.sep = QComboBox()
        for key, text in ((".", "点（图1.1）"), ("-", "短横（图1-1）"),
                          ("", "无分隔（图11）")):
            self.sep.addItem(text, key)
        i = self.sep.findData(getattr(settings, "caption_sep", "."))
        self.sep.setCurrentIndex(i if i >= 0 else 0)
        self.row("章号与序号分隔", self.sep)

        self.by_chapter = MarkCheck("按章编号（第 N 章的第 M 图）")
        self.by_chapter.setChecked(getattr(settings, "caption_by_chapter", True))
        self.row("", self.by_chapter)

        self.start = QSpinBox()
        self.start.setRange(1, 99)
        self.start.setValue(int(getattr(settings, "caption_start", 1) or 1))
        self.row("起始序号", self.start)

        self.keep = MarkCheck("保留原稿已有编号（只给缺编号的补号）")
        self.keep.setChecked(bool(getattr(settings, "number_keep_existing", True)))
        self.row("", self.keep)

        self.demo = DemoWidget("number", self)
        self.row("示意", self.demo)
        self.chk.toggled.connect(lambda *_: self._refresh())
        self.prefix.textChanged.connect(lambda *_: self._refresh())
        self.auto_pfx.toggled.connect(lambda *_: self._refresh())
        self.sep.currentIndexChanged.connect(lambda *_: self._refresh())
        self.by_chapter.toggled.connect(lambda *_: self._refresh())
        self.start.valueChanged.connect(lambda *_: self._refresh())
        self._refresh()

    def _refresh(self) -> None:
        on = self.chk.isChecked()
        for w in (self.auto_pfx, self.sep, self.by_chapter, self.start, self.keep):
            w.setEnabled(on)
        self.prefix.setEnabled(on and not self.auto_pfx.isChecked())
        if not on:
            items = [("（未启用）", 0)]
        else:
            pfx = self.prefix.text() or ("图" if self.kind == "figure" else "表")
            sep = self.sep.currentData() or ""
            if self.by_chapter.isChecked():
                items = [(f"{pfx}1{sep}1  第 1 章第 1 项", 0),
                         (f"{pfx}1{sep}2  第 1 章第 2 项", 0)]
            else:
                items = [(f"{pfx}1  全文第 1 项", 0), (f"{pfx}2  全文第 2 项", 0)]
        self.demo.data = {"items": items}
        self.demo.update()

    def apply_to(self, settings) -> None:
        setattr(settings, self._en_attr, self.chk.isChecked())
        pfx = "caption_fig_prefix" if self.kind == "figure" else "caption_tab_prefix"
        setattr(settings, pfx, self.prefix.text())
        settings.caption_prefix_auto = self.auto_pfx.isChecked()
        settings.caption_sep = self.sep.currentData() or ""
        settings.caption_by_chapter = self.by_chapter.isChecked()
        settings.caption_start = self.start.value()
        settings.number_keep_existing = self.keep.isChecked()


class ParaFormatDialog(QDialog):
    """点预览里的某一段弹出的浮动格式条（段落级手动格式）。"""

    def __init__(self, para_index: int, value: dict, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.Popup | Qt.FramelessWindowHint)
        self.para_index = para_index
        lay = QVBoxLayout(self)
        lay.setContentsMargins(12, 10, 12, 10)
        lay.setSpacing(6)
        self.setObjectName("para-pop")
        self.title = QLabel("")
        self.title.setObjectName("para-pop-title")
        self.title.setWordWrap(True)
        self.title.setFixedWidth(300)
        lay.addWidget(self.title)

        # 「设为」：把这一段改成某个角色（一级标题 / 图注 / 表注 …）
        r0 = QHBoxLayout()
        r0.addWidget(QLabel("设为"))
        self.role = QComboBox()
        for val, text in ROLE_CHOICES:
            self.role.addItem(text, val)
        self.role.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        r0.addWidget(self.role, 1)
        lay.addLayout(r0)

        r1 = QHBoxLayout()
        self.font = QComboBox(); self.font.addItems(FONTS)
        self.latin = QComboBox(); self.latin.addItems(LATIN_FONTS)
        for w, label in ((self.font, "中文"), (self.latin, "西文")):
            lab = QLabel(label)
            r1.addWidget(lab)
            r1.addWidget(w, 1)
        lay.addLayout(r1)

        r2 = QHBoxLayout()
        r2.addWidget(QLabel("字号"))
        self.size = QDoubleSpinBox()
        self.size.setRange(4, 72)
        self.size.setDecimals(1)
        self.size.setFixedWidth(84)
        self.bold = MarkCheck("加粗")
        r2.addWidget(self.size)
        r2.addSpacing(8)
        r2.addWidget(self.bold)
        r2.addStretch()
        lay.addLayout(r2)

        r2b = QHBoxLayout()
        r2b.addWidget(QLabel("行距"))
        self.line_rule = QComboBox()
        for key, text in LINE_RULES:
            self.line_rule.addItem(text, key)
        self.line_rule.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.line = QDoubleSpinBox()
        self.line.setRange(1, 100)
        self.line.setDecimals(1)
        self.line.setFixedWidth(84)
        r2b.addWidget(self.line_rule, 1)
        r2b.addWidget(self.line)
        lay.addLayout(r2b)

        self.align = QComboBox()
        for key, text in ALIGNS:
            self.align.addItem(text, key)
        lay.addWidget(self.align)

        r4 = QHBoxLayout()
        r4.addWidget(QLabel("字色"))
        self.color = QComboBox()
        for key, text in COLOR_CHOICES:
            self.color.addItem(text, key)
        self.color.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        r4.addWidget(self.color, 1)
        lay.addLayout(r4)

        r3 = QHBoxLayout()
        clear = QPushButton("恢复为角色默认")
        clear.clicked.connect(self._clear)
        allc = QPushButton("清除全部手动")
        allc.clicked.connect(self._clear_all)
        r3.addWidget(clear)
        r3.addWidget(allc)
        lay.addLayout(r3)

        for w in (self.font, self.latin, self.align, self.color):
            w.currentIndexChanged.connect(lambda *_: self._changed())
        self.line_rule.currentIndexChanged.connect(self._line_rule_changed)
        for w in (self.size, self.line):
            w.valueChanged.connect(lambda *_: self._changed())
        self.bold.toggled.connect(lambda *_: self._changed())
        self.role.currentIndexChanged.connect(self._role_changed)
        self.value_before = dict(value)
        self.changed = False
        self.on_change = None
        self.on_clear = None
        self.on_clear_all = None
        self.on_role = None
        # 批量改格式时，一次改多段（Ctrl 多选）
        self.para_indices: list = [para_index]

    # ---------------------------------------------------------------- #
    def set_para_indices(self, indices) -> None:
        self.para_indices = list(indices) if isinstance(indices, (list, tuple, set)) \
            else [indices]

    def primary_index(self) -> int:
        return self.para_indices[0] if self.para_indices else self.para_index

    def set_title(self, text: str) -> None:
        self.title.setText(text)

    def set_role(self, role: str) -> None:
        i = self.role.findData(role)
        self.role.setCurrentIndex(i if i >= 0 else 0)

    def _role_changed(self) -> None:
        if self.on_role is not None:
            self.on_role(self.para_indices, self.role.currentData())

    def _line_rule_changed(self) -> None:
        rule = self.line_rule.currentData() or "exact"
        need_pt = rule in ("exact", "minimum")
        self.line.setEnabled(need_pt)
        if need_pt and self.line.value() < 6.0:
            # 从 1.5 倍切到固定值时，避免把 1.5 当成 1.5 磅导致文字重叠
            self.line.blockSignals(True)
            self.line.setValue(max(18.0, round(self.size.value() * 1.5, 1)))
            self.line.blockSignals(False)
        self._changed()

    def set_value(self, v: dict) -> None:
        widgets = (self.font, self.latin, self.size, self.line_rule,
                   self.line, self.bold, self.align, self.color)
        for w in widgets:
            w.blockSignals(True)
        try:
            idx_f = self.font.findText(v.get("font", "宋体"))
            if idx_f >= 0:
                self.font.setCurrentIndex(idx_f)
            else:
                self.font.addItem(v.get("font", "宋体"))
                self.font.setCurrentIndex(self.font.count() - 1)
            idx_l = self.latin.findText(v.get("ascii") or "Times New Roman")
            if idx_l >= 0:
                self.latin.setCurrentIndex(idx_l)
            sz = float(v.get("size", 12))
            self.size.setValue(sz)
            rule = v.get("line_rule") or "exact"
            i_r = self.line_rule.findData(rule)
            self.line_rule.setCurrentIndex(i_r if i_r >= 0 else 0)
            need_pt = rule in ("exact", "minimum")
            self.line.setEnabled(need_pt)
            lv = float(v.get("line", 22) or 22)
            if need_pt and lv < 6.0:
                lv = max(18.0, round(sz * 1.5, 1))
            self.line.setValue(lv)
            self.bold.setChecked(bool(v.get("bold")))
            i = self.align.findData(v.get("align", "left"))
            self.align.setCurrentIndex(i if i >= 0 else 0)
            i = self.color.findData(v.get("color") or "")
            self.color.setCurrentIndex(i if i >= 0 else 0)
        finally:
            for w in widgets:
                w.blockSignals(False)

    def value(self) -> dict:
        rule = self.line_rule.currentData() or "exact"
        return {
            "font": self.font.currentText(),
            "ascii": self.latin.currentText(),
            "size": self.size.value(),
            "bold": self.bold.isChecked(),
            "line": self.line.value(),
            "line_rule": rule,
            "align": self.align.currentData(),
            "color": self.color.currentData() or "",
        }

    def _changed(self) -> None:
        self.changed = True
        if self.on_change is not None:
            self.on_change(self.para_indices, self.value())

    def _clear(self) -> None:
        if self.on_clear is not None:
            self.on_clear(self.para_indices)
        self.hide()

    def _clear_all(self) -> None:
        if self.on_clear_all is not None:
            self.on_clear_all()
        self.hide()


class RolesDialog(QDialog):
    """识别明细：逐段查看自动识别结果，并可手工改正角色/层级。

    自动识别对"无样式、无编号"的稿子只能靠字数/词数/间距猜层级，必然有错。
    没有这个入口，猜错的一级标题会被永久套用错误的格式。
    """

    def __init__(self, analysis, doc=None, parent=None):
        super().__init__(parent)
        self.analysis = analysis
        self.doc = doc
        self.changed = 0
        self.setWindowTitle("识别明细（可手工改正）")
        self.resize(760, 620)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 16, 16, 16)
        lay.setSpacing(10)

        tip = QLabel(
            f"共 {analysis.paragraph_count} 段。下面列出结构段（标题/题目/图表注等），"
            "改「角色」列后点确定，预览与排版会立刻按新角色重算。"
        )
        tip.setWordWrap(True)
        tip.setObjectName("tip")
        lay.addWidget(tip)

        bar = QHBoxLayout()
        self.search = QLineEdit()
        self.search.setPlaceholderText("搜索文字…")
        self.search.textChanged.connect(self._filter)
        bar.addWidget(self.search, 1)
        self.only_err = MarkCheck("只显示层级把握低（<0.6）的标题")
        self.only_err.toggled.connect(self._filter)
        bar.addWidget(self.only_err)
        lay.addLayout(bar)

        self.table = QTableWidget(0, 4, self)
        self.table.setHorizontalHeaderLabels(["段落", "内容", "角色", "依据"])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        # 行高必须显式给：单元格里没有 item 时默认行高只有 12px，
        # 会把下拉框压扁、文字被裁掉（实测 combo 只剩 12px 高）。
        self.table.verticalHeader().setDefaultSectionSize(30)
        self.table.verticalHeader().setMinimumSectionSize(30)
        # 列宽必须写死：角色列是单元格控件（不是 item），
        # 用 ResizeToContents 只会按表头"角色"两个字算，把下拉框挤到 78px 看不见字。
        self.table.setColumnWidth(0, 56)
        self.table.setColumnWidth(1, 210)
        self.table.setColumnWidth(2, 150)
        lay.addWidget(self.table, 1)

        # 底部：选中行 → 改成某个角色
        edit = QHBoxLayout()
        edit.addWidget(QLabel("把选中行改为："))
        self.role_box = QComboBox()
        for val, text in ROLE_CHOICES:
            self.role_box.addItem(text, val)
        edit.addWidget(self.role_box, 1)
        apply_role = QPushButton("应用")
        apply_role.clicked.connect(self._apply_role)
        edit.addWidget(apply_role)
        self.status = QLabel("")
        self.status.setObjectName("card-note")
        edit.addWidget(self.status)
        lay.addLayout(edit)
        self.table.itemSelectionChanged.connect(self._on_select)

        btns = QHBoxLayout()
        btns.addStretch()
        reset = QPushButton("重置为自动识别结果")
        reset.clicked.connect(self._reset)
        btns.addWidget(reset)
        cancel = QPushButton("取消")
        cancel.clicked.connect(self.reject)
        ok = QPushButton("确定")
        ok.setObjectName("primary")
        ok.clicked.connect(self.accept)
        btns.addWidget(cancel)
        btns.addWidget(ok)
        lay.addLayout(btns)

        self._rows: list[int] = []          # 表格行 → 段落 index
        self._edits: dict[int, str] = {}    # 段落 index → 新角色
        self._fill()

    # ---------------------------------------------------------------- #
    def _candidates(self) -> list[int]:
        out = []
        for d in self.analysis.detections:
            if d.role in ("body", "table", "figure", "other"):
                continue
            out.append(d.index)
        return out

    def _fill(self) -> None:
        self.table.setRowCount(0)
        self._rows = []
        for i in self._candidates():
            d = self.analysis.detections[i]
            f = self.analysis.features[i]
            row = self.table.rowCount()
            self.table.insertRow(row)
            self._rows.append(i)

            self.table.setItem(row, 0, QTableWidgetItem(str(i)))
            self.table.setItem(row, 1, QTableWidgetItem(f.text[:60]))
            why = "；".join(d.reasons) if d.reasons else ""
            if d.level is not None:
                why += f"｜层级 {d.level}（把握 {d.level_confidence:.2f}）"
            self.table.setItem(row, 3, QTableWidgetItem(why))
            # 角色列用纯文本：单元格控件（下拉框）在本项目的样式表下
            # 会被压扁/画不出字（实测行高 12px、文字不可见），改成
            # 只读文本 + 底部统一的角色选择器。
            self.table.setItem(row, 2, QTableWidgetItem(
                ROLE_TEXT.get(effective_role(d), effective_role(d))))
        self._filter()

    def _apply_role(self) -> None:
        """把底部选择器里的角色写到选中的行上。"""
        role = self.role_box.currentData()
        if not role:
            return
        text = ROLE_TEXT.get(role, role)
        n = 0
        for row in self.table.selectionModel().selectedRows():
            r = row.row()
            if 0 <= r < len(self._rows):
                self._edits[self._rows[r]] = role
                self.table.setItem(r, 2, QTableWidgetItem(text))
                n += 1
        if n:
            self.status.setText(f"已改 {n} 行")

    def _filter(self) -> None:
        key = self.search.text().strip().lower()
        only_err = self.only_err.isChecked()
        for row, i in enumerate(self._rows):
            d = self.analysis.detections[i]
            f = self.analysis.features[i]
            show = True
            if key and key not in f.text.lower():
                show = False
            if only_err and not (d.role == "heading"
                                 or (d.level or 0) and d.level_confidence < 0.6):
                show = False
            self.table.setRowHidden(row, not show)

    def _reset(self) -> None:
        """丢弃所有手工改动，重新自动识别。"""
        if self.doc is None:
            QMessageBox.information(self, "提示", "没有原始文档，无法重算。")
            return
        from ..classify import analyze

        fresh = analyze(self.doc, source=self.analysis.source)
        a = self.analysis
        a.scheme, a.scheme_stats = fresh.scheme, fresh.scheme_stats
        a.detections, a.features = fresh.detections, fresh.features
        a.level_note = fresh.level_note
        self._edits.clear()
        self.status.setText("")
        self._fill()

    def _on_select(self) -> None:
        rows = self.table.selectionModel().selectedRows()
        self.status.setText(f"已选 {len(rows)} 行" if rows else "")

    # ---------------------------------------------------------------- #
    def apply_to(self, analysis) -> None:
        """把手工改过的角色写回 analysis.detections。"""
        changed = 0
        for i, role in self._edits.items():
            d = analysis.detections[i]
            if effective_role(d) == role:
                continue
            d.role = role
            d.level = int(role[-1]) if role.startswith("heading") and role != "heading" else None
            d.level_confidence = 1.0
            d.user_confirmed = True
            d.reasons = list(d.reasons) + ["手工改正"]
            changed += 1
        self.changed = changed

