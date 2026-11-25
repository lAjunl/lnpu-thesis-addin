"""主窗口：标题行 + Ribbon 功能区 + 工作区（Word 实时对比）。"""
from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QEvent, Qt, QTimer, QPoint, QPointF, QRectF
from PySide6.QtGui import (QActionGroup, QBrush, QColor, QKeySequence,
                           QLinearGradient, QPainter, QPen, QShortcut)
from PySide6.QtWidgets import (
    QApplication,
    QDialog,
    QSplitter,
    QFrame,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QMenu,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QToolButton,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)

from ..classify import analyze, effective_role
from .. import __version__
from .. import schemes
from .dialogs import (
    CaptionNumberDialog,
    ROLE_TEXT,
    TableDialog,
    ParaFormatDialog,
    AboutDialog,
    IndentDialog,
    LayoutDialog,
    NumberingDialog,
    PageDialog,
    RoleDialog,
    RolesDialog,
)
from .pageview import PageView, PreviewScroll
from .widgets import PillButton
from .settings import (
    EDITABLE_ROLES,
    JOURNAL_DISABLED,
    ROLE_LABELS,
    Settings,
    ribbon_groups,
)
from .word_preview import WordPreview

NUM_KIND = {"num_heading": "heading", "num_figure": "figure", "num_table": "table"}
LAYOUT_KIND = {"layout_section": "section", "layout_pagenum": "pagenum",
               "layout_header": "header"}


def _as_list(x) -> list:
    """单个段号 / 列表 / 元组 / 集合，统一成列表（批量改格式用）。"""
    if isinstance(x, (list, tuple, set)):
        return list(x)
    return [x]


class _LogoIcon(QWidget):
    """左上角图标：圆角方块里一张纸 + 三条段落线（QPainter 画的，不是字符）。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(32, 32)

    def paintEvent(self, _e) -> None:
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        r = QRectF(self.rect()).adjusted(1, 1, -1, -1)
        g = QLinearGradient(r.topLeft(), r.bottomRight())
        g.setColorAt(0.0, QColor("#0a84ff"))
        g.setColorAt(1.0, QColor("#0058d0"))
        p.setPen(Qt.NoPen)
        p.setBrush(QBrush(g))
        p.drawRoundedRect(r, 9, 9)

        # 纸
        page = QRectF(r.left() + 7.5, r.top() + 5.5, r.width() - 15, r.height() - 11)
        p.setBrush(QColor("#ffffff"))
        p.drawRoundedRect(page, 2.5, 2.5)
        # 折角
        p.setBrush(QColor("#d7e8ff"))
        p.drawPolygon([
            page.topRight() + QPointF(-5.5, 0),
            page.topRight() + QPointF(0, 5.5),
            page.topRight() + QPointF(-5.5, 5.5),
        ])
        # 段落线（长短不一，最后一行短 = 段落结尾）
        p.setBrush(QColor("#0a84ff"))
        for i, ratio in enumerate((0.92, 0.72, 0.52)):
            y = page.top() + 4.6 + i * 4.6
            w = (page.width() - 5) * ratio
            p.drawRoundedRect(QRectF(page.left() + 2.5, y, w, 2.1), 1.05, 1.05)
        p.end()


class TitleBlock(QWidget):
    """标题区：图标 + 主标题 + 副标题 + 右侧渐隐圆点装饰带。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("title-block")
        self.setFixedHeight(46)
        # 横向撑开：右侧空档用它画渐隐点阵，不会显得空
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(24, 0, 24, 0)
        lay.setSpacing(11)
        lay.addWidget(_LogoIcon())
        col = QWidget()
        col.setObjectName("title-col")
        cl = QVBoxLayout(col)
        cl.setContentsMargins(0, 0, 0, 0)
        cl.setSpacing(1)
        self.title = QLabel("论文格式助手")
        self.title.setObjectName("app-title")
        self.sub = QLabel("中文学位论文 · 期刊论文　自动排版")
        self.sub.setObjectName("app-sub")
        cl.addWidget(self.title)
        cl.addWidget(self.sub)
        lay.addWidget(col)
        lay.addStretch(1)

    def paintEvent(self, e) -> None:
        super().paintEvent(e)
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        y = int(self.height() / 2)
        x0 = 350
        x1 = self.width() - 32
        if x1 <= x0:
            p.end()
            return
        n = int((x1 - x0) / 14)
        for i in range(n + 1):
            t = i / max(1.0, float(n))
            a = int(95 * (1.0 - t) ** 1.4)
            if a <= 8:
                continue
            p.setPen(QPen(QColor(10, 132, 255, a), 1.5))
            p.drawPoint(x0 + int(i * 14), y)
        p.end()


class PageJump(QWidget):
    """「跳至 [N] 页」：长文档不用一格格滚。"""

    def __init__(self, view, scroll, parent=None):
        super().__init__(parent)
        self.setObjectName("page-jump")
        self.view = view
        self.scroll = scroll
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(4)
        lab = QLabel("跳至")
        lab.setObjectName("jump-label")
        self.box = QSpinBox()
        self.box.setRange(1, 1)
        self.box.setFixedWidth(52)
        self.box.setAlignment(Qt.AlignCenter)
        self.box.setFocusPolicy(Qt.ClickFocus)
        self.box.setToolTip("输入页码后回车，直接跳到该页")
        of = QLabel("页")
        of.setObjectName("jump-label")
        lay.addWidget(lab)
        lay.addWidget(self.box)
        lay.addWidget(of)
        self._syncing = False
        self.box.valueChanged.connect(self._goto)
        scroll.verticalScrollBar().valueChanged.connect(self._on_scroll)

    def set_page_count(self, n: int) -> None:
        self.box.setRange(1, max(1, n))
        self._on_scroll(self.scroll.verticalScrollBar().value())

    def _goto(self, p: int) -> None:
        if self._syncing:
            return
        y = self.view.y_for_page(max(0, p - 1))
        vh = self.scroll.viewport().height()
        ph = self.view._paper[1] * self.view._zoom
        val = int(y - max(0.0, (vh - ph) / 2.0))
        self._syncing = True
        self.scroll.verticalScrollBar().setValue(max(0, val))
        self._syncing = False

    def _on_scroll(self, value: int) -> None:
        if self._syncing:
            return
        i = self.view.page_for_y(int(value))
        self._syncing = True
        self.box.setValue(max(1, i + 1))
        self._syncing = False


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"论文格式助手 v{__version__}")
        self._fit_window()

        self.path: Path | None = None
        self.doc = None
        self.analysis = None
        self.settings = Settings()
        # 记住上次用的方案：换电脑/发给别人时是空配置，仍是默认两套
        self.current_scheme = schemes.builtin(self.settings.profile)
        last = schemes.load_state()
        sc = schemes.find(last.get("last_name", ""),
                          last.get("last_kind")) if last.get("last_name") else None
        if sc is not None:
            self.settings = sc.settings_obj()
            self.current_scheme = sc
        self.preview: WordPreview | None = None
        self.applied = False          # 是否已经排过版（缓存里有结果）
        self.cached_out: Path | None = None
        self.overrides: dict = {}     # 段索引 → 手动格式（预览里点段落改的）
        self._jump: dict = {}         # 每个预览区的「跳至指定页」控件
        self.setAcceptDrops(True)    # 把 docx 拖进窗口直接打开

        self._build()

    # ---------------------------------------------------------------- #
    def _build(self) -> None:
        root = QWidget()
        root.setObjectName("root")
        outer = QVBoxLayout(root)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        outer.addWidget(self._topbar())
        self._ribbon_layout = QVBoxLayout()
        self._ribbon_layout.setContentsMargins(0, 0, 0, 0)
        self._ribbon_layout.setSpacing(0)
        outer.addLayout(self._ribbon_layout)
        self._ribbon_layout.addWidget(self._ribbon())

        split = QSplitter(Qt.Horizontal)
        split.setHandleWidth(1)
        self.before_view = PageView()
        self.after_view = PageView()
        self.after_view.set_paragraph_handler(self._para_menu)
        self._para_dlg = None
        self.before_head = QLabel("原稿")
        self.after_head = QLabel("排版后")
        self.sync_btn = PillButton("同步")
        self.sync_btn.setObjectName("sync")
        self.sync_btn.setCheckable(True)
        self.sync_btn.setToolTip("左右两栏一起滑动")
        self.sync_btn.clicked.connect(self.toggle_sync)
        self.before_wrap, self.before_scroll = self._pane(
            self.before_head, self.before_view, extra=self.sync_btn)
        self.after_wrap, self.after_scroll = self._pane(self.after_head, self.after_view)
        split.addWidget(self.before_wrap)
        split.addWidget(self.after_wrap)
        split.setSizes([680, 680])
        outer.addWidget(split, 1)

        self._connect_sync()
        self.setCentralWidget(root)
        self._build_profile_switch()
        # Esc = 取消段落多选
        QShortcut(QKeySequence("Esc"), self).activated.connect(
            self._clear_para_selection)
        self.statusBar().showMessage("就绪 — 先打开一个 docx")

    def _clear_para_selection(self) -> None:
        self.after_view.clear_selection()

    def eventFilter(self, obj, ev) -> bool:
        # 段落浮动格式条一消失就清空多选，避免下一次误把上一批再改一遍
        if obj is getattr(self, "_para_dlg", None) and ev.type() == QEvent.Type.Hide:
            self._clear_para_selection()
        return super().eventFilter(obj, ev)

    # ---------------------------------------------------------------- #
    def _pane(self, head: QLabel, view: QWidget, extra: QWidget = None
              ) -> tuple[QWidget, QScrollArea]:
        """返回 (外壳, 滚动区)；外壳必须留在 self 上，否则滚动区会被 GC 掉。"""
        wrap = QWidget()
        lay = QVBoxLayout(wrap)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)
        sa = PreviewScroll(view)
        head_row = QWidget()
        head_row.setFixedHeight(34)
        head_row.setObjectName("pane-head-row")
        hl = QHBoxLayout(head_row)
        hl.setContentsMargins(12, 0, 10, 0)
        hl.setSpacing(6)
        head.setObjectName("pane-head")
        hl.addWidget(head)
        hl.addStretch(1)
        if extra is not None:
            hl.addWidget(extra)
        j = PageJump(view, sa)
        self._jump[id(view)] = j
        hl.addWidget(j)
        hl.addWidget(self._zoom_row(view))
        lay.addWidget(head_row)
        lay.addWidget(sa, 1)
        return wrap, sa

    # ---------------------------------------------------------------- #
    def _zoom_row(self, view: PageView) -> QWidget:
        """缩放控件：− 比例 + ｜适应窗口｜适应整页（像 Word 那样）。"""
        w = QWidget()
        lay = QHBoxLayout(w)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(3)

        minus = QPushButton("−")
        minus.setObjectName("zoom-btn")
        minus.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        minus.setToolTip("缩小")
        pct = QLabel("100%")
        pct.setObjectName("zoom-pct")
        plus = QPushButton("＋")
        plus.setObjectName("zoom-btn")
        plus.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        plus.setToolTip("放大")
        lay.addWidget(minus)
        lay.addWidget(pct)
        lay.addWidget(plus)
        lay.addSpacing(6)

        fit_w = PillButton("适应窗口")
        fit_p = PillButton("适应整页")
        for b in (fit_w, fit_p):
            b.setObjectName("fit-btn")
            b.setCheckable(True)
        fit_w.setChecked(True)
        lay.addWidget(fit_w)
        lay.addWidget(fit_p)

        def show_zoom() -> None:
            pct.setText(f"{view.zoom_percent()}%")

        def step(d: int) -> None:
            base = view.zoom_percent() if view._fit_mode != "pct" else view._zoom_pct
            view.set_fit_mode("pct", base + d)
            show_zoom()
            fit_w.setChecked(False)
            fit_p.setChecked(False)

        minus.clicked.connect(lambda: step(-10))
        plus.clicked.connect(lambda: step(10))

        def use_window() -> None:
            view.set_fit_mode("window")
            show_zoom()
            fit_w.setChecked(True)
            fit_p.setChecked(False)

        def use_page() -> None:
            view.set_fit_mode("page")
            show_zoom()
            fit_p.setChecked(True)
            fit_w.setChecked(False)

        fit_w.clicked.connect(use_window)
        fit_p.clicked.connect(use_page)
        view.zoom_changed = show_zoom
        show_zoom()
        return w

    # ---------------- 标题行 ---------------- #
    def _topbar(self) -> QWidget:
        bar = QWidget()
        bar.setObjectName("topbar")
        bar.setFixedHeight(52)
        hl = QHBoxLayout(bar)
        hl.setContentsMargins(24, 0, 24, 0)

        title = TitleBlock()
        hl.addWidget(title)
        hl.addSpacing(18)

        self.open_btn = QPushButton("打开文档")
        self.open_btn.clicked.connect(self.open_file)
        hl.addWidget(self.open_btn)

        # 主体是普通按钮（文字居中），右侧紧贴一个小箭头按钮出二级菜单（0 间距容器）
        split_box = QWidget()
        split_box.setSizePolicy(
            QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        split_hl = QHBoxLayout(split_box)
        split_hl.setContentsMargins(0, 0, 0, 0)
        split_hl.setSpacing(0)
        self.detect_btn = QPushButton("从原稿识别")
        self.detect_btn.setObjectName("split-main")
        self.detect_btn.setEnabled(False)
        self.detect_btn.clicked.connect(self.open_detect_dialog)
        self.detect_arrow = QPushButton("▾")
        self.detect_arrow.setObjectName("split-arrow")
        detect_menu = QMenu(self.detect_arrow)
        act_cur = detect_menu.addAction("从当前文档识别…")
        act_cur.triggered.connect(self.open_detect_dialog)
        act_tpl = detect_menu.addAction("从模板文件识别…")
        act_tpl.triggered.connect(lambda: self.open_detect_dialog(template=True))
        detect_menu.addSeparator()
        act_save_sc = detect_menu.addAction("保存当前设置为方案…")
        act_save_sc.triggered.connect(self.save_current_as_scheme)
        act_scheme = detect_menu.addAction("管理格式方案…")
        act_scheme.triggered.connect(self.open_schemes)
        self.detect_arrow.setMenu(detect_menu)
        split_hl.addWidget(self.detect_btn)
        split_hl.addWidget(self.detect_arrow)
        hl.addWidget(split_box)

        self.compare_btn = QPushButton("在 Word 中对比")
        self.compare_btn.setEnabled(False)
        self.compare_btn.clicked.connect(self.open_compare)
        hl.addWidget(self.compare_btn)

        self.apply_btn = QPushButton("一键排版")
        self.apply_btn.setObjectName("primary")
        self.apply_btn.setEnabled(False)
        self.apply_btn.clicked.connect(self.apply)
        hl.addWidget(self.apply_btn)

        # 保存：一键排版只生成结果，落盘由这个按钮决定（覆盖 / 另存为）
        self.save_btn = QPushButton("保存")
        self.save_btn.setEnabled(False)
        self.save_btn.setToolTip("选择覆盖原文件或另存为新文件")
        save_menu = QMenu(self)
        act_ow = save_menu.addAction("覆盖原文件（自动备份 .bak.docx）")
        act_ow.triggered.connect(lambda: self.save("overwrite"))
        act_as = save_menu.addAction("另存为…")
        act_as.triggered.connect(lambda: self.save("as"))
        self.save_btn.setMenu(save_menu)
        hl.addWidget(self.save_btn)

        # 窗口偏窄时按钮收缩而不是被截断
        for b in (self.open_btn, self.detect_btn, self.detect_arrow,
                  self.compare_btn, self.apply_btn, self.save_btn):
            b.setMinimumWidth(0)
        return bar

    # ---------------- 格式方案（状态栏） ---------------- #
    def _build_profile_switch(self) -> None:
        """状态栏右侧：当前格式方案 + 二级菜单（勾选即用）。"""
        lab = QLabel("格式方案：")
        lab.setObjectName("seg-label")
        self.scheme_btn = QPushButton(self.current_scheme.label)
        self.scheme_btn.setObjectName("scheme-btn")
        self.scheme_btn.setToolTip("点一下选择要套用的格式方案")
        self.scheme_btn.clicked.connect(self._show_scheme_menu)
        holder = QWidget()
        hl = QHBoxLayout(holder)
        hl.setContentsMargins(0, 2, 8, 2)
        hl.setSpacing(6)
        hl.addWidget(lab)
        hl.addWidget(self.scheme_btn)
        self.statusBar().addPermanentWidget(holder)

    def _update_scheme_btn(self) -> None:
        self.scheme_btn.setText(self.current_scheme.label)
        self.scheme_btn.setToolTip(
            f"当前方案：{self.current_scheme.label}\n点一下选择其他方案")

    def _scheme_menu(self) -> QMenu:
        menu = QMenu(self)
        group = QActionGroup(menu)
        group.setExclusive(True)
        for kind in schemes.KINDS:
            menu.addSection(schemes.KIND_LABEL[kind])
            items = [schemes.builtin(kind)]
            items += [s for s in schemes.load_user_schemes() if s.kind == kind]
            for sc in items:
                act = group.addAction(sc.label)
                act.setCheckable(True)
                act.setChecked(sc.name == self.current_scheme.name
                               and sc.kind == self.current_scheme.kind)
                act.triggered.connect(
                    lambda _=False, s=sc: self._apply_scheme(s))
        menu.addSeparator()
        if not self.current_scheme.builtin:
            menu.addAction(f"更新方案「{self.current_scheme.name}」",
                           self.update_current_scheme)
            menu.addAction("另存当前设置为新方案…", self.save_current_as_scheme)
        else:
            menu.addAction("保存当前设置为方案…", self.save_current_as_scheme)
        menu.addAction("管理方案…", self.open_schemes)
        menu.addAction("从模板导入…", lambda: self.open_detect_dialog(template=True))
        return menu

    def _show_scheme_menu(self) -> None:
        self._scheme_menu().exec(self.scheme_btn.mapToGlobal(
            QPoint(0, self.scheme_btn.height())))

    def _apply_scheme(self, sc) -> None:
        self.settings = sc.settings_obj()
        self.settings._profiles = {}      # 旧快照作废，避免切换时换回旧值
        self.current_scheme = sc
        self._update_scheme_btn()
        self._rebuild_ribbon()
        if self.doc is not None:
            self.refresh_preview()
        schemes.save_state(sc.name, sc.kind)
        self.statusBar().showMessage(f"已套用方案「{sc.name}」")

    # ---------------- Ribbon ---------------- #
    def _ribbon(self) -> QWidget:
        wrap = QWidget()
        wrap.setObjectName("ribbon")
        lay = QVBoxLayout(wrap)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll.setStyleSheet("QScrollArea{border:none;background:transparent;}")

        inner = QWidget()
        hl = QHBoxLayout(inner)
        hl.setContentsMargins(12, 8, 12, 8)
        hl.setSpacing(10)

        for group_name, buttons in ribbon_groups(self.settings.profile):
            group = QFrame()
            group.setObjectName("ribbon-group")
            gl = QVBoxLayout(group)
            gl.setContentsMargins(6, 5, 6, 5)
            gl.setSpacing(4)
            for key, label in buttons:
                b = QPushButton(label)
                b.setObjectName("ribbon-btn")
                b.setFixedHeight(26)
                # 期刊论文没有目录：卡片照样显示，但画成灰色、
                # 指上去有说明、点击会解释原因（不能用 setEnabled(False)，
                # Qt 对禁用控件不投递鼠标事件，tooltip 永远不显示）
                tip = JOURNAL_DISABLED.get(key)
                if tip and self.settings.profile == "journal":
                    b.setProperty("locked", "true")
                    b.setToolTip(tip + "（当前论文形态不可用）")
                    st = b.style()
                    st.unpolish(b)
                    st.polish(b)
                b.clicked.connect(lambda _=False, k=key: self.on_ribbon(k))
                gl.addWidget(b)
            # 标题固定贴底：各组按钮数不一样多时，
            # 组名仍然对齐在同一条线上（旧版跟着按钮走，参差不齐）
            gl.addStretch(1)
            lab = QLabel(group_name)
            lab.setObjectName("ribbon-group-title")
            lab.setAlignment(Qt.AlignHCenter)
            gl.addWidget(lab)
            # 每个组都带拉伸系数：7 个组均分整行宽度，
            # 否则全挤在左边、右侧留一大片空白
            hl.addWidget(group, 1)

        scroll.setWidget(inner)
        lay.addWidget(scroll)

        def _fit():
            inner.adjustSize()
            h = inner.sizeHint().height()
            if h > 0:
                scroll.setFixedHeight(h)

        QTimer.singleShot(0, _fit)
        QTimer.singleShot(80, _fit)
        return wrap

    def _rebuild_ribbon(self) -> None:
        """切换论文形态后重建 Ribbon（期刊没有分节、目录是置灰的）。"""
        old = self._ribbon_layout.itemAt(0)
        w = old.widget() if old is not None else None
        if w is not None:
            self._ribbon_layout.removeWidget(w)
            w.setParent(None)
            w.deleteLater()
        self._ribbon_layout.addWidget(self._ribbon())


    # ---------------------------------------------------------------- #
    def on_ribbon(self, key: str) -> None:
        s = self.settings
        # 期刊形态下显示但不可用的卡片：点一下说清楚原因
        tip = JOURNAL_DISABLED.get(key)
        if tip and s.profile == "journal":
            QMessageBox.information(self, "该卡片不可用", tip)
            self.statusBar().showMessage(tip)
            return
        if key in EDITABLE_ROLES:
            dlg = RoleDialog(ROLE_LABELS.get(key, key), s.roles[key], self)
            if dlg.exec():
                s.roles[key].update(dlg.value())
        elif key == "page":
            dlg = PageDialog(s.page, self)
            if dlg.exec():
                s.page.update(dlg.value())
        elif key in ("indent", "not_indent"):
            dlg = IndentDialog(s, self, focus_not_indent=(key == "not_indent"))
            if dlg.exec():
                dlg.apply_to(s)
        elif key in ("num_heading", "num_figure", "num_table"):
            dlg = (CaptionNumberDialog(s, self, kind=NUM_KIND[key])
                   if key in ("num_figure", "num_table")
                   else NumberingDialog(s, self, kind=NUM_KIND[key]))
            if dlg.exec():
                dlg.apply_to(s)
        elif key == "table_style":
            dlg = TableDialog(s, self)
            if dlg.exec():
                dlg.apply_to(s)
        elif key in ("layout_section", "layout_pagenum", "layout_header"):
            dlg = LayoutDialog(s, self, kind=LAYOUT_KIND[key])
            if dlg.exec():
                dlg.apply_to(s)
        elif key == "roles":
            if self.analysis is None:
                QMessageBox.information(self, "识别明细", "请先打开一个文档。")
                return
            dlg = RolesDialog(self.analysis, self.doc, self)
            if dlg.exec():
                dlg.apply_to(self.analysis)
                self._after_change()
                self.statusBar().showMessage(
                    f"识别明细已更新：手工改正 {dlg.changed} 段"
                )
            return
        elif key == "schemes":
            self.open_schemes()
            return
        elif key == "check":
            self.run_check()
            return
        elif key == "about":
            AboutDialog(self).exec()
            return
        else:
            return
        self._after_change()

    def _after_change(self) -> None:
        """设置改动后：立刻刷新内嵌预览；如果 Word 对比开着，也刷新它。"""
        self.refresh_preview()
        if self.preview is not None and self.preview.ready:
            self.preview.refresh(self._render)
            self.statusBar().showMessage("内嵌预览与 Word 预览均已更新（原稿只读，未改动）")

    def refresh_preview(self) -> None:
        """用 QTextDocument 分页渲染左右两栏。"""
        if self.doc is None or self.analysis is None:
            return
        from ..preview_render import derive_cut_points
        from .render import page_count, page_geometry, page_labels, render_side

        paper, margins, content = page_geometry(self.settings)
        profile = getattr(self.settings, "profile", "thesis")
        _, body = derive_cut_points(self.analysis) if profile == "thesis" else (None, None)

        before, b_blocks = render_side(self.doc, self.analysis, self.settings,
                                       "before", content)
        after, a_blocks = render_side(self.doc, self.analysis, self.settings,
                                      "after", content, body_index=body,
                                      overrides=self.overrides)

        nb, na = page_count(before), page_count(after)
        nums = self.settings.enable_page_numbers
        b_labels = page_labels(before, b_blocks, body, content[1], nb,
                               profile=profile, show_numbers=nums)
        a_labels = page_labels(after, a_blocks, body, content[1], na,
                               profile=profile, show_numbers=nums)

        self.before_view.set_paper(paper, margins)
        self.after_view.set_paper(paper, margins)
        self.before_view.set_document(before, b_labels, header="", show_header=False,
                                      show_pagenum=nums, page_count=nb)
        self.after_view.set_document(
            after, a_labels,
            header=self.settings.header_text if self.settings.enable_header else "",
            show_header=bool(self.settings.enable_header),
            show_pagenum=nums,
            page_count=na,
            src_of_block={b.position(): i for i, b in a_blocks.items()},
        )
        self.after_view.set_header_line(
            self.settings.header_line, self.settings.header_line_color,
            getattr(self.settings, "header_line_style", "single"))
        self.after_view.set_header_style(
            getattr(self.settings, "header_font", "宋体"),
            getattr(self.settings, "header_ascii", "Times New Roman"),
            getattr(self.settings, "header_size", 10.5),
            getattr(self.settings, "header_bold", False),
            getattr(self.settings, "header_color", ""),
            getattr(self.settings, "header_gap", 4.0))

        n_ov = len(self.overrides)
        self.before_head.setText(f"原稿（{nb} 页）")
        self.after_head.setText(
            f"排版后（{na} 页）" + (f"｜{n_ov} 处手动" if n_ov else ""))
        for v, n in ((self.before_view, nb), (self.after_view, na)):
            j = self._jump.get(id(v))
            if j is not None:
                j.set_page_count(n)
        self.statusBar().showMessage(
            f"{self.path.name} ｜ 段落 {self.analysis.paragraph_count} ｜ "
            f"原稿 {nb} 页 / 排版后 {na} 页"
        )

    # ---------------------------------------------------------------- #
    def _connect_sync(self) -> None:
        """左右栏同步滚动：按页对齐（两栏页数不同，按像素没意义）。"""
        self.sync_on = False
        self._syncing = False
        for src, dst in ((self.before_scroll, self.after_scroll),
                         (self.after_scroll, self.before_scroll)):
            src.verticalScrollBar().valueChanged.connect(
                lambda v, s=src, d=dst: self._on_scroll(s, d, v))

    def _on_scroll(self, src: QScrollArea, dst: QScrollArea, value: int) -> None:
        """按相对位置同步：两边滚动条按比例走，滑起来是齐的。"""
        if not self.sync_on or self._syncing:
            return
        sb = src.verticalScrollBar()
        rng = max(1, sb.maximum())
        frac = max(0.0, min(1.0, value / rng))
        self._syncing = True
        try:
            dsb = dst.verticalScrollBar()
            dsb.setValue(int(frac * max(1, dsb.maximum())))
        finally:
            self._syncing = False

    def toggle_sync(self) -> None:
        self.sync_on = not self.sync_on
        self.sync_btn.setChecked(self.sync_on)
        # 中性灰的"开"，不用主色——主色只给「一键排版」这个级别的动作
        self.sync_btn.setText("同步中" if self.sync_on else "同步")
        self.statusBar().showMessage(
            "同步滚动已开启：左右两栏一起滑动" if self.sync_on else "同步滚动已关闭")

    # ---------------------------------------------------------------- #
    def _para_menu(self, src: int, global_pos, block=None) -> None:
        """点中排版后预览里的某一段：弹浮动格式条，改动即时生效。

        Ctrl 多选过若干段时，再点其中任意一段 = 批量改这些段。
        """
        from .dialogs import ParaFormatDialog

        sel = self.after_view.selected_srcs()
        if src in sel and len(sel) > 1:
            targets = sorted(sel)
        else:
            # 点了一个没选中的段：改成只改它，并清掉旧的多选
            targets = [src]
            sel = {src}
        self.after_view.set_selected_srcs(sel)

        role = effective_role(self.analysis.detections[src])
        base = dict(self.settings.roles.get(role) or self.settings.roles.get("body") or {})
        cur = self.overrides.get(src) or base

        # 先把这一段高亮出来，用户才知道自己点的是哪段
        if block is not None:
            self.after_view.set_selected_block(block)

        dlg = ParaFormatDialog(src, cur, self)
        dlg.set_para_indices(targets)
        dlg.set_value(cur)
        dlg.set_role(role)
        _preview_txt = self.analysis.features[src].text.strip().replace("\n", " ")[:26]
        head = (f"已选 {len(targets)} 段" if len(targets) > 1
                else f"第 {src + 1} 段 · {ROLE_LABELS.get(role, role)}")
        dlg.set_title(
            head
            + ("（已手动改）" if src in self.overrides else "")
            + (f"\n「{_preview_txt}」" if _preview_txt else ""))
        dlg.adjustSize()
        x = global_pos.x()
        y = global_pos.y()
        scr = QApplication.screenAt(global_pos) or QApplication.primaryScreen()
        avail = scr.availableGeometry()
        w, h = dlg.width(), dlg.height()
        if x + w > avail.right():
            x = avail.right() - w - 4
        if y + h > avail.bottom():
            y = max(avail.top(), y - h - 8)
        dlg.move(x, y)

        dlg.on_change = self._override_changed
        dlg.on_clear = self._override_clear
        dlg.on_clear_all = self._override_clear_all
        dlg.on_role = self._apply_role_to_para
        dlg.show()
        dlg.raise_()
        dlg.activateWindow()
        self._para_dlg = dlg
        # 浮动条一消失就清空多选，避免下一次误把上一批再改一遍
        dlg.installEventFilter(self)
        self.after_view.on_selection_changed = None

    def _apply_role_to_para(self, indices, role: str) -> None:
        """把一段（或 Ctrl 多选的若干段）直接改成某个角色。

        除了写手动格式，还要把 analysis 里的角色一起改掉——否则编号不会重算
        （用户实测：把某段设成二级标题后，标题号还是老的）。
        """
        p = self.settings.roles.get(role)
        if not p:
            return
        targets = _as_list(indices)
        for i in targets:
            self.overrides[i] = dict(p)
            d = self.analysis.detections[i]
            d.role = role
            d.level = (int(role[-1])
                       if role.startswith("heading") and role != "heading" else None)
            d.level_confidence = 1.0
            d.user_confirmed = True
            if "手工改正" not in d.reasons:
                d.reasons = list(d.reasons) + ["手工改正"]
        dlg = self._para_dlg
        if isinstance(dlg, ParaFormatDialog):
            dlg.set_value(p)
            dlg.set_title(f"已把 {len(targets)} 段设为「{ROLE_TEXT.get(role, role)}」")
        self._after_change()
        self.statusBar().showMessage(
            f"{len(targets)} 段已设为「{ROLE_TEXT.get(role, role)}」，编号已重算")

    def _override_changed(self, indices, fmt: dict) -> None:
        """indices 可以是单个段号，也可以是列表（Ctrl 多选批量改）。"""
        for i in _as_list(indices):
            self.overrides[i] = fmt
        self.refresh_preview()

    def _override_clear(self, indices) -> None:
        n = 0
        for i in _as_list(indices):
            if self.overrides.pop(i, None) is not None:
                n += 1
        self.refresh_preview()
        self.statusBar().showMessage(
            f"已恢复 {n} 段的角色默认格式" if n else "这些段本来就没有手动格式")

    def _override_clear_all(self) -> None:
        n = len(self.overrides)
        self.overrides.clear()
        self.refresh_preview()
        self.statusBar().showMessage(f"已清除全部手动格式（{n} 处）")

    # ---------------------------------------------------------------- #
    def _render(self, out_path: Path) -> None:
        from ..preview_render import render_preview

        render_preview(self.path, self.settings, out_path, overrides=self.overrides)

    def open_file(self) -> None:
        from PySide6.QtWidgets import QFileDialog

        path, _ = QFileDialog.getOpenFileName(
            self, "打开 Word 文档", "", "Word 文档 (*.docx *.doc)")
        if not path:
            return
        self.open_path(Path(path))

    def open_path(self, path: Path) -> bool:
        """打开指定路径的 docx/doc（按钮、命令行参数、拖拽共用）。"""
        suf = path.suffix.lower()
        if suf not in (".docx", ".doc"):
            QMessageBox.information(
                self, "打不开这个文件",
                "只支持 .docx 或 .doc 格式的 Word 文档。",
            )
            return False
        self._orig_path = path
        if suf == ".doc":
            self.statusBar().showMessage(f"正在转换旧版 .doc：{path.name} …")
            QApplication.processEvents()
            try:
                from ..features import ensure_docx
                docx_path = ensure_docx(path)
            except Exception as exc:  # noqa: BLE001
                QMessageBox.critical(self, "转换 .doc 失败", str(exc))
                return False
            self.path = docx_path
        else:
            self.path = path
        try:
            from docx import Document

            self.doc = Document(str(self.path))
            self.analysis = analyze(self.doc, source=path.name)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "打开失败", str(exc))
            return False

        self.applied = False
        self.cached_out = None
        self.save_btn.setEnabled(False)
        self.detect_btn.setEnabled(True)
        self.detect_arrow.setEnabled(True)
        self.compare_btn.setEnabled(True)
        self.apply_btn.setEnabled(True)
        # 关键：打开后立刻渲染双栏预览，否则两栏是空的
        self.refresh_preview()

        c = self.analysis.summary()["by_role"]
        heads = self.analysis.summary()["headings"]
        self.statusBar().showMessage(
            f"已打开 {self.path.name} ｜ {self.analysis.paragraph_count} 段 ｜ "
            f"标题 {heads} ｜ 图 {c.get('figure', 0)} ｜ 图注 {c.get('figure_caption', 0)} ｜ "
            f"文献 {c.get('reference', 0)} ｜ 编号体系 {self.analysis.scheme}"
        )
        return True

    # ---------------------------------------------------------------- #
    # 拖拽：把 docx 拖进窗口（任意位置）直接打开
    # ---------------------------------------------------------------- #
    def _drop_docx(self, event) -> Path | None:
        """拖拽里的第一个本地 .docx / .doc；不是就返回 None。"""
        for u in event.mimeData().urls():
            if not u.isLocalFile():
                continue
            p = Path(u.toLocalFile())
            return p if p.suffix.lower() in (".docx", ".doc") else None
        return None

    def _drop_hl(self, on: bool, name: str = "") -> None:
        """拖拽悬停时的视觉反馈：内容区淡蓝 + 状态栏提示。"""
        cw = self.centralWidget()
        if cw is not None:
            cw.setStyleSheet("background:#e8f1ff;" if on else "")
        if on:
            self.statusBar().showMessage(f"松开鼠标，打开 {name}")

    def handle_drag_event(self, event) -> bool:
        """处理拖拽事件（由 App.notify() 转发过来）。返回 True 表示已消费。"""
        t = event.type()
        if t == QEvent.Type.DragEnter:
            p = self._drop_docx(event)
            if p is None:
                event.ignore()
            else:
                event.acceptProposedAction()
                self._drop_hl(True, p.name)
            return True
        if t == QEvent.Type.DragMove:
            p = self._drop_docx(event)
            if p is None:
                event.ignore()
            else:
                event.acceptProposedAction()
            return True
        if t == QEvent.Type.DragLeave:
            self._drop_hl(False)
            return True
        if t == QEvent.Type.Drop:
            p = self._drop_docx(event)
            if p is None:
                event.ignore()
            else:
                event.acceptProposedAction()
                self._drop_hl(False)
                self.open_path(p)
            return True
        return False

    def _monitor_work(self, hwnd) -> tuple[tuple[float, float, float, float], float] | None:
        """窗口**所在那台显示器**的可用区（物理像素）+ 真实缩放比。

        不能用 SPI_GETWORKAREA——它只返回主显示器。这台机器是双屏
        （主 2880×1800、副 QHD 1700×2560），窗口落在副屏时按主屏居中，
        又会有一截跑到屏幕外。
        """
        import ctypes

        user32 = ctypes.WinDLL("user32", use_last_error=True)

        class RECT(ctypes.Structure):
            _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                        ("right", ctypes.c_long), ("bottom", ctypes.c_long)]

        class MONITORINFO(ctypes.Structure):
            _fields_ = [("cbSize", ctypes.c_uint32), ("rcMonitor", RECT),
                        ("rcWork", RECT), ("dwFlags", ctypes.c_uint32)]

        hmon = user32.MonitorFromWindow(ctypes.c_void_p(hwnd), 2)   # DEFAULTTONEAREST
        mi = MONITORINFO()
        mi.cbSize = ctypes.sizeof(MONITORINFO)
        if not hmon or not user32.GetMonitorInfoW(ctypes.c_void_p(hmon),
                                                  ctypes.byref(mi)):
            return None
        cl = RECT()
        if not user32.GetClientRect(ctypes.c_void_p(hwnd), ctypes.byref(cl)):
            return None
        scale = (cl.right - cl.left) / max(1.0, float(self.width()))
        if scale <= 0.2 or scale > 8.0:
            return None
        w = mi.rcWork
        return (float(w.left), float(w.top),
                float(w.right - w.left), float(w.bottom - w.top)), scale

    def _win_rect(self, hwnd) -> tuple[float, float, float, float]:
        import ctypes

        user32 = ctypes.WinDLL("user32", use_last_error=True)

        class RECT(ctypes.Structure):
            _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                        ("right", ctypes.c_long), ("bottom", ctypes.c_long)]

        r = RECT()
        user32.GetWindowRect(ctypes.c_void_p(hwnd), ctypes.byref(r))
        return (float(r.left), float(r.top),
                float(r.right - r.left), float(r.bottom - r.top))

    def _fit_window(self) -> None:
        """按屏幕可用区决定窗口大小，并保证整块窗口落在屏幕里。

        坑（实测）：Qt 报的 dpr 与 Windows 实际缩放不一致（Qt 说 1.75，
        系统按 1.5 缩放），拿 QScreen::availableGeometry() 直接乘 0.94
        会算出比屏幕还大的窗口——右半截被挤出屏幕，鼠标点不到
        「关键词 / 目录 / 参考文献」这些靠右的 Ribbon 按钮。
        所以这里用 Win32 量"物理像素 / 逻辑像素"的真实比例来换算。
        """
        self.resize(1400, 860)
        try:
            self.winId()                       # 先建窗口句柄，才能量到真实尺寸
            got = self._monitor_work(int(self.winId()))
            if got is None:
                return
            (wx, wy, ww, wh), scale = got
            self.resize(max(1000, min(1500, int((ww / scale) * 0.94))),
                        max(680, min(960, int((wh / scale) * 0.90))))
        except Exception:
            pass

    def showEvent(self, event) -> None:
        super().showEvent(event)
        # 窗口真正显示后才会按我们的要求摆位置（show 之前 move 会被系统改掉）
        if not getattr(self, "_placed", False):
            self._placed = True
            self._place_on_screen()
            # 系统的默认摆位可能发生在 showEvent 之后，再补一次
            QTimer.singleShot(120, self._place_on_screen)

    def _place_on_screen(self) -> None:
        """把窗口收缩进屏幕可用区并居中（标题栏也占高，必须量真实外框）。"""
        try:
            hwnd = int(self.winId())
            got = self._monitor_work(hwnd)
            if got is None:
                return
            (wx, wy, ww, wh), scale = got
            for _ in range(4):
                rx, ry, rw, rh = self._win_rect(hwnd)
                over_w = rw - ww
                over_h = rh - wh
                if over_w <= 0 and over_h <= 0:
                    break
                self.resize(max(900, int(self.width() - over_w / scale)),
                            max(640, int(self.height() - over_h / scale)))
            rx, ry, rw, rh = self._win_rect(hwnd)
            self.move(int((wx + (ww - rw) / 2) / scale),
                      int((wy + (wh - rh) / 2) / scale))
        except Exception:
            pass

    def open_detect_dialog(self, template: bool = False) -> None:
        """「格式识别」弹窗：从当前文档或模板文件学格式，核对后再保存/覆盖。"""
        if self.doc is None or self.analysis is None:
            QMessageBox.information(self, "格式识别", "请先打开一个文档。")
            return
        from .scheme_dialogs import DetectDialog

        dlg = DetectDialog(self.doc, self.analysis,
                           base_settings=self.settings, parent=self)
        if template:
            dlg.tpl_radio.setChecked(True)
            dlg._pick()
            if not dlg.template_path:
                return
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        if dlg.scheme is not None:
            self._apply_scheme(dlg.scheme)
        elif dlg.settings is not None:
            self.settings = dlg.settings
            self.current_scheme = schemes.builtin(self.settings.profile)
            self._update_scheme_btn()
            self._after_change()
            self.statusBar().showMessage("已把识别结果写入当前设置")

    def open_schemes(self) -> None:
        """「格式方案」管理弹窗。"""
        from .scheme_dialogs import FormatSchemeDialog

        dlg = FormatSchemeDialog(self.current_scheme, parent=self,
                                 import_handler=self._import_scheme,
                                 save_current_handler=self.save_current_as_scheme)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        if dlg.current is not None:
            self._apply_scheme(dlg.current)

    def save_current_as_scheme(self):
        """把当前手动调好的整套设置存成格式方案（不经过格式识别器）。"""
        from .scheme_dialogs import SaveSchemeDialog

        dlg = SaveSchemeDialog(self.current_scheme, self.settings, parent=self)
        if dlg.exec() != QDialog.DialogCode.Accepted or dlg.scheme is None:
            return None
        sc = dlg.scheme
        self.current_scheme = sc
        self.settings.profile = sc.kind
        self._update_scheme_btn()
        self._rebuild_ribbon()
        schemes.save_state(sc.name, sc.kind)
        self.statusBar().showMessage(f"已将当前设置保存为方案「{sc.name}」")
        return sc

    def update_current_scheme(self) -> None:
        """把当前修改直接写回正在使用的用户方案。"""
        sc = self.current_scheme
        if sc.builtin:
            self.save_current_as_scheme()
            return
        payload = self.settings.scheme_payload()
        payload["profile"] = sc.kind
        sc.settings = payload
        schemes.save(sc)
        self.statusBar().showMessage(f"已更新方案「{sc.name}」")

    def _import_scheme(self):
        """给方案弹窗用：打开格式识别，返回新建的方案（取消返回 None）。"""
        if self.doc is None or self.analysis is None:
            QMessageBox.information(self, "格式识别", "请先打开一个文档。")
            return None
        from .scheme_dialogs import DetectDialog

        dlg = DetectDialog(self.doc, self.analysis,
                           base_settings=self.settings, parent=self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return None
        if dlg.scheme is not None:
            return dlg.scheme
        if dlg.settings is not None:
            self.settings = dlg.settings
            self._after_change()
        return None

    # ---------------------------------------------------------------- #
    def open_compare(self) -> None:
        if self.path is None:
            return
        if self.preview is None:
            self.preview = WordPreview(self.path, self)
        ok = self.preview.open(self._render)
        if ok:
            self.statusBar().showMessage(
                f"Word 对比已开启（原稿只读）｜预览目录：{self.preview.dir_path}"
            )

    def run_check(self) -> None:
        if self.doc is None:
            QMessageBox.information(self, "格式检查", "请先打开一个文档。")
            return
        from ..check import check_document

        findings = check_document(self.doc, self.analysis, self.settings)
        if not findings:
            QMessageBox.information(self, "格式检查", "✓ 未发现明显问题。")
            return
        text = "\n".join(f"• {f}" for f in findings[:40])
        more = f"\n… 其余 {len(findings) - 40} 条" if len(findings) > 40 else ""
        QMessageBox.information(self, "格式检查", f"发现 {len(findings)} 条：\n\n{text}{more}")

    # ---------------------------------------------------------------- #
    def _reload_source(self) -> None:
        """重新读原稿，保证「一键排版」可以反复点（幂等）。"""
        from docx import Document

        self.doc = Document(str(self.path))
        self.analysis = analyze(self.doc, source=self.path.name)

    def apply(self) -> None:
        """一键排版：只生成结果并刷新预览，不弹保存框。

        保存交给顶栏的「保存」按钮——用户自己决定覆盖原文件还是另存为。
        """
        if self.doc is None or self.path is None:
            return
        from ..numbering import (
            apply_plans,
            plan_caption_numbers,
            plan_heading_numbers,
        )

        # 已经排过一次的话，先从原稿重来，避免编号叠加
        if getattr(self, "applied", False):
            self._reload_source()

        from .preview import PreviewDialog

        front = body = None
        if self.settings.profile == "thesis":
            for d in self.analysis.detections:
                if front is None and d.role in ("abstract", "keywords", "toc"):
                    front = d.index
                if (front is not None and body is None and d.index > front
                        and (d.role == "heading" or d.role.startswith("heading"))):
                    body = d.index
        if self.settings.profile == "thesis" and (front is None or body is None):
            QMessageBox.warning(
                self, "无法定位分节点",
                "没有识别到“摘要/目录”，或找不到正文起点（摘要之后的第一个标题）。",
            )
            return

        plans = []
        if self.settings.enable_number_heading:
            plans += plan_heading_numbers(self.analysis.detections, body,
                                          self.settings)
        if self.settings.enable_number_figure:
            plans += plan_caption_numbers(self.analysis.detections, "figure",
                                          self.settings, self.analysis.features)
        if self.settings.enable_number_table:
            plans += plan_caption_numbers(self.analysis.detections, "table",
                                          self.settings, self.analysis.features)
        if plans and PreviewDialog(self.analysis, plans, self).exec() != 1:
            self.statusBar().showMessage("已取消")
            return

        try:
            from ..reflow import apply_formats
            from ..sections import apply as apply_sections
            from ..sections import diagnose, plans_from_settings

            if self.settings.profile == "thesis" and (
                self.settings.enable_sections or self.settings.enable_page_numbers
            ):
                apply_sections(self.doc, front_index=front, body_index=body,
                               plans=plans_from_settings(self.settings),
                               settings=self.settings)
            elif self.settings.profile == "journal":
                # 期刊论文：不切分，只统一页码格式/页眉
                from ..sections import configure

                configure(self.doc, plans_from_settings(self.settings), self.settings)
            from ..reflow import indent_chars_of, style_tables

            stats = apply_formats(self.doc, self.analysis, self.settings.roles,
                                    overrides=self.overrides,
                                    label_bold=bool(self.settings.bold_label),
                                    indent_chars=indent_chars_of(self.settings))
            style_tables(self.doc, self.settings)
            if self.settings.enable_marks:
                from ..marks import apply_marks
                apply_marks(self.doc)
            n_num = apply_plans(self.doc, plans) if plans else 0
            self._write_cache()
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "排版失败", str(exc))
            return

        self.applied = True
        self.refresh_preview()
        self.after_head.setText(f"排版后（{self.after_view.page_count()} 页）· 已排版")
        self.save_btn.setEnabled(True)
        stats_text = "、".join(
            f"{ROLE_LABELS.get(k, k)} {v} 段" for k, v in sorted(stats.items(), key=lambda x: -x[1])
        ) or "无"
        secs = " ｜ ".join(
            f"第{s['index']}节 {s['header'] or '无页眉'} {s['pgNumType_fmt'] or '无页码'}"
            for s in diagnose(self.doc)["sections"]
        )
        self.statusBar().showMessage(
            f"排版完成：{stats_text} ｜ 编号 {n_num} 个 ｜ {secs} ｜ "
            "结果已缓存，点「保存」决定去向"
        )

    # ---------------------------------------------------------------- #
    def _cache_dir(self) -> Path:
        if getattr(self, "_cache", None) is None:
            import tempfile

            self._cache = Path(tempfile.mkdtemp(prefix="thesis_out_"))
        return self._cache

    def _write_cache(self) -> Path:
        """把排版结果写到内部缓存文件，供预览对比与保存使用。"""
        out = self._cache_dir() / "排版结果.docx"
        self.doc.save(str(out))
        self.cached_out = out
        return out

    def save(self, mode: str) -> None:
        """用户明确点「保存」之后才落盘。mode = 'overwrite' | 'as'。"""
        cached = getattr(self, "cached_out", None)
        if cached is None:
            QMessageBox.information(self, "保存", "请先点「一键排版」生成结果。")
            return
        orig = getattr(self, "_orig_path", None) or self.path
        if mode == "overwrite":
            if orig.suffix.lower() == ".doc":
                target = orig.with_name(orig.stem + "_排版.docx")
            else:
                target = orig
        else:
            from PySide6.QtWidgets import QFileDialog

            chosen, _ = QFileDialog.getSaveFileName(
                self, "另存为",
                str(orig.with_name(orig.stem + "_排版.docx")),
                "Word 文档 (*.docx)",
            )
            if not chosen:
                return
            target = Path(chosen)
        try:
            import shutil

            if target == orig and orig.suffix.lower() == ".docx":
                backup = orig.with_suffix(".bak.docx")
                if not backup.exists():
                    shutil.copyfile(orig, backup)
            shutil.copyfile(cached, target)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "保存失败", str(exc))
            return
        self.statusBar().showMessage(f"已保存到 {target}")

    # ---------------------------------------------------------------- #
    def closeEvent(self, event) -> None:
        if self.preview is not None:
            self.preview.cleanup()
            self.preview = None
        if getattr(self, "_cache", None) is not None:
            import shutil

            shutil.rmtree(self._cache, ignore_errors=True)
        super().closeEvent(event)
