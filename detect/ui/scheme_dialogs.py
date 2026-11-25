"""方案库与格式识别两个弹窗（草图见 README/对话）。

- `FormatSchemeDialog`：格式方案管理。左栏按「期刊论文 / 学位论文」分组列出方案，
  右栏显示选中方案的摘要；可应用、从模板导入、重命名、删除、导出。
- `DetectDialog`：格式识别。选来源（当前文档 / 模板文件）→ 识别 →
  核对表（未识别的标出来，可改）→ 保存为方案或直接覆盖当前设置。
"""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtGui import QBrush, QColor, QPalette
from PySide6.QtWidgets import (
    QButtonGroup,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QDoubleSpinBox,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QRadioButton,
    QTableWidget,
    QTableWidgetItem,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from .. import schemes
from ..recognize import recognize_from_doc, settings_from_report
from .dialogs import ALIGNS, FONTS, LATIN_FONTS
from .settings import ROLE_LABELS
from .widgets import MarkCheck


DETECT_ROLES = ["title", "heading1", "heading2", "heading3", "body",
                "abstract", "abstract_body", "keywords", "meta", "toc",
                "figure_caption", "table_caption", "equation", "reference",
                "back"]

LINE_CHOICES = [("single", "单倍"), ("multiple", "1.5 倍"), ("multiple2", "2 倍"),
                ("minimum", "最小值"), ("exact", "固定磅值")]

SUMMARY_ROLES = ["title", "heading1", "heading2", "body", "reference"]


def _scheme_summary(sc: schemes.Scheme) -> str:
    s = sc.settings_obj()
    lines = [f"当前方案：{sc.label}", ""]
    p = s.page
    lines.append(f"页边距：上 {p.get('top', 0):g}  下 {p.get('bottom', 0):g}  "
                 f"左 {p.get('left', 0):g}  右 {p.get('right', 0):g} mm")
    lines.append("")
    for role in SUMMARY_ROLES:
        r = s.roles.get(role)
        if not r:
            continue
        parts = [f"{float(r.get('size', 12)):g} 磅"]
        parts.append(f"{r.get('font', '宋体')} / {r.get('ascii', 'Times New Roman')}")
        if r.get("bold"):
            parts.append("加粗")
        rule = r.get("line_rule") or "exact"
        value = r.get("line", 22)
        if rule == "multiple":
            parts.append("1.5 倍行距")
        elif rule == "multiple2":
            parts.append("2 倍行距")
        elif rule == "single":
            parts.append("单倍行距")
        elif rule == "minimum":
            parts.append(f"最小 {value:g} 磅")
        else:
            parts.append(f"{value:g} 磅")
        parts.append({"left": "左对齐", "center": "居中", "right": "右对齐",
                      "both": "两端对齐"}.get(r.get("align", "left"), ""))
        lines.append(f"{ROLE_LABELS.get(role, role)}：" + " · ".join(x for x in parts if x))
    tbl_st = {"three_line": "三线表", "grid": "网格表", "none": "无线表"}.get(
        getattr(s, "table_style", "three_line"), "三线表")
    outer = float(getattr(s, "table_line", 1.5) or 0)
    inner = float(getattr(s, "table_line_inner", 0.75) or 0)
    if getattr(s, "table_style", "three_line") == "none":
        lines.append(f"表格：{tbl_st}")
    else:
        lines.append(f"表格：{tbl_st}（外/顶底 {outer:g} 磅 · 内/栏目 {inner:g} 磅）")
    lines.append("")
    if sc.source:
        lines.append(f"来源：{sc.source}")
    if sc.updated:
        lines.append(f"保存：{sc.updated}")
    return "\n".join(lines)


class SaveSchemeDialog(QDialog):
    """把当前手动调好的整套设置存成一个新的（或覆盖已有）格式方案。"""

    def __init__(self, current: schemes.Scheme, settings, parent=None):
        super().__init__(parent)
        self.setWindowTitle("保存当前设置为格式方案")
        self.resize(500, 380)
        self.settings = settings
        self.scheme: schemes.Scheme | None = None

        lay = QVBoxLayout(self)
        lay.setContentsMargins(18, 18, 18, 14)
        lay.setSpacing(12)

        form = QFormLayout()
        form.setSpacing(10)
        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("例如：XX大学硕士论文 / XX期刊模板")
        if current is not None and not current.builtin:
            self.name_edit.setText(current.name)
            self.name_edit.selectAll()
        form.addRow("方案名称", self.name_edit)

        self.kind_cb = QComboBox()
        for k in schemes.KINDS:
            self.kind_cb.addItem(schemes.KIND_LABEL[k], k)
        idx = self.kind_cb.findData(getattr(settings, "profile", "journal"))
        self.kind_cb.setCurrentIndex(idx if idx >= 0 else 0)
        form.addRow("论文类型", self.kind_cb)

        self.note_edit = QLineEdit("手动设置")
        self.note_edit.setPlaceholderText("可选：记录适用期刊、学校或用途")
        form.addRow("备注说明", self.note_edit)
        lay.addLayout(form)

        lay.addWidget(QLabel("将保存的设置摘要："))
        preview = QPlainTextEdit()
        preview.setReadOnly(True)
        tmp = schemes.Scheme(
            name="（待保存）",
            kind=getattr(settings, "profile", "journal"),
            settings=settings.scheme_payload(),
        )
        preview.setPlainText(_scheme_summary(tmp))
        lay.addWidget(preview, 1)

        btns = QHBoxLayout()
        btns.addStretch()
        cancel = QPushButton("取消")
        cancel.clicked.connect(self.reject)
        ok = QPushButton("保存方案")
        ok.setObjectName("primary")
        ok.clicked.connect(self._save)
        btns.addWidget(cancel)
        btns.addWidget(ok)
        lay.addLayout(btns)
        self.name_edit.setFocus()

    def _save(self) -> None:
        name = self.name_edit.text().strip()
        if not name:
            QMessageBox.warning(self, "方案名称", "请先给这套格式方案起个名字。")
            self.name_edit.setFocus()
            return
        if name in schemes.BUILTIN_NAME.values():
            QMessageBox.warning(self, "名称冲突",
                                "不能与内置方案同名，请换一个名称。")
            self.name_edit.setFocus()
            return
        kind = self.kind_cb.currentData() or "journal"
        existing = schemes.find(name, kind)
        if existing is not None and not existing.builtin:
            if QMessageBox.question(
                self, "覆盖方案",
                f"「{schemes.KIND_LABEL[kind]}」里已有方案「{name}」，确定用当前设置覆盖吗？",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes,
            ) != QMessageBox.Yes:
                return
        payload = self.settings.scheme_payload()
        payload["profile"] = kind
        sc = schemes.Scheme(
            name=name,
            kind=kind,
            settings=payload,
            source=self.note_edit.text().strip() or "手动设置",
            builtin=False,
        )
        schemes.save(sc)
        self.scheme = sc
        self.accept()


class FormatSchemeDialog(QDialog):
    """格式方案管理（对应草图 1）。"""

    def __init__(self, current: schemes.Scheme, parent=None,
                 import_handler=None, save_current_handler=None):
        super().__init__(parent)
        self.setWindowTitle("格式方案")
        self.resize(820, 520)
        self.current = current
        self.selected = current
        self.import_handler = import_handler
        self.save_current_handler = save_current_handler

        lay = QVBoxLayout(self)
        lay.setContentsMargins(18, 18, 18, 14)
        lay.setSpacing(12)

        body = QHBoxLayout()
        self.tree = QTreeWidget()
        self.tree.setHeaderHidden(True)
        self.tree.setColumnCount(1)
        self.tree.setMinimumWidth(300)
        # 选中行颜色直接写进调色板：QSS 压不住非活动窗口的系统高亮（白字浅底）
        pal = self.tree.palette()
        for grp in (QPalette.ColorGroup.Active, QPalette.ColorGroup.Inactive):
            pal.setColor(grp, QPalette.ColorRole.Highlight, QColor("#e8f1ff"))
            pal.setColor(grp, QPalette.ColorRole.HighlightedText, QColor("#1d1d1f"))
        self.tree.setPalette(pal)
        self.tree.currentItemChanged.connect(self._on_select)
        body.addWidget(self.tree, 3)

        self.summary = QPlainTextEdit()
        self.summary.setReadOnly(True)
        body.addWidget(self.summary, 4)
        lay.addLayout(body, 1)

        row = QHBoxLayout()
        self.apply_btn = QPushButton("应用此方案")
        self.apply_btn.setObjectName("primary")
        self.apply_btn.clicked.connect(self._apply)
        self.save_cur_btn = QPushButton("保存当前设置…")
        self.save_cur_btn.setToolTip("把软件里现在调好的所有格式存成一套方案")
        self.save_cur_btn.clicked.connect(self._save_current)
        self.save_cur_btn.setVisible(self.save_current_handler is not None)
        self.import_btn = QPushButton("从模板导入…")
        self.import_btn.clicked.connect(self._import)
        self.rename_btn = QPushButton("重命名")
        self.rename_btn.clicked.connect(self._rename)
        self.move_btn = QPushButton("移到学位论文")
        self.move_btn.setToolTip("在「期刊论文」和「学位论文」分类之间移动此方案")
        self.move_btn.clicked.connect(self._move_kind)
        self.delete_btn = QPushButton("删除")
        self.delete_btn.clicked.connect(self._delete)
        self.export_btn = QPushButton("导出…")
        self.export_btn.clicked.connect(self._export)
        for b in (self.apply_btn, self.save_cur_btn, self.import_btn,
                  self.rename_btn, self.move_btn, self.delete_btn,
                  self.export_btn):
            row.addWidget(b)
        row.addStretch()
        cancel = QPushButton("取消")
        cancel.clicked.connect(self.reject)
        row.addWidget(cancel)
        lay.addLayout(row)

        self._fill()

    # ---------------------------------------------------------------- #
    def _fill(self) -> None:
        self.tree.clear()
        self.groups: dict[str, QTreeWidgetItem] = {}
        for kind in schemes.KINDS:
            top = QTreeWidgetItem([schemes.KIND_LABEL[kind]])
            top.setFlags(top.flags() & ~Qt.ItemIsSelectable)
            f = top.font(0)
            f.setBold(True)
            top.setFont(0, f)
            self.tree.addTopLevelItem(top)
            self.groups[kind] = top
            items = [schemes.builtin(kind)]
            items += [s for s in schemes.load_user_schemes() if s.kind == kind]
            for sc in items:
                child = QTreeWidgetItem([sc.label])
                child.setData(0, Qt.UserRole, sc.name + "|" + sc.kind)
                top.addChild(child)
                if sc.name == self.current.name and sc.kind == self.current.kind:
                    self.tree.setCurrentItem(child)
            top.setExpanded(True)
        if self.tree.currentItem() is None:
            first = self.tree.topLevelItem(0)
            if first and first.childCount():
                self.tree.setCurrentItem(first.child(0))

    def _on_select(self, item: QTreeWidgetItem) -> None:
        if item is None or item.parent() is None:
            return
        # 选中行直接画死颜色：Qt 在非活动窗口里会用系统高亮色（白字浅底），
        # 光靠 QSS 的 ::item:selected 压不住
        for i in range(self.tree.topLevelItemCount()):
            top = self.tree.topLevelItem(i)
            for j in range(top.childCount()):
                it = top.child(j)
                it.setBackground(0, QBrush(QColor("#e8f1ff") if it is item
                                           else QColor("#ffffff")))
                it.setForeground(0, QBrush(QColor("#1d1d1f")))
        key = item.data(0, Qt.UserRole) or ""
        name, _, kind = key.partition("|")
        sc = schemes.find(name, kind)
        if sc is None:
            return
        self.selected = sc
        self.summary.setPlainText(_scheme_summary(sc))
        builtin = sc.builtin
        self.rename_btn.setEnabled(not builtin)
        self.move_btn.setEnabled(not builtin)
        other = "thesis" if sc.kind == "journal" else "journal"
        self.move_btn.setText(f"移到{schemes.KIND_LABEL[other]}")
        self.delete_btn.setEnabled(not builtin)

    def _apply(self) -> None:
        self.current = self.selected
        self.accept()

    def _save_current(self) -> None:
        if self.save_current_handler is None:
            return
        sc = self.save_current_handler()
        if sc is None:
            return
        self.current = sc
        self._fill()

    def _import(self) -> None:
        if self.import_handler is None:
            return
        sc = self.import_handler()
        if sc is None:
            return
        self.current = sc
        self._fill()
        QMessageBox.information(self, "已导入", f"方案「{sc.name}」已保存。")

    def _rename(self) -> None:
        sc = self.selected
        if sc.builtin:
            return
        name, ok = QInputDialog.getText(self, "重命名方案", "新名字：",
                                        QLineEdit.Normal, sc.name)
        if not ok or not name.strip() or name.strip() == sc.name:
            return
        if schemes.rename(sc, name.strip()):
            self.current = sc
            self._fill()
        else:
            QMessageBox.warning(self, "重命名失败", "这个名字已被占用。")

    def _move_kind(self) -> None:
        sc = self.selected
        if sc.builtin:
            return
        other = "thesis" if sc.kind == "journal" else "journal"
        if schemes.move_kind(sc, other):
            self.current = sc
            self._fill()
        else:
            QMessageBox.warning(
                self, "移动失败",
                f"「{schemes.KIND_LABEL[other]}」里已有同名方案「{sc.name}」。")

    def _delete(self) -> None:
        sc = self.selected
        if sc.builtin:
            return
        if QMessageBox.question(self, "删除方案",
                                f"确定删除「{sc.name}」？此操作不可撤销。",
                                QMessageBox.Yes | QMessageBox.No,
                                QMessageBox.No) != QMessageBox.Yes:
            return
        schemes.delete(sc)
        self._fill()

    def _export(self) -> None:
        sc = self.selected
        path, _ = QFileDialog.getSaveFileName(
            self, "导出方案", f"{sc.name}.json", "格式方案 (*.json)")
        if not path:
            return
        schemes.export(sc, path)
        QMessageBox.information(self, "已导出", f"已导出到：\n{path}")


class DetectDialog(QDialog):
    """格式识别（对应草图 2）。"""

    def __init__(self, doc, analysis, base_settings=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("格式识别")
        self.resize(920, 640)
        self.doc = doc
        self.analysis = analysis
        self.base = base_settings
        self.report: dict = {}
        self.settings = None
        self.scheme = None
        self.template_path = ""

        lay = QVBoxLayout(self)
        lay.setContentsMargins(18, 18, 18, 14)
        lay.setSpacing(10)

        # ---- 来源 ----
        # 两组单选按钮必须装进不同的 QButtonGroup：否则同属 DetectDialog，
        # Qt 会把 4 个单选框合并成一个互斥组，选「学位」会把「来源」取消、
        # 点「从模板识别」又会把「学位」取消掉（导致永远存进期刊论文）。
        src = QHBoxLayout()
        self.cur_radio = QRadioButton("从当前文档识别")
        self.cur_radio.setChecked(True)
        self.tpl_radio = QRadioButton("从模板文件识别")
        self._src_group = QButtonGroup(self)
        self._src_group.addButton(self.cur_radio)
        self._src_group.addButton(self.tpl_radio)
        src.addWidget(self.cur_radio)
        self.tpl_path = QLineEdit()
        self.tpl_path.setReadOnly(True)
        self.tpl_path.setPlaceholderText("尚未选择文件")
        src.addWidget(self.tpl_radio)
        src.addWidget(self.tpl_path, 1)
        self.pick_btn = QPushButton("选择文件…")
        self.pick_btn.clicked.connect(self._pick)
        src.addWidget(self.pick_btn)
        self.detect_btn = QPushButton("识别")
        self.detect_btn.setObjectName("primary")
        self.detect_btn.clicked.connect(self._detect)
        src.addWidget(self.detect_btn)
        lay.addLayout(src)

        # ---- 核对表 ----
        lay.addWidget(QLabel("识别结果（可修改后再保存）"))
        self.table = QTableWidget(0, 8)
        self.table.setHorizontalHeaderLabels(
            ["角色", "中文字体", "西文字体", "字号", "加粗", "行距", "对齐", "状态"])
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(38)
        # 按内容自适应：列宽不够时 Qt 会把组合框裁到看不见字
        head = self.table.horizontalHeader()
        for i in range(7):
            head.setSectionResizeMode(i, QHeaderView.ResizeToContents)
        head.setStretchLastSection(True)
        for i, w in enumerate((104, 92, 150, 78, 52, 100, 92)):
            self.table.setColumnWidth(i, w)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        lay.addWidget(self.table, 1)

        # ---- 保存为方案 ----
        save = QHBoxLayout()
        save.addWidget(QLabel("保存为方案"))
        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("方案名，例如：A 会议格式")
        save.addWidget(self.name_edit, 1)
        save.addWidget(QLabel("类型"))
        self.kind_journal = QRadioButton("期刊")
        self.kind_thesis = QRadioButton("学位")
        self._kind_group = QButtonGroup(self)
        self._kind_group.addButton(self.kind_journal)
        self._kind_group.addButton(self.kind_thesis)
        if getattr(base_settings, "profile", "journal") == "thesis":
            self.kind_thesis.setChecked(True)
        else:
            self.kind_journal.setChecked(True)
        save.addWidget(self.kind_journal)
        save.addWidget(self.kind_thesis)
        lay.addLayout(save)

        box = QDialogButtonBox()
        self.save_scheme_btn = QPushButton("保存方案")
        self.save_scheme_btn.setObjectName("primary")
        self.save_scheme_btn.clicked.connect(self._save_scheme)
        self.apply_btn = QPushButton("直接覆盖当前设置")
        self.apply_btn.clicked.connect(self._apply_now)
        cancel = QPushButton("取消")
        cancel.clicked.connect(self.reject)
        box.addButton(self.save_scheme_btn, QDialogButtonBox.ActionRole)
        box.addButton(self.apply_btn, QDialogButtonBox.ActionRole)
        box.addButton(cancel, QDialogButtonBox.RejectRole)
        lay.addWidget(box)

        self._cells: dict[str, list] = {}

    # ---------------------------------------------------------------- #
    def _pick(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "选择模板文件", "", "Word 文档 (*.docx *.doc)")
        if path:
            self.template_path = path
            self.tpl_path.setText(path)
            self.tpl_radio.setChecked(True)

    def _source(self):
        """返回 (doc, analysis, 来源名)。"""
        if self.tpl_radio.isChecked() and self.template_path:
            from pathlib import Path
            from docx import Document

            from ..classify import analyze
            from ..features import ensure_docx

            src = Path(self.template_path)
            docx_path = ensure_docx(src)
            doc = Document(str(docx_path))
            analysis = analyze(doc, source=src.name)
            return doc, analysis, analysis.source
        return self.doc, self.analysis, getattr(self.analysis, "source", "")

    def _detect(self) -> None:
        try:
            doc, analysis, source = self._source()
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "打不开模板", str(exc))
            return
        report: dict = {}
        try:
            recognize_from_doc(doc, analysis, report=report)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "识别失败", str(exc))
            return
        self.report = report
        self._fill_table()

    def _fill_table(self) -> None:
        roles = [r for r in DETECT_ROLES if r in self.report.get("roles", {})]
        self.table.setRowCount(len(roles))
        self._cells = {}
        for row, role in enumerate(roles):
            attrs = self.report["roles"][role]
            base = (self.base.roles.get(role) if self.base else {}) or {}
            self.table.setItem(row, 0, QTableWidgetItem(ROLE_LABELS.get(role, role)))

            cb = QComboBox()
            cb.addItems(FONTS)
            cb.setEditable(True)
            cb.setCurrentText(attrs.get("font") or base.get("font") or "宋体")
            cb.lineEdit().setCursorPosition(0)
            self.table.setCellWidget(row, 1, cb)

            cl = QComboBox()
            cl.addItems(LATIN_FONTS)
            cl.setEditable(True)
            cl.setCurrentText(attrs.get("ascii") or base.get("ascii") or "Times New Roman")
            cl.lineEdit().setCursorPosition(0)
            self.table.setCellWidget(row, 2, cl)

            sz = QDoubleSpinBox()
            sz.setRange(4, 72)
            sz.setDecimals(1)
            sz.setSingleStep(0.5)
            sz.setValue(float(attrs.get("size") or base.get("size") or 12))
            self.table.setCellWidget(row, 3, sz)

            bold = MarkCheck("")
            bold.setChecked(bool(attrs.get("bold") if attrs.get("bold") is not None
                                 else base.get("bold")))
            self.table.setCellWidget(row, 4, bold)

            lc = QComboBox()
            for key, text in LINE_CHOICES:
                lc.addItem(text, key)
            line = attrs.get("line")
            rule = line[0] if line else (base.get("line_rule") or "exact")
            idx = lc.findData(rule)
            lc.setCurrentIndex(idx if idx >= 0 else 0)
            self.table.setCellWidget(row, 5, lc)

            al = QComboBox()
            for key, text in ALIGNS:
                al.addItem(text, key)
            idx = al.findData(attrs.get("align") or base.get("align") or "left")
            al.setCurrentIndex(idx if idx >= 0 else 0)
            self.table.setCellWidget(row, 6, al)

            detected = any(attrs.get(k) is not None for k in
                           ("size", "font", "ascii", "bold", "align")) \
                or attrs.get("line") is not None
            st = QTableWidgetItem("未识别" if not detected else "实测")
            if not detected:
                st.setForeground(Qt.GlobalColor.gray)
                # 未识别的行：值来自预设默认，灰掉以免被当成识别结果
                for w in (cb, cl, sz, bold, lc, al):
                    w.setStyleSheet("color: #8a8f98;")
            self.table.setItem(row, 7, st)

            self._cells[role] = [cb, cl, sz, bold, lc, al]

    def _collect(self) -> dict:
        """把（可能被改过的）核对表读回报告。"""
        roles = self.report.setdefault("roles", {})
        for role, cells in self._cells.items():
            cb, cl, sz, bold, lc, al = cells
            attrs = roles.setdefault(role, {})
            attrs["font"] = cb.currentText().strip()
            attrs["ascii"] = cl.currentText().strip()
            attrs["size"] = float(sz.value())
            attrs["bold"] = bool(bold.isChecked())
            rule = lc.currentData() or "exact"
            prev = attrs.get("line")
            prev_val = float(prev[1]) if prev and prev[1] is not None else None
            size = float(attrs.get("size") or 12)
            if rule == "multiple":
                attrs["line"] = ("multiple", 1.5)
            elif rule == "multiple2":
                attrs["line"] = ("multiple2", 2.0)
            elif rule == "single":
                attrs["line"] = ("single", None)
            elif rule == "minimum":
                attrs["line"] = ("minimum", prev_val or round(size * 1.2, 1))
            else:
                attrs["line"] = ("exact", prev_val or round(size * 1.4, 1))
            attrs["align"] = al.currentData() or "left"
        return self.report

    def _save_scheme(self) -> None:
        name = self.name_edit.text().strip()
        if not name:
            QMessageBox.warning(self, "缺少方案名", "请先填一个方案名。")
            return
        if not self._cells:
            QMessageBox.warning(self, "还没有识别结果", "请先点「识别」。")
            return
        kind = "thesis" if self.kind_thesis.isChecked() else "journal"
        report = self._collect()
        s = settings_from_report(self.base, report) if self.base else \
            recognize_from_doc(*self._source()[:2])
        payload = s.scheme_payload()
        payload["profile"] = kind
        sc = schemes.Scheme(name=name, kind=kind, settings=payload,
                            source=self.report.get("source", ""))
        try:
            schemes.save(sc)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "保存失败", str(exc))
            return
        self.scheme = sc
        self.accept()

    def _apply_now(self) -> None:
        if not self._cells:
            QMessageBox.warning(self, "还没有识别结果", "请先点「识别」。")
            return
        report = self._collect()
        if self.base is None:
            QMessageBox.warning(self, "无法覆盖", "当前没有打开的文档设置。")
            return
        self.settings = settings_from_report(self.base, report)
        self.accept()
