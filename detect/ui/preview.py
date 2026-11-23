"""应用前预览对话框：列出将要插入的编号与风险提示。"""
from __future__ import annotations

from PySide6.QtWidgets import (
    QDialog,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QHeaderView,
)

from ..numbering import summarize


class PreviewDialog(QDialog):
    def __init__(self, analysis, plans: list, parent=None):
        super().__init__(parent)
        self.plans = plans
        self.setWindowTitle("应用预览")
        self.resize(780, 580)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 24, 24, 24)
        lay.setSpacing(14)

        head = QLabel(f"将要插入：{summarize(plans)}")
        head.setObjectName("card-title")
        lay.addWidget(head)

        doubtful = sum(
            1 for d in analysis.detections
            if (d.role == "heading" or d.role.startswith("heading"))
            and not d.number and not d.user_confirmed
        )
        warn = QLabel(
            f"⚠ 其中有 {doubtful} 个标题的层级是推断出来的（无编号锚点）。"
            f"层级若判错，后续编号会整体错位——建议先核对这些标题。"
            if doubtful else "✓ 所有标题层级都由编号锚点确定。"
        )
        warn.setObjectName("warn" if doubtful else "ok")
        warn.setWordWrap(True)
        lay.addWidget(warn)

        table = QTableWidget()
        table.setColumnCount(4)
        table.setHorizontalHeaderLabels(["#", "段落", "将插入", "原编号"])
        table.verticalHeader().setVisible(False)
        table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        for col in (0, 2, 3):
            table.horizontalHeader().setSectionResizeMode(col, QHeaderView.ResizeToContents)

        shown = plans[:200]
        table.setRowCount(len(shown))
        for row, p in enumerate(shown):
            table.setItem(row, 0, QTableWidgetItem(str(p.index)))
            table.setItem(row, 1, QTableWidgetItem(analysis.features[p.index].text[:60]))
            table.setItem(row, 2, QTableWidgetItem(p.text.strip()))
            table.setItem(row, 3, QTableWidgetItem(p.old or ""))
        lay.addWidget(table, 1)

        if len(plans) > 200:
            more = QLabel(f"… 其余 {len(plans) - 200} 条未列出")
            more.setObjectName("card-note")
            lay.addWidget(more)

        btns = QHBoxLayout()
        btns.addStretch()
        cancel = QPushButton("取消")
        cancel.clicked.connect(self.reject)
        ok = QPushButton("确认应用")
        ok.setObjectName("primary")
        ok.clicked.connect(self.accept)
        btns.addWidget(cancel)
        btns.addWidget(ok)
        lay.addLayout(btns)
