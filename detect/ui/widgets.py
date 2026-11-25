"""自绘复选框。

为什么不用 QCheckBox：Qt 的 QSS 在 `::indicator` 上不支持 `image` /
`background-image`（实测蓝底能出来、对勾画不出来），
所以对勾只能自己画。尺寸、圆角、勾的粗细都在这里调。
"""
from __future__ import annotations

from PySide6.QtCore import QPointF, QRectF, Qt, Signal
from PySide6.QtGui import QColor, QFont, QPainter, QPen
from PySide6.QtWidgets import QAbstractButton, QWidget

BOX = 18
FONT = 13
ON = QColor("#007aff")
OFF_BORDER = QColor("#c7c7cc")
TEXT = QColor("#1d1d1f")
TEXT_OFF = QColor("#b0b0b6")
PILL_OFF = QColor("#e4e4ea")
PILL_ON = QColor("#d3d3da")


class PillButton(QAbstractButton):
    """自绘的胶囊开关/切换钮。

    为什么自绘：Qt 原生对 checkable 按钮的 checked 状态会按系统主题上色，
    在同一个窗口里出现"同样状态不同颜色"的情况（实测右栏的按钮变蓝），
    而且容易和主操作按钮（蓝底白字）撞级别。这里颜色完全由 QSS 之外控制。
    """

    def __init__(self, text: str = "", parent=None):
        super().__init__(parent)
        self._text = text
        self.setCheckable(True)
        self.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self._syncAccessible()
        self.adjustSize()

    def setText(self, text: str) -> None:
        self._text = text
        self._syncAccessible()
        self.adjustSize()
        self.update()

    def text(self) -> str:
        return self._text

    def setChecked(self, value: bool) -> None:
        super().setChecked(bool(value))
        self._syncAccessible()

    def _syncAccessible(self) -> None:
        # 自绘控件不会自动把文字暴露给 UIAutomation/屏幕阅读器，手动给
        state = "已开启" if self.isChecked() else "未开启"
        self.setAccessibleName(f"{self._text}（{state}）")
        self.setAccessibleDescription(self.toolTip() or self._text)

    def sizeHint(self):
        from PySide6.QtCore import QSize
        from PySide6.QtGui import QFontMetrics
        m = QFontMetrics(self.font())
        return QSize(m.horizontalAdvance(self._text) + 22, 22)

    def paintEvent(self, _e) -> None:
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        r = QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5)
        if self.isChecked():
            p.setPen(QPen(QColor("#c7c7cc"), 1))
            p.setBrush(PILL_ON)
        else:
            p.setPen(QPen(OFF_BORDER, 1))
            p.setBrush(PILL_OFF)
        p.drawRoundedRect(r, r.height() / 2, r.height() / 2)
        f = QFont()
        f.setPixelSize(12.5)
        f.setWeight(QFont.Weight.Bold if self.isChecked() else QFont.Weight.Normal)
        p.setFont(f)
        p.setPen(TEXT if self.isEnabled() else TEXT_OFF)
        p.drawText(self.rect(), Qt.AlignCenter, self._text)
        p.end()

    def mousePressEvent(self, e) -> None:
        if not self.isEnabled():
            return
        self.setChecked(not self.isChecked())
        # 必须自己发 clicked：QAbstractButton.clicked 是基类在
        # mouseReleaseEvent 里发出的，重写 mousePressEvent 后不会再有。
        # 漏了这一句会导致「同步 / 适应窗口 / 适应整页」点了完全没反应。
        self.clicked.emit()


class MarkCheck(QWidget):
    """带文字的复选框：圆角方框 + 白色对勾。"""

    toggled = Signal(bool)

    def __init__(self, text: str = "", parent=None):
        super().__init__(parent)
        self._on = False
        self._text = text
        self._enabled = True
        self.setFixedHeight(26)
        self.setCursor(Qt.CursorShape.PointingHandCursor)

    # ---------------------------------------------------------------- #
    def setText(self, text: str) -> None:
        self._text = text
        self.update()

    def text(self) -> str:
        return self._text

    def setChecked(self, value: bool) -> None:
        if self._on == bool(value):
            return
        self._on = bool(value)
        self.update()
        self.toggled.emit(self._on)

    def isChecked(self) -> bool:
        return self._on

    def setEnabled(self, enabled: bool) -> None:
        self._enabled = bool(enabled)
        self.update()
        super().setEnabled(enabled)

    def sizeHint(self):
        from PySide6.QtCore import QSize
        from PySide6.QtGui import QFontMetrics
        f = QFont()
        f.setPixelSize(FONT)
        w = BOX + (9 + QFontMetrics(f).horizontalAdvance(self._text) if self._text else 0) + 4
        return QSize(max(BOX + 4, w), 26)

    def minimumSizeHint(self):
        return self.sizeHint()

    # ---------------------------------------------------------------- #
    def paintEvent(self, _e) -> None:
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        box = QRectF(0, (self.height() - BOX) / 2, BOX, BOX)
        if self._on:
            p.setPen(QPen(ON, 1.5))
            p.setBrush(ON)
        else:
            p.setPen(QPen(OFF_BORDER, 1.5))
            p.setBrush(QColor("#ffffff"))
        p.drawRoundedRect(box, 5, 5)
        if self._on:
            p.setPen(QPen(QColor("#ffffff"), 2.1, Qt.PenStyle.SolidLine,
                           Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin))
            p.drawPolyline([
                QPointF(box.left() + 4.2, box.center().y() + 0.4),
                QPointF(box.left() + 7.6, box.bottom() - 3.6),
                QPointF(box.right() - 3.6, box.top() + 3.2),
            ])
        f = QFont()
        f.setPixelSize(FONT)
        p.setFont(f)
        p.setPen(TEXT if self._enabled else TEXT_OFF)
        p.drawText(QRectF(box.right() + 9, 0, self.width() - box.right() - 9,
                          self.height()),
                   Qt.AlignVCenter | Qt.AlignLeft, self._text)
        p.end()

    def mousePressEvent(self, e) -> None:
        if not self._enabled:
            return
        self.setChecked(not self._on)
