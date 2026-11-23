"""界面样式：简洁现代通用风（mac 设计语言的落地）。

设计令牌（改这里就能整体换肤）：
    背景  #f5f5f7      卡片  #ffffff      边框  #e5e5ea
    主文字 #1d1d1f     次文字 #86868b
    强调  #007aff      成功  #34c759      警告  #ff9f0a      错误  #ff3b30

四条硬规矩：
    1. 排版分四级（17/600、14/600、13/400、12 灰），主次分明
    2. 行高 44px、卡片内边距 20px、卡片间距 20px、页面边距 24px
    3. 分隔用 1px 浅灰线，不用重边框；分组靠卡片
    4. 强调色只有一个蓝；警告色只给徽标和小提示条，不大面积铺色
"""
from __future__ import annotations

import base64
import io


def _check_mark(size: int = 34) -> str:
    """画一个白色对勾，base64 PNG。

    用 PNG 而不是 SVG data URL：实测 Qt 的 QSS 对 data:image/svg+xml 支持不稳定，
    勾画不出来（蓝底是出来了，勾没有）。PNG 一定画得出来。
    """
    from PIL import Image, ImageDraw

    im = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    w = max(2, round(size / 11))
    d.line([(size * 0.17, size * 0.53), (size * 0.42, size * 0.78),
            (size * 0.84, size * 0.25)],
           fill=(255, 255, 255, 255), width=w, joint="curve")
    buf = io.BytesIO()
    im.save(buf, "PNG")
    return base64.b64encode(buf.getvalue()).decode("ascii")


_CHECK_PNG = _check_mark()


LIGHT = """
* {
    font-family: "Microsoft YaHei", "PingFang SC", "Segoe UI", sans-serif;
    font-size: 13px;
    color: #1d1d1f;
}

QMainWindow, QWidget#root { background: #f5f5f7; }

/* ============ 顶部栏 ============ */
QWidget#topbar {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                stop:0 #ffffff, stop:1 #f6f6f9);
    border-bottom: 1px solid #e5e5ea;
}
QWidget#title-block { background: transparent; }
QLabel#app-title { font-size: 17px; font-weight: 600; }
QLabel#app-sub { color: #86868b; font-size: 11.5px; }
QLabel#about-version { color: #007aff; font-size: 12px; font-weight: 600; }

/* ============ 顶栏：论文形态切换 ============ */
QDialog#para-pop {
    background: #ffffff;
    border: 1px solid #d2d2d7;
    border-radius: 10px;
}
QDialog#para-pop QComboBox { padding: 3px 8px; }
QDialog#para-pop QPushButton { padding: 4px 8px; font-size: 12px; }
QLabel#para-pop-title { color: #1d1d1f; font-weight: 600; font-size: 12px; }
QWidget#seg {
    background: #ececf0;
    border-radius: 8px;
    padding: 2px;
}
QPushButton#seg-btn {
    background: transparent;
    border: none;
    border-radius: 6px;
    padding: 4px 10px;
    color: #86868b;
    font-size: 12px;
    font-weight: 600;
}
QPushButton#seg-btn:hover { color: #1d1d1f; }
QPushButton#seg-btn:checked {
    background: #ffffff;
    color: #1d1d1f;
}
QPushButton#seg-btn:checked:hover { color: #1d1d1f; }

/* 同步：中性灰开关。主色只留给「一键排版」那个级别的动作 */
QPushButton#sync { background: #e8e8ed; border: 1px solid transparent; color: #555557; }
QPushButton#sync:hover { background: #dedee4; }
QPushButton#sync:checked {
    background: #dcdce2;
    border: 1px solid #c7c7cc;
    color: #1d1d1f;
    font-weight: 600;
}

/* ============ 预览面板标题行 / 缩放 ============ */
QPushButton#zoom-btn { padding: 1px 7px; font-size: 14px; line-height: 1.2; }
QPushButton#fit-btn { padding: 3px 9px; font-size: 12px; color: #86868b; }
QPushButton#fit-btn:hover { color: #1d1d1f; }
QPushButton#fit-btn:checked {
    background: #e4e4ea;
    border: 1px solid #c7c7cc;
    color: #1d1d1f;
    font-weight: 600;
}
QLabel#zoom-pct {
    color: #86868b;
    font-size: 12px;
    min-width: 56px;
    text-align: center;
}

/* ============ 左侧导航 ============ */
QWidget#nav-wrap { background: #f5f5f7; border-right: 1px solid #e5e5ea; }
QLabel#nav-title { color: #86868b; font-size: 12px; padding: 20px 20px 6px 20px; }
QListWidget#nav {
    background: transparent;
    border: none;
    outline: none;
    padding: 0 12px;
}
QListWidget#nav::item {
    padding: 11px 14px;
    border-radius: 8px;
    margin: 2px 0;
    color: #555557;
}
QListWidget#nav::item:selected {
    background: #e8e8ed;
    color: #1d1d1f;
    font-weight: 600;
}
QListWidget#nav::item:hover:!selected { background: #eeeeF2; }

/* ============ 右侧内容 ============ */
QWidget#content { background: #f5f5f7; }

/* ============ 卡片 ============ */
QFrame.card {
    background: #ffffff;
    border: 1px solid #e5e5ea;
    border-radius: 12px;
}
QLabel.card-title { font-size: 14px; font-weight: 600; }
QLabel.card-note { color: #86868b; font-size: 12px; }

/* ============ 统计卡片 ============ */
QFrame.stat { background: #ffffff; border: 1px solid #e5e5ea; border-radius: 12px; }
QLabel.stat-num { font-size: 28px; font-weight: 600; color: #1d1d1f; }
QLabel.stat-label { font-size: 12px; color: #86868b; }

/* ============ 分层结果卡片 ============ */
QFrame#verdict {
    background: #f0f6ff;
    border: 1px solid #b8d8ff;
    border-radius: 12px;
}
QLabel#verdict-num { font-size: 28px; font-weight: 600; color: #007aff; }
QLabel#verdict-main { font-size: 14px; font-weight: 600; }
QLabel#verdict-note { color: #555557; font-size: 12px; }

/* ============ 按钮 ============ */
QPushButton {
    border-radius: 8px;
    padding: 8px 18px;
    border: 1px solid transparent;
    background: #e8e8ed;
    color: #1d1d1f;
}
QPushButton:hover { background: #dedee4; }
QPushButton:pressed { background: #d3d3d9; }
QPushButton:disabled { color: #b0b0b6; background: #f0f0f3; }

QPushButton#primary { background: #007aff; color: #ffffff; font-weight: 600; }
QPushButton#primary:hover { background: #006fe6; }
QPushButton#primary:pressed { background: #0064d2; }
QPushButton#primary:disabled { background: #a9cfff; color: #ffffff; }

QPushButton#ghost { background: transparent; border: 1px solid #d2d2d7; color: #1d1d1f; }
QPushButton#ghost:hover { background: #f5f5f7; }

/* 「关于」里的帮助文档：画成一行链接文字，但用 QPushButton 才能
   被 UIAutomation/键盘正常触发（QLabel 的 linkActivated 只认鼠标）。 */
QPushButton#help-link {
    background: transparent;
    border: none;
    color: #007aff;
    padding: 0;
    font-size: 13px;
    text-align: left;
}
QPushButton#help-link:hover { text-decoration: underline; }

/* ============ 大纲树 ============ */
QTreeWidget {
    background: #ffffff;
    border: 1px solid #e5e5ea;
    border-radius: 12px;
    outline: none;
    padding: 8px;
}
QTreeWidget::item {
    height: 44px;
    padding: 0 8px;
    border-radius: 6px;
    border-bottom: 1px solid #f5f5f7;
}
QTreeWidget::item:selected { background: #e8f1ff; }
QTreeWidget::item:hover { background: #f7f9fc; }
QTreeWidget::branch { image: none; }
QHeaderView::section {
    background: transparent;
    border: none;
    border-bottom: 1px solid #e5e5ea;
    padding: 10px 12px;
    font-weight: 600;
    color: #86868b;
    font-size: 12px;
}

/* ============ 徽标 ============ */
QLabel#badge-ok { color: #34c759; font-weight: 600; }
QLabel#badge-warn { color: #ff9f0a; font-weight: 600; }
QLabel#badge-none { color: #c7c7cc; font-weight: 600; }
QLabel#badge-num { color: #007aff; font-weight: 600; }
QLabel#badge-none-num { color: #86868b; font-weight: 600; }

/* ============ 输入 ============ */
QLineEdit, QComboBox {
    border: 1px solid #d2d2d7;
    border-radius: 8px;
    padding: 7px 12px;
    background: #ffffff;
}
QLineEdit:focus, QComboBox:focus { border: 1px solid #007aff; }
QComboBox::drop-down { border: none; width: 22px; }

/* ============ 参数表 ============ */
QTableWidget {
    background: #ffffff;
    border: 1px solid #e5e5ea;
    border-radius: 12px;
    gridline-color: #f5f5f7;
    selection-background-color: #e8f1ff;
    outline: none;
}
QTableWidget::item { padding: 8px; border-bottom: 1px solid #f5f5f7; }
QHeaderView#param-head::section {
    background: #fafafc;
    border: none;
    border-bottom: 1px solid #e5e5ea;
    padding: 10px 12px;
    font-weight: 600;
    color: #555557;
}

/* ============ Ribbon 功能区 ============ */
QWidget#ribbon { background: #ffffff; border-bottom: 1px solid #e5e5ea; }
QFrame#ribbon-group {
    background: #fafafc;
    border: 1px solid #e5e5ea;
    border-radius: 10px;
}
QLabel#ribbon-group-title {
    color: #86868b;
    font-size: 11px;
    text-align: center;
    padding-top: 2px;
}
QPushButton#ribbon-btn {
    background: #ffffff;
    border: 1px solid #e5e5ea;
    padding: 4px 12px;
    text-align: center;
    border-radius: 6px;
}
QPushButton#ribbon-btn:hover { background: #f0f6ff; border: 1px solid #b8d8ff; }
QPushButton#ribbon-btn:pressed { background: #e0edff; }
/* 期刊形态下"目录"这种看得见但不可用的卡片：画成灰色，但仍然可指向、
   能弹 tooltip（所以不能用 :disabled，那样鼠标事件根本不投递） */
QPushButton#ribbon-btn[locked="true"] {
    color: #b0b0b6;
    background: #f0f0f3;
    border: 1px solid #e5e5ea;
}
QPushButton#ribbon-btn[locked="true"]:hover { background: #f0f0f3; }

/* ============ 双栏预览 ============ */
QWidget#pane-head-row { background: #f5f5f7; }
QLabel#pane-head {
    background: transparent;
    border-bottom: 1px solid #e5e5ea;
    padding: 0;
    font-weight: 600;
}
/* 跳至指定页 */
QLabel#jump-label { color: #86868b; font-size: 12px; }
QSpinBox#page-jump, QWidget#page-jump QSpinBox {
    background: #ffffff;
    border: 1px solid #e5e5ea;
    border-radius: 6px;
    padding: 1px 2px;
    font-size: 12px;
    selection-background-color: #007aff;
}
QWidget#page-jump QSpinBox:focus { border: 1px solid #007aff; }
QWidget#page-jump QSpinBox::up-button,
QWidget#page-jump QSpinBox::down-button {
    width: 14px;
    background: #f0f0f3;
    border-left: 1px solid #e5e5ea;
}
QWidget#page-jump QSpinBox::up-button:hover,
QWidget#page-jump QSpinBox::down-button:hover { background: #e4e4ea; }
QTextBrowser#preview { background: #ececf0; border: none; }
QTextBrowser#preview > QWidget { background: transparent; }

/* ============ 底部任务栏 ============ */
QWidget#taskbar {
    background: #ffffff;
    border-top: 1px solid #e5e5ea;
}
QCheckBox { spacing: 8px; padding: 1px 0; }
QCheckBox::indicator {
    width: 17px; height: 17px;
    border: 1.5px solid #c7c7cc;
    border-radius: 5px;
    background: #ffffff;
}
QCheckBox::indicator:hover { border-color: #007aff; }
QCheckBox::indicator:checked {
    background: #007aff;
    border: 1.5px solid #007aff;
    image: url("data:image/png;base64,__CHECK_PNG__");
}
QCheckBox::indicator:indeterminate {
    background: #007aff;
    border: 1.5px solid #007aff;
}
QCheckBox::indicator:disabled {
    background: #f0f0f3;
    border: 1.5px solid #e5e5ea;
}
QCheckBox::indicator:checked:disabled {
    background: #a9cfff;
    border: 1.5px solid #a9cfff;
    image: none;
}

/* ============ 参数面板行 ============ */
QFrame.card QComboBox { padding: 5px 10px; }

/* ============ 状态栏 ============ */
QStatusBar {
    background: #f5f5f7;
    border-top: 1px solid #e5e5ea;
    color: #86868b;
    font-size: 12px;
}

/* ============ 滚动条 ============ */
QScrollBar:vertical { background: transparent; width: 10px; margin: 0; }
QScrollBar::handle:vertical {
    background: #c7c7cc; border-radius: 5px; min-height: 30px;
}
QScrollBar::add-line, QScrollBar::sub-line { height: 0; }
"""


def apply(app) -> None:
    app.setStyleSheet(LIGHT.replace("__CHECK_PNG__", _CHECK_PNG))
