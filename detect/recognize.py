"""从原稿识别格式：读文档里各角色的实际生效格式，填进设置。

「从原稿识别」按钮的背后就是这个——不是让用户手填，
而是先把原稿的格式读出来当起点。
"""
from __future__ import annotations

from collections import Counter

from docx.shared import Emu

from .classify import effective_role
from .ui.settings import Settings


def _mode(values):
    vals = [v for v in values if v is not None]
    return Counter(vals).most_common(1)[0][0] if vals else None


def _mm(length) -> float:
    if length is None:
        return 0.0
    return round(length.cm, 2)


def recognize_from_doc(doc, analysis) -> Settings:
    s = Settings()
    feats = analysis.features

    # ---- 页面：读第一个节的实际页边距 ----
    if doc.sections:
        sec = doc.sections[0]
        try:
            s.page.update({
                "top": _mm(sec.top_margin),
                "bottom": _mm(sec.bottom_margin),
                "left": _mm(sec.left_margin),
                "right": _mm(sec.right_margin),
                "header": _mm(sec.header_distance),
                "footer": _mm(sec.footer_distance),
            })
        except Exception:
            pass
        # 页眉文字
        try:
            if sec.header.paragraphs and sec.header.paragraphs[0].text.strip():
                s.header_text = sec.header.paragraphs[0].text.strip()
        except Exception:
            pass

    # ---- 角色格式：按角色聚类，取众数 ----
    buckets: dict[str, list] = {}
    for d in analysis.detections:
        role = effective_role(d)
        buckets.setdefault(role, []).append(feats[d.index])

    for role, items in buckets.items():
        if role not in s.roles:
            continue
        size = _mode([f.size_pt for f in items])
        ea = _mode([f.east_asia for f in items if f.east_asia])
        ascii_f = _mode([f.ascii_font for f in items if f.ascii_font])
        bold = _mode([bool(f.bold) for f in items])
        align = _mode([f.align for f in items if f.align])
        line = _mode([f.line_value for f in items
                      if f.line_value is not None and f.line_rule == "exactly"])
        cur = s.roles[role]
        if size:
            cur["size"] = float(size)
        if ea:
            cur["font"] = ea
        if ascii_f:
            cur["ascii"] = ascii_f
        if bold is not None:
            cur["bold"] = bool(bold)
        if align:
            cur["align"] = align
        if line:
            cur["line"] = float(line)

    # ---- 行距（全局）----
    exact = [f.line_value for f in feats
             if f.line_rule == "exactly" and f.line_value is not None]
    if exact:
        s.line_spacing = float(Counter(exact).most_common(1)[0][0])

    # ---- 中文缩进：看正文首行缩进 ----
    bodies = [f for f in buckets.get("body", [])]
    if bodies:
        indents = [f.first_indent_pt for f in bodies if f.first_indent_pt]
        if indents:
            s.apply_indent = True
            # 首行缩进（磅）≈ 字号 × 字符数
            typical_size = s.roles.get("body", {}).get("size", 12.0)
            chars = round((Counter(indents).most_common(1)[0][0] / typical_size) * 2) / 2
            s.first_line_indent_chars = max(0.0, chars)
        else:
            s.apply_indent = False

    return s
