"""从原稿识别格式：读文档里各角色的实际生效格式，填进设置。

「从原稿识别」按钮的背后就是这个——不是让用户手填，
而是先把原稿的格式读出来当起点。

除了产出 `Settings`，还会填一份 **report**（`report` 参数）：
每个角色识别到了哪些属性、哪些没识别到。给「格式识别」弹窗的核对表用，
避免把"没识别到"和"用了预设默认值"混为一谈。
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
    return round(length.mm, 1)


def _line_setting(rule, value):
    """把 features 的 (rule, value) 映射成角色字典的 (line_rule, line)。"""
    if rule is None:
        return None
    if rule == "multiple":
        return ("multiple" if float(value or 1.0) < 1.75 else "multiple2",
                float(value or 1.0))
    if rule == "at_least":
        return ("minimum", float(value)) if value is not None else ("minimum", None)
    if rule == "exactly":
        return ("exact", float(value)) if value is not None else ("exact", None)
    if rule == "single":
        return ("single", None)
    return None


def _line_text(setting) -> str:
    if not setting:
        return "—"
    rule, value = setting
    if rule == "multiple":
        return "1.5 倍"
    if rule == "multiple2":
        return "2 倍"
    if rule == "minimum":
        return f"最小 {value:g} 磅"
    if rule == "single":
        return "单倍"
    return f"{value:g} 磅"


def recognize_from_doc(doc, analysis, report: dict | None = None) -> Settings:
    s = Settings()
    feats = analysis.features

    if report is not None:
        report["source"] = getattr(analysis, "source", "") or ""
        report["roles"] = {}
        report["page"] = {}
        report["indent"] = {}

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
        # 页眉文字：封面节通常没页眉，向后扫到第一个非空页眉
        for sec_h in doc.sections:
            try:
                ht = "".join(
                    p.text.strip() for p in sec_h.header.paragraphs if p.text.strip()
                )
                if ht:
                    s.header_text = ht
                    break
            except Exception:
                pass
    if report is not None:
        report["page"] = dict(s.page)
        report["page"]["header_text"] = s.header_text

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
        setting = _mode([_line_setting(f.line_rule, f.line_value) for f in items])
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
        if setting:
            cur["line_rule"], cur["line"] = setting[0], setting[1]

        if report is not None:
            report["roles"][role] = {
                "size": float(size) if size else None,
                "font": ea or None,
                "ascii": ascii_f or None,
                "bold": None if bold is None else bool(bold),
                "line": setting,
                "line_text": _line_text(setting),
                "align": align or None,
                "count": len(items),
            }

    # 没出现过的角色也要在报告里占位（表格要显示"未识别"）
    if report is not None:
        for role in s.roles:
            report["roles"].setdefault(role, {
                "size": None, "font": None, "ascii": None, "bold": None,
                "line": None, "line_text": "—", "align": None, "count": 0})

    # ---- 行距（全局）----
    exact = [f.line_value for f in feats
             if f.line_rule == "exactly" and f.line_value is not None]
    if exact:
        s.line_spacing = float(Counter(exact).most_common(1)[0][0])

    # ---- 正文首行缩进 ----
    bodies = buckets.get("body", [])
    chars = None
    if bodies:
        indents = [f.first_indent_pt for f in bodies if f.first_indent_pt]
        if indents:
            s.apply_indent = True
            typical_size = s.roles.get("body", {}).get("size", 12.0)
            chars = round((Counter(indents).most_common(1)[0][0] / typical_size) * 2) / 2
            s.first_line_indent_chars = max(0.0, chars)
        else:
            s.apply_indent = False
    if report is not None:
        report["indent"] = {"apply": bool(s.apply_indent), "chars": chars}

    # ---- 表格样式与表内字体 ----
    if doc.tables:
        from docx.oxml.ns import qn as _qn
        t_styles = []
        t_fonts = []
        t_sizes = []
        for tbl in doc.tables[:15]:
            tblPr = tbl._tbl.find(_qn("w:tblPr"))
            bdr = tblPr.find(_qn("w:tblBorders")) if tblPr is not None else None
            st_name = (tbl.style.name if tbl.style else "").lower()
            has_v = False
            has_h = False
            if bdr is not None:
                for side in ("left", "right", "insideV"):
                    el = bdr.find(_qn("w:" + side))
                    if el is not None and el.get(_qn("w:val")) not in (None, "none", "nil"):
                        has_v = True
                for side in ("top", "bottom"):
                    el = bdr.find(_qn("w:" + side))
                    if el is not None and el.get(_qn("w:val")) not in (None, "none", "nil"):
                        has_h = True
            if has_v or "grid" in st_name:
                t_styles.append("grid")
            elif has_h:
                t_styles.append("three_line")
            for row in tbl.rows[:4]:
                for cell in row.cells[:6]:
                    for p in cell.paragraphs:
                        if not p.text.strip():
                            continue
                        for r_el in p._p.iter(_qn("w:r")):
                            if not "".join(t.text or "" for t in r_el.findall(_qn("w:t"))).strip():
                                continue
                            rpr = r_el.find(_qn("w:rPr"))
                            if rpr is None:
                                continue
                            rf = rpr.find(_qn("w:rFonts"))
                            if rf is not None and rf.get(_qn("w:eastAsia")):
                                t_fonts.append(rf.get(_qn("w:eastAsia")))
                            se = rpr.find(_qn("w:sz"))
                            if se is not None and se.get(_qn("w:val")):
                                try:
                                    t_sizes.append(int(se.get(_qn("w:val"))) / 2.0)
                                except ValueError:
                                    pass
        if t_styles:
            s.table_style = _mode(t_styles) or s.table_style
        if t_fonts:
            s.table_font = _mode(t_fonts) or s.table_font
        if t_sizes:
            s.table_font_size = float(_mode(t_sizes) or s.table_font_size)
        if report is not None:
            report["table"] = {
                "style": s.table_style,
                "font": s.table_font,
                "size": s.table_font_size,
            }

    return s


def settings_from_report(base: Settings, report: dict) -> Settings:
    """把（可能被用户改过的）识别报告套用到一份设置上。

    只有报告里非 None 的属性才会覆盖；`—`（未识别）保持 base 的值。
    """
    s = Settings.from_json(base.to_json())
    for role, attrs in (report.get("roles") or {}).items():
        if role not in s.roles:
            continue
        cur = s.roles[role]
        if attrs.get("size"):
            cur["size"] = float(attrs["size"])
        if attrs.get("font"):
            cur["font"] = attrs["font"]
        if attrs.get("ascii"):
            cur["ascii"] = attrs["ascii"]
        if attrs.get("bold") is not None:
            cur["bold"] = bool(attrs["bold"])
        if attrs.get("align"):
            cur["align"] = attrs["align"]
        setting = attrs.get("line")
        if setting:
            cur["line_rule"], cur["line"] = setting[0], setting[1]
    page = report.get("page") or {}
    for key in ("top", "bottom", "left", "right", "header", "footer"):
        if page.get(key) is not None:
            s.page[key] = float(page[key])
    if page.get("header_text") is not None:
        s.header_text = page["header_text"]
    ind = report.get("indent") or {}
    if ind.get("apply") is not None:
        s.apply_indent = bool(ind["apply"])
    if ind.get("chars") is not None:
        s.first_line_indent_chars = float(ind["chars"])
    tbl = report.get("table") or {}
    if tbl.get("style"):
        s.table_style = tbl["style"]
    if tbl.get("font"):
        s.table_font = tbl["font"]
    if tbl.get("size"):
        s.table_font_size = float(tbl["size"])
    return s
