"""格式检查：把文档实际格式与当前设置对比，列出偏差。"""
from __future__ import annotations

from .classify import effective_role, is_heading_role


def check_document(doc, analysis, settings) -> list[str]:
    findings: list[str] = []
    feats = analysis.features

    # ---- 页面 ----
    if doc.sections:
        sec = doc.sections[0]
        pairs = [
            ("top", "上边距", "mm"), ("bottom", "下边距", "mm"),
            ("left", "左边距", "mm"), ("right", "右边距", "mm"),
            ("header", "页眉距离", "mm"), ("footer", "页脚距离", "mm"),
        ]
        for key, label, unit in pairs:
            try:
                actual = round(getattr(sec, f"{key}_margin").cm, 2)
            except Exception:
                continue
            want = settings.page.get(key)
            if want and abs(actual - want) > 0.5:
                findings.append(f"{label} 实际 {actual}mm，应为 {want}mm")

    # ---- 角色格式 ----
    seen: dict[str, dict] = {}
    for d in analysis.detections:
        role = effective_role(d)
        if role not in settings.roles:
            continue
        seen.setdefault(role, []).append(feats[d.index])

    for role, items in seen.items():
        p = settings.roles[role]
        name = {"heading1": "一级标题", "heading2": "二级标题", "heading3": "三级标题",
                "body": "正文", "figure_caption": "图注", "table_caption": "表注"}.get(role, role)
        sizes = [f.size_pt for f in items if f.size_pt is not None]
        if sizes:
            bad = [x for x in sizes if abs(x - float(p["size"])) > 0.1]
            if bad:
                findings.append(
                    f"{name}：{len(bad)}/{len(sizes)} 段字号不是 {p['size']}pt"
                )
        fonts = [f.east_asia or f.ascii_font for f in items
                 if (f.east_asia or f.ascii_font)]
        if fonts:
            want = p.get("font", "宋体")
            bad = [x for x in fonts if x != want]
            if bad:
                findings.append(
                    f"{name}：{len(bad)}/{len(fonts)} 段字体不是 {want}（如 {bad[0]}）"
                )
        lines = [f.line_value for f in items
                 if f.line_value is not None and f.line_rule == "exactly"]
        if lines:
            bad = [x for x in lines if abs(x - float(p["line"])) > 0.1]
            if bad:
                findings.append(
                    f"{name}：{len(bad)}/{len(lines)} 段行距不是固定值 {p['line']} 磅"
                )

    # ---- 编号 ----
    headings = [d for d in analysis.detections if is_heading_role(d.role)]
    if headings:
        no_num = [d for d in headings if not d.number]
        if no_num and settings.enable_number_heading:
            findings.append(
                f"标题编号：{len(no_num)}/{len(headings)} 个标题没有编号"
            )
    return findings
