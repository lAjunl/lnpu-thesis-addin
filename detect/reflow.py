"""重排引擎：把"角色 → 格式"套用到文档。

设计取舍（重要）：
    只动字体名、字号、加粗，以及段落的行距/对齐/段后。
    **不动**斜体、下划线、颜色、上下标——正文里的上标引用、强调格式要保留。
    这样排版不会把作者已有的强调信息抹掉。
"""
from __future__ import annotations

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

from .classify import effective_role
from .labels import LABEL_ROLES

ALIGN_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "both": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def _set_rpr_font(rPr, cjk: str, latin: str, size: float, bold: bool,
                  family_only: bool = False) -> None:
    """写 rPr 里的字体：中文槽位用 cjk，英文/数字槽位用 latin。

    四个槽位都要写，并且清掉 *Theme——只写 w:ascii 或残留主题字体，
    中文会回退成 Calibri / 等线（实测"字体不合适"就是这个原因）。
    family_only=True 时只统一字体族，不动字号/加粗（给上/下标用）。
    """
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn("w:ascii"), latin)
    rFonts.set(qn("w:hAnsi"), latin)
    rFonts.set(qn("w:eastAsia"), cjk)
    rFonts.set(qn("w:cs"), latin)
    for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rFonts.get(qn(attr)) is not None:
            del rFonts.attrib[qn(attr)]
    if family_only:
        return
    sz = rPr.get_or_add_sz()
    sz.set(qn("w:val"), str(int(round(size * 2))))
    szCs = rPr.find(qn("w:szCs"))
    if szCs is None:
        szCs = OxmlElement("w:szCs")
        sz.addnext(szCs)
    szCs.set(qn("w:val"), str(int(round(size * 2))))
    _remove_toggle(rPr, "b")
    if bold:
        rPr.get_or_add_b()


def set_run_font(run, cjk: str, latin: str, size: float, bold: bool,
                 color: str = "") -> None:
    """给一个 run 写字体（中文/西文分开）+ 字号 + 加粗 + 字色。"""
    rPr = run._r.get_or_add_rPr()
    _set_rpr_font(rPr, cjk, latin, size, bold)
    for el in rPr.findall(qn("w:color")):
        rPr.remove(el)
    if color:
        rPr.get_or_add_color().set(qn("w:val"), color.replace("#", ""))


def _set_para_default_font(paragraph, p: dict) -> None:
    """给没有 run 的空段落设置默认字体。"""
    pPr = paragraph._p.get_or_add_pPr()
    rPr = pPr.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        pPr.append(rPr)
    _set_rpr_font(rPr, p["font"], p.get("ascii", p["font"]),
                  float(p["size"]), bool(p.get("bold")))


def _remove_toggle(rPr, tag: str) -> None:
    for el in rPr.findall(qn("w:" + tag)):
        rPr.remove(el)
    for el in rPr.findall(qn("w:" + tag + "Cs")):
        rPr.remove(el)


RULE_MAP = {
    "exact": WD_LINE_SPACING.EXACTLY,
    "minimum": WD_LINE_SPACING.AT_LEAST,
    "single": WD_LINE_SPACING.SINGLE,
    "multiple": WD_LINE_SPACING.ONE_POINT_FIVE,
    "multiple2": WD_LINE_SPACING.DOUBLE,
}
RULE_MULT = {"multiple": 1.5, "multiple2": 2.0}


def _set_line_spacing(para, p: dict) -> None:
    """按设置写行距：固定值 / 单倍 / 1.5倍 / 2倍 / 最小值。"""
    rule = p.get("line_rule") or "exact"
    pf = para.paragraph_format
    try:
        if rule in ("multiple", "multiple2"):
            pf.line_spacing_rule = RULE_MAP[rule]
            pf.line_spacing = RULE_MULT[rule]
        elif rule == "single":
            pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
        elif rule == "minimum":
            pf.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
            pf.line_spacing = Pt(float(p.get("line", 22)))
        else:
            pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            pf.line_spacing = Pt(float(p.get("line", 22)))
    except Exception:
        pass


def _set_first_indent(para, chars: float | None, size: float = 12.0) -> None:
    """正文首行缩进：chars 个字符（None = 顶格，清掉缩进）。

    同时写 `w:firstLineChars` 和 `w:firstLine`：
    Word 优先按字符数缩进（换字号也是"2 个字"），`w:firstLine` 只是兜底，
    给不认 Chars 属性的软件用。两种写法都要先清掉原来的 hanging。
    """
    pPr = para._p.get_or_add_pPr()
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        if not chars:
            return
        ind = pPr.get_or_add_ind()
    for attr in ("w:firstLineChars", "w:firstLine", "w:hanging",
                 "w:hangingChars"):
        if ind.get(qn(attr)) is not None:
            del ind.attrib[qn(attr)]
    if chars and chars > 0:
        ind.set(qn("w:firstLineChars"), str(int(round(chars * 100))))
        ind.set(qn("w:firstLine"), str(int(round(chars * size * 20))))


def _set_equation_tabs(para, content_twips: float) -> None:
    """公式段落：写「居中 + 右对齐」两个制表位，编号才贴得到右边距。

    Word 里"公式 + 右编号"靠制表符分列（MathType 就这么排），
    源文件常常只剩制表符、没有制表位 → 编号紧跟在公式后面。
    这里按正文栏宽重写制表位，并把段落改成左对齐（制表位才有意义）。
    """
    p_el = para._p
    if content_twips <= 0 or p_el.find(".//" + qn("w:tab")) is None:
        return
    from .features import equation_label
    if not equation_label(para):
        return                       # 无编号的居中公式，不掺和
    pPr = p_el.get_or_add_pPr()
    for old in pPr.findall(qn("w:tabs")):
        pPr.remove(old)
    tabs = pPr.get_or_add_tabs()
    for pos, val in ((content_twips / 2.0, "center"),
                     (content_twips, "right")):
        el = OxmlElement("w:tab")
        el.set(qn("w:val"), val)
        el.set(qn("w:pos"), str(int(round(pos))))
        tabs.append(el)
    para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT


def indent_chars_of(settings) -> float:
    """从设置里算正文首行缩进字符数：0 表示顶格。预览和排版共用这一口径。"""
    if getattr(settings, "apply_not_indent", False):
        return 0.0
    if not getattr(settings, "apply_indent", True):
        return 0.0
    return max(0.0, float(getattr(settings, "first_line_indent_chars", 2.0) or 0.0))


def _bold_leading_label(para, base_bold: bool = False) -> bool:
    """把段落开头的 "Abstract:" / "关键词:" 标签加粗，正文部分不动。

    只加粗标签这几个字（期刊习惯），所以要把第一个 run 拆成两段。
    """
    from .labels import split_label

    runs = para.runs
    if not runs:
        return False
    label, _ = split_label("".join(r.text for r in runs))
    if label is None:
        return False
    r0 = runs[0]
    if not r0.text.startswith(label):
        return False            # 标签跨 run，极少见，退化为不做
    import copy

    leftover = r0.text[len(label):]
    r0.text = label
    r0.font.bold = True
    if leftover:
        new_r = copy.deepcopy(r0._r)
        r0._r.addnext(new_r)
        from docx.text.run import Run

        nr = Run(new_r, para)
        nr.text = leftover
        nr.font.bold = base_bold
    return True


def _apply_para(para, p: dict, role: str | None = None,
                label_bold: bool = False, indent_chars: float | None = None) -> None:
    """把一份格式字典写进一个段落（行距/对齐/首行缩进 + 字体/字号/加粗/字色）。

    indent_chars: 正文段落的首行缩进字符数。None = 不处理，0 = 顶格（清掉缩进）。
    只作用于 body：其它角色（尤其参考文献的悬挂缩进）不能被改。
    """
    pf = para.paragraph_format
    _set_line_spacing(para, p)
    align = ALIGN_MAP.get(p.get("align", "left"))
    if align is not None:
        pf.alignment = align
    if role == "body" and indent_chars is not None:
        _set_first_indent(para, indent_chars,
                          float(p.get("size", 12) or 12))

    # 字色：留空表示"不指定"（沿用 Word 的自动色），写 w:color 会覆盖掉
    color = (p.get("color") or "").strip()

    if para.runs:
        cjk = p.get("font") or "宋体"
        latin = p.get("ascii") or cjk
        size = float(p["size"])
        bold = bool(p.get("bold"))
        # 一次遍历**所有** w:r（含超链接 / 域 / sdt 内部的）。
        # 不能用 paragraph.runs + id() 去重：lxml 的元素代理每次访问都是新对象，
        # id() 不可靠；而超链接里的 run 本来就不在 paragraph.runs 里，
        # 这正是"图注里 [1, 3, 4] 字体和周围不一样"的原因。
        for r_el in para._p.iter(qn("w:r")):
            rPr = r_el.get_or_add_rPr()
            va = rPr.find(qn("w:vertAlign"))
            script = va is not None and va.get(qn("w:val")) in ("superscript",
                                                                "subscript")
            _set_rpr_font(rPr, cjk, latin, size, bold, family_only=script)
            for el in rPr.findall(qn("w:color")):
                rPr.remove(el)
            if color:
                rPr.get_or_add_color().set(qn("w:val"), color.replace("#", ""))
    else:
        _set_para_default_font(para, p)

    if label_bold and role in LABEL_ROLES:
        _bold_leading_label(para, bool(p.get("bold")))


def _table_fonts(settings) -> tuple[str, str, float, bool, str]:
    return (
        getattr(settings, "table_font", "") or "宋体",
        getattr(settings, "table_ascii", "") or "Times New Roman",
        float(getattr(settings, "table_font_size", 10.5) or 10.5),
        bool(getattr(settings, "table_bold", False)),
        (getattr(settings, "table_color", "") or "").strip(),
    )


def _apply_table_fonts(doc, settings) -> int:
    """表内文字统一成设置的字体/字号/加粗/字色/对齐（表头与数据行一视同仁）。"""
    cjk, latin, size, bold, color = _table_fonts(settings)
    align = ALIGN_MAP.get(getattr(settings, "table_align", "left"))
    n = 0
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if align is not None:
                        para.paragraph_format.alignment = align
                    if para.runs:
                        for run in para.runs:
                            set_run_font(run, cjk, latin, size, bold, color)
                    else:
                        _set_para_default_font(
                            para, {"font": cjk, "ascii": latin, "size": size,
                                   "bold": bold})
        n += 1
    return n


_TABLE_POS = {"left": "LEFT", "center": "CENTER", "right": "RIGHT"}


def _apply_table_position(doc, settings) -> int:
    """表格在页面里的位置（左 / 居中 / 右）。"""
    try:
        from docx.enum.table import WD_TABLE_ALIGNMENT
    except Exception:
        return 0
    pos = _TABLE_POS.get(getattr(settings, "table_pos", "center"))
    if pos is None:
        return 0
    attr = getattr(WD_TABLE_ALIGNMENT, pos, None)
    if attr is None:
        return 0
    n = 0
    for table in doc.tables:
        try:
            table.alignment = attr
        except Exception:
            continue
        n += 1
    return n


def style_tables(doc, settings) -> int:
    """表格样式：三线表（两粗一细：顶/底线粗、栏目线细）/ 网格 / 无线；并统一表内字体。"""
    style = getattr(settings, "table_style", "three_line")
    outer = float(getattr(settings, "table_line", 1.5) or 0)
    inner = float(getattr(settings, "table_line_inner", 0.75) or 0)
    _apply_table_fonts(doc, settings)
    _apply_table_position(doc, settings)
    n = 0
    for table in doc.tables:
        tblPr = table._tbl.tblPr
        _remove(tblPr, "tblBorders")
        borders = OxmlElement("w:tblBorders")
        sides = ("top", "bottom", "left", "right", "insideH", "insideV")
        if style == "grid":
            vals = {k: "single" for k in sides}
        elif style == "none":
            vals = {k: "none" for k in sides}
        else:   # three_line
            vals = {"top": "single", "bottom": "single",
                    "left": "none", "right": "none",
                    "insideH": "none", "insideV": "none"}
        for side in sides:
            val = vals[side]
            el = OxmlElement("w:" + side)
            el.set(qn("w:val"), val)
            w_pt = inner if (style == "grid" and side in ("insideH", "insideV") and inner > 0) else outer
            if val == "single" and w_pt > 0:
                el.set(qn("w:sz"), str(max(1, int(round(w_pt * 8)))))
                el.set(qn("w:color"), "000000")
            borders.append(el)
        tblPr.append(borders)

        # 清掉旧单元格边框，避免原网格线在三线表/无线表里残留
        for row in table.rows:
            for cell in row.cells:
                tcPr = cell._tc.find(qn("w:tcPr"))
                if tcPr is not None:
                    _remove(tcPr, "tcBorders")

        # 三线表的栏目线（表头下线）：用较细的 inner（默认 0.75 磅）
        head_w = inner if inner > 0 else outer
        if style == "three_line" and head_w > 0 and table.rows:
            for cell in table.rows[0].cells:
                tcPr = cell._tc.get_or_add_tcPr()
                tb = OxmlElement("w:tcBorders")
                el = OxmlElement("w:bottom")
                el.set(qn("w:val"), "single")
                el.set(qn("w:sz"), str(max(1, int(round(head_w * 8)))))
                el.set(qn("w:color"), "000000")
                tb.append(el)
                tcPr.append(tb)
        n += 1
    return n


def _remove(parent, tag: str) -> None:
    if parent is None:
        return
    for el in parent.findall(qn("w:" + tag)):
        parent.remove(el)


def _remove_color(run) -> None:
    rPr = run._r.get_or_add_rPr()
    for el in rPr.findall(qn("w:color")):
        rPr.remove(el)


def apply_formats(doc, analysis, params: dict,
                  overrides: dict | None = None,
                  label_bold: bool = False,
                  indent_chars: float | None = None) -> dict:
    """按角色套格式。返回统计 {角色: 段数}。

    overrides: {段索引: 格式字典}，用户在预览里手动改的段落，
    在该段上覆盖角色格式（相当于 Word 的"直接格式"）。
    label_bold: 是否把段落开头的 Abstract: / 关键词: 标签加粗。
    indent_chars: 正文首行缩进字符数，用 indent_chars_of(settings) 算，
                  None = 不动缩进。
    """
    content_twips = 0.0
    try:
        sec = doc.sections[0]
        content_twips = (float(sec.page_width) - float(sec.left_margin)
                         - float(sec.right_margin)) / 635.0   # EMU → twips
    except Exception:
        content_twips = 0.0

    stats: dict[str, int] = {}
    for d in analysis.detections:
        role = effective_role(d)
        p = params.get(role)
        if not p:
            continue
        p = dict(p)
        ov = (overrides or {}).get(d.index)
        if ov:
            p.update({k: v for k, v in ov.items() if v is not None})
        _apply_para(doc.paragraphs[d.index], p, role=role,
                    label_bold=label_bold, indent_chars=indent_chars)
        if role == "equation":
            _set_equation_tabs(doc.paragraphs[d.index], content_twips)
        stats[role] = stats.get(role, 0) + 1
    return stats
