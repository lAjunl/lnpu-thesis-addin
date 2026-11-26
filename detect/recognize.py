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
from .features import (
    StyleIndex,
    _align_of,
    _ascii_of,
    _bold_of,
    _east_asia_of,
    _find_in_pPr,
    _get_w,
    _size_pt,
    _visible_runs,
    style_chain,
)
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


def _doc_defaults_rpr(doc):
    from docx.oxml.ns import qn as _qn
    try:
        defaults = doc.styles.element.find(_qn("w:docDefaults"))
        if defaults is not None:
            rpr_def = defaults.find(_qn("w:rPrDefault"))
            if rpr_def is not None:
                return rpr_def.find(_qn("w:rPr"))
    except Exception:      # noqa: BLE001
        pass
    return None


def _tbl_first_row_bold(tbl, tbl_chain) -> bool | None:
    """检查表格样式是否通过 tblStylePr(firstRow) 给表头设了加粗。"""
    from docx.oxml.ns import qn as _qn
    tblPr = tbl._tbl.find(_qn("w:tblPr"))
    if tblPr is not None:
        look = tblPr.find(_qn("w:tblLook"))
        if look is not None:
            fr = look.get(_qn("w:firstRow"))
            if fr in ("0", "false", "off"):
                return None
    for st in tbl_chain:
        for tsp in st.element.findall(_qn("w:tblStylePr")):
            if tsp.get(_qn("w:type")) != "firstRow":
                continue
            rpr = tsp.find(_qn("w:rPr"))
            if rpr is None:
                continue
            b = _find_in_pPr(rpr, "b")
            if b is not None:
                v = _get_w(b, "val")
                return True if v is None else v not in ("0", "false", "off")
    return None


def _border_pt(el) -> float | None:
    from docx.oxml.ns import qn as _qn
    if el is None:
        return None
    v = el.get(_qn("w:val"))
    if v in (None, "none", "nil"):
        return None
    sz = el.get(_qn("w:sz"))
    if not sz:
        return None
    try:
        return round(int(sz) / 8.0, 2)
    except ValueError:
        return None


def inspect_table_format(doc, tables=None, style_index: StyleIndex | None = None) -> dict:
    """从文档表格中实测样式、表内字体/字号、表头/表体加粗、对齐与线宽。

    未实测到的项一律返回 None，绝不拿默认值冒充识别结果。
    """
    from docx.oxml.ns import qn as _qn
    from docx.text.paragraph import Paragraph

    target = list(doc.tables[:15]) if tables is None else list(tables)
    if not target:
        return {
            "style": None, "font": None, "ascii": None, "size": None,
            "header_bold": None, "bold": None, "align": None, "pos": None,
            "line": None, "line_inner": None, "count": 0,
        }

    si = style_index if style_index is not None else StyleIndex(doc)
    defaults_rpr = _doc_defaults_rpr(doc)
    t_styles: list[str] = []
    t_fonts: list[str] = []
    t_asciis: list[str] = []
    t_sizes: list[float] = []
    t_header_bolds: list[bool] = []
    t_body_bolds: list[bool] = []
    t_aligns: list[str] = []
    t_pos: list[str] = []
    t_outer: list[float] = []
    t_inner: list[float] = []

    pos_map = {"left": "left", "start": "left", "center": "center",
               "right": "right", "end": "right"}
    w_tblPr = _qn("w:tblPr")
    w_tblBorders = _qn("w:tblBorders")
    w_jc = _qn("w:jc")
    w_val = _qn("w:val")
    w_tcPr = _qn("w:tcPr")
    w_tcBorders = _qn("w:tcBorders")
    w_bottom = _qn("w:bottom")
    w_p = _qn("w:p")
    w_pPr = _qn("w:pPr")

    for tbl in target:
        tblPr = tbl._tbl.find(w_tblPr)
        bdr = tblPr.find(w_tblBorders) if tblPr is not None else None
        t_st = si.table_style(tbl)
        st_name = (t_st.name if t_st is not None else "").lower()
        tbl_chain = si.chain_of(t_st)
        first_row_style_b = _tbl_first_row_bold(tbl, tbl_chain)

        if tblPr is not None:
            jc = tblPr.find(w_jc)
            if jc is not None:
                pv = pos_map.get((jc.get(w_val) or "").lower())
                if pv:
                    t_pos.append(pv)

        has_v = False
        has_h = False
        if bdr is not None:
            for side in ("left", "right", "insideV"):
                el = bdr.find(_qn("w:" + side))
                if el is not None and el.get(w_val) not in (None, "none", "nil"):
                    has_v = True
            for side in ("top", "bottom"):
                el = bdr.find(_qn("w:" + side))
                if el is not None and el.get(w_val) not in (None, "none", "nil"):
                    has_h = True
                    pt = _border_pt(el)
                    if pt:
                        t_outer.append(pt)
            for side in ("insideH", "insideV"):
                pt = _border_pt(bdr.find(_qn("w:" + side)))
                if pt:
                    t_inner.append(pt)

        tr_lst = tbl._tbl.tr_lst
        if has_v or "grid" in st_name:
            t_styles.append("grid")
        elif has_h:
            t_styles.append("three_line")
            # 三线表栏目线通常写在首行单元格的 tcBorders/bottom
            if tr_lst:
                for tc in tr_lst[0].tc_lst[:8]:
                    tcPr = tc.find(w_tcPr)
                    tb = tcPr.find(w_tcBorders) if tcPr is not None else None
                    pt = _border_pt(tb.find(w_bottom) if tb is not None else None)
                    if pt:
                        t_inner.append(pt)

        for r_idx, tr in enumerate(tr_lst[:6]):
            for tc in tr.tc_lst[:8]:
                for p_el in tc.findall(w_p):
                    p = Paragraph(p_el, tbl)
                    if not _visible_runs(p) or not p.text.strip():
                        continue
                    p_chain = si.para_chain(p)
                    full_chain = p_chain + [st for st in tbl_chain if st not in p_chain]
                    pPr = p_el.find(w_pPr)

                    ea = _east_asia_of(p, pPr, full_chain, defaults_rpr)
                    if ea:
                        t_fonts.append(ea)
                    asc = _ascii_of(p, pPr, full_chain, defaults_rpr)
                    if asc:
                        t_asciis.append(asc)
                    sz = _size_pt(p, pPr, full_chain, defaults_rpr)
                    if sz:
                        t_sizes.append(float(sz))
                    al = _align_of(pPr, full_chain)
                    if al:
                        t_aligns.append(al)

                    b = _bold_of(p, pPr, full_chain, defaults_rpr)
                    if r_idx == 0:
                        if b is None:
                            b = first_row_style_b
                        t_header_bolds.append(bool(b))
                    else:
                        t_body_bolds.append(bool(b))

    sz_mode = _mode(t_sizes)
    return {
        "style": _mode(t_styles),
        "font": _mode(t_fonts),
        "ascii": _mode(t_asciis),
        "size": float(sz_mode) if sz_mode is not None else None,
        "header_bold": _mode(t_header_bolds),
        "bold": _mode(t_body_bolds),
        "align": _mode(t_aligns),
        "pos": _mode(t_pos),
        "line": _mode(t_outer),
        "line_inner": _mode(t_inner),
        "count": len(target),
    }


def _is_hidden_run(r_el, _qn) -> bool:
    """判断页眉里的某个 w:r 是否被 Word 隐藏（vanish / webHidden / 纯白字）。"""
    rPr = r_el.find(_qn("w:rPr"))
    if rPr is None:
        return False
    for tag in ("w:vanish", "w:webHidden"):
        el = rPr.find(_qn(tag))
        if el is not None:
            v = el.get(_qn("w:val"))
            if v is None or v.lower() not in ("0", "false", "off"):
                return True
    col = rPr.find(_qn("w:color"))
    if col is not None:
        cv = (col.get(_qn("w:val")) or "").strip().lower()
        if cv == "ffffff":
            return True
    return False


def _header_para_visible_text(para) -> str:
    """提取页眉段落里真正可见的文字，过滤隐藏字、纯白字与纯页码域（PAGE/NUMPAGES）。"""
    import re
    from docx.oxml.ns import qn as _qn

    page_field_re = re.compile(r"\b(PAGE|NUMPAGES|SECTIONPAGES)\b", re.IGNORECASE)
    parts: list[str] = []
    # 域状态栈：每层 [phase('instr'|'result'), instr_buf, is_page_field]
    stack: list[list] = []

    def walk(el) -> None:
        tag = el.tag
        if tag == _qn("w:del") or tag == _qn("w:moveFrom"):
            return
        if tag == _qn("w:fldSimple"):
            instr = el.get(_qn("w:instr")) or ""
            if page_field_re.search(instr):
                return
            for ch in el:
                walk(ch)
            return
        if tag == _qn("w:r"):
            hidden = _is_hidden_run(el, _qn)
            for ch in el:
                ct = ch.tag
                if ct == _qn("w:fldChar"):
                    ft = ch.get(_qn("w:fldCharType"))
                    if ft == "begin":
                        stack.append(["instr", "", False])
                    elif ft == "separate" and stack:
                        top = stack[-1]
                        top[0] = "result"
                        top[2] = bool(page_field_re.search(top[1]))
                    elif ft == "end" and stack:
                        stack.pop()
                elif ct == _qn("w:instrText"):
                    if stack and stack[-1][0] == "instr":
                        stack[-1][1] += ch.text or ""
                elif ct == _qn("w:t"):
                    if hidden:
                        continue
                    if any(frame[2] for frame in stack):
                        continue
                    if ch.text:
                        parts.append(ch.text)
            return
        for ch in el:
            walk(ch)

    for child in para._p:
        walk(child)
    return "".join(parts).strip()


def _header_part_text(hdr) -> tuple[str, list]:
    """返回 (可见页眉文字, 有可见文字的段落列表)。"""
    texts: list[str] = []
    paras: list = []
    try:
        for p in hdr.paragraphs:
            t = _header_para_visible_text(p)
            if t:
                texts.append(t)
                paras.append(p)
    except Exception:      # noqa: BLE001
        return "", []
    return " ".join(texts).strip(), paras


def _section_body_buckets(doc) -> list[list]:
    """把 body 下每个块级元素（段落/表格）归到对应节，顺序与 doc.sections 对齐。"""
    from docx.oxml.ns import qn as _qn

    body = doc.element.body
    buckets: list[list] = []
    cur: list = []
    for child in body:
        if child.tag == _qn("w:sectPr"):
            continue
        cur.append(child)
        if child.tag == _qn("w:p"):
            pPr = child.find(_qn("w:pPr"))
            if pPr is not None and pPr.find(_qn("w:sectPr")) is not None:
                buckets.append(cur)
                cur = []
    buckets.append(cur)
    return buckets


def _section_stats(elements: list) -> tuple[int, int, bool]:
    """返回 (可见字符数, 非空段落数, 是否明确跨多页[分页符/lastRenderedPageBreak/页前分页])。"""
    from docx.oxml.ns import qn as _qn

    chars = 0
    non_empty = 0
    has_break = False
    p_idx = 0
    for el in elements:
        if el.find(".//" + _qn("w:lastRenderedPageBreak")) is not None:
            has_break = True
        for br in el.findall(".//" + _qn("w:br")):
            if br.get(_qn("w:type")) == "page":
                has_break = True
        if el.tag == _qn("w:p"):
            if p_idx > 0:
                pPr = el.find(_qn("w:pPr"))
                pb = pPr.find(_qn("w:pageBreakBefore")) if pPr is not None else None
                if pb is not None:
                    v = pb.get(_qn("w:val"))
                    if v is None or v.lower() not in ("0", "false", "off"):
                        has_break = True
            p_idx += 1
            txt = "".join(t.text or "" for t in el.findall(".//" + _qn("w:t"))).strip()
            if txt:
                chars += len(txt)
                non_empty += 1
        elif el.tag == _qn("w:tbl"):
            txt = "".join(t.text or "" for t in el.findall(".//" + _qn("w:t"))).strip()
            chars += len(txt)
            non_empty += len(el.findall(_qn("w:tr")))
    return chars, non_empty, has_break


def inspect_header_format(doc) -> dict:
    """识别文档中真正显示的页眉及其格式。

    关键防误判：
    - 某节开启「首页不同」(different_first_page_header_footer) 且本节只有 1 页时，
      Word 只显示 first_page_header，sec.header 里的模板残留文字（如旧校名）根本不会出现；
    - 封面/空节残留页眉但正文各节无页眉时，按各节正文体量加权投票，不因第一节残留误判；
    - 忽略隐藏字 (vanish)、纯白字 (FFFFFF) 与纯页码域 (PAGE/NUMPAGES)。
    """
    from docx.oxml.ns import qn as _qn

    empty = {
        "enabled": False, "text": "",
        "font": None, "ascii": None, "size": None, "bold": None,
        "line": None, "line_style": None, "line_color": None, "gap": None,
        "detected": False,
    }
    sections = list(doc.sections)
    if not sections:
        return empty

    buckets = _section_body_buckets(doc)
    has_any_lrpb = doc.element.body.find(".//" + _qn("w:lastRenderedPageBreak")) is not None
    try:
        odd_even = bool(doc.settings.odd_and_even_pages_header_footer)
    except Exception:      # noqa: BLE001
        odd_even = False

    sec_results: list[tuple[str, int, list]] = []
    for idx, sec in enumerate(sections):
        elems = buckets[idx] if idx < len(buckets) else []
        chars, non_empty, has_break = _section_stats(elems)
        # 判断本节是否会显示到第 2 页（决定 diff1=True 时 default header 是否可见）
        if has_break:
            multipage = True
        elif has_any_lrpb:
            # 文档保留了 Word 分页标记，本节无标记时只有在字数明显超一页时才视作多页
            multipage = chars >= 2000 or non_empty >= 36
        else:
            # 未经 Word 分页保存（如 WPS / python-docx 生成）
            multipage = chars >= 1000 or non_empty >= 22

        diff1 = bool(getattr(sec, "different_first_page_header_footer", False))
        def_txt, def_paras = _header_part_text(sec.header)
        first_txt, first_paras = (_header_part_text(sec.first_page_header)
                                  if diff1 else ("", []))
        even_txt, even_paras = (_header_part_text(sec.even_page_header)
                                if (odd_even and multipage) else ("", []))

        if not diff1:
            vis_txt, vis_paras = def_txt, def_paras
            if not vis_txt and even_txt:
                vis_txt, vis_paras = even_txt, even_paras
        elif multipage:
            if def_txt:
                vis_txt, vis_paras = def_txt, def_paras
            elif even_txt:
                vis_txt, vis_paras = even_txt, even_paras
            else:
                vis_txt, vis_paras = first_txt, first_paras
        else:
            # 1 页且「首页不同」：Word 只渲染首页页眉，绝不显示 sec.header 的残留内容
            vis_txt, vis_paras = first_txt, first_paras

        sec_results.append((vis_txt, chars, vis_paras))

    # 有正文的节优先参与投票，排除全空节
    non_zero = [r for r in sec_results if r[1] > 0]
    voters = non_zero if non_zero else [(t, 1, ps) for t, _, ps in sec_results]

    w_has = sum(w for t, w, _ in voters if t)
    w_empty = sum(w for t, w, _ in voters if not t)
    if w_has <= 0 or w_has < w_empty:
        return empty

    by_text: dict[str, int] = {}
    paras_by_text: dict[str, list] = {}
    for t, w, ps in voters:
        if not t:
            continue
        by_text[t] = by_text.get(t, 0) + w
        paras_by_text.setdefault(t, []).extend(ps)
    winner = max(by_text.items(), key=lambda kv: kv[1])[0]
    win_paras = paras_by_text.get(winner, [])

    defaults_rpr = _doc_defaults_rpr(doc)
    h_fonts: list[str] = []
    h_asciis: list[str] = []
    h_sizes: list[float] = []
    h_bolds: list[bool] = []
    h_lines: list[float] = []
    h_lstyles: list[str] = []
    h_lcolors: list[str] = []
    h_gaps: list[float] = []

    si = StyleIndex(doc)
    for p in win_paras[:6]:
        p_chain = si.para_chain(p)
        pPr = p._p.find(_qn("w:pPr"))
        ea = _east_asia_of(p, pPr, p_chain, defaults_rpr)
        if ea:
            h_fonts.append(ea)
        asc = _ascii_of(p, pPr, p_chain, defaults_rpr)
        if asc:
            h_asciis.append(asc)
        sz = _size_pt(p, pPr, p_chain, defaults_rpr)
        if sz:
            h_sizes.append(float(sz))
        b = _bold_of(p, pPr, p_chain, defaults_rpr)
        if b is not None:
            h_bolds.append(bool(b))
        # 页眉下横线（w:pPr/w:pBdr/w:bottom，逐级回退样式链）
        bdr_el = None
        if pPr is not None:
            pb = pPr.find(_qn("w:pBdr"))
            if pb is not None:
                bdr_el = pb.find(_qn("w:bottom"))
        if bdr_el is None:
            for st in p_chain:
                sp = st.element.find(_qn("w:pPr"))
                pb = sp.find(_qn("w:pBdr")) if sp is not None else None
                if pb is not None:
                    bdr_el = pb.find(_qn("w:bottom"))
                    if bdr_el is not None:
                        break
        if bdr_el is not None:
            val = bdr_el.get(_qn("w:val"))
            if val in ("none", "nil"):
                h_lines.append(0.0)
            else:
                pt = _border_pt(bdr_el)
                if pt is not None:
                    h_lines.append(pt)
                if val in ("single", "double", "dashed", "dotted", "dashDot"):
                    h_lstyles.append(val)
                col = bdr_el.get(_qn("w:color"))
                if col and col.lower() != "auto":
                    h_lcolors.append(col.upper())
                sp_v = bdr_el.get(_qn("w:space"))
                if sp_v and sp_v.isdigit():
                    h_gaps.append(float(int(sp_v)))

    sz_mode = _mode(h_sizes)
    return {
        "enabled": True,
        "text": winner,
        "font": _mode(h_fonts),
        "ascii": _mode(h_asciis),
        "size": float(sz_mode) if sz_mode is not None else None,
        "bold": _mode(h_bolds),
        "line": _mode(h_lines),
        "line_style": _mode(h_lstyles),
        "line_color": _mode(h_lcolors),
        "gap": _mode(h_gaps),
        "detected": True,
    }


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

    # ---- 页眉：只认真正显示的页眉，过滤首页不同残留 / 封面残留 / 隐藏字 ----
    hdr_info = inspect_header_format(doc)
    s.enable_header = bool(hdr_info["enabled"])
    s.header_text = hdr_info["text"] or ""
    if hdr_info["enabled"]:
        if hdr_info.get("font"):
            s.header_font = hdr_info["font"]
        if hdr_info.get("ascii"):
            s.header_ascii = hdr_info["ascii"]
        if hdr_info.get("size") is not None:
            s.header_size = float(hdr_info["size"])
        if hdr_info.get("bold") is not None:
            s.header_bold = bool(hdr_info["bold"])
        if hdr_info.get("line") is not None:
            s.header_line = float(hdr_info["line"])
        if hdr_info.get("line_style"):
            s.header_line_style = hdr_info["line_style"]
        if hdr_info.get("line_color"):
            s.header_line_color = hdr_info["line_color"]
        if hdr_info.get("gap") is not None:
            s.header_gap = float(hdr_info["gap"])
    if report is not None:
        report["page"] = dict(s.page)
        report["page"]["header_text"] = s.header_text
        report["page"]["enable_header"] = s.enable_header
        report["header"] = hdr_info

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
            cur["line_rule"] = setting[0]
            if setting[1] is not None:
                cur["line"] = setting[1]

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

    # ---- 表格样式与表内字体（含表头/表体加粗、西文字体、对齐、线宽）----
    tbl_info = inspect_table_format(doc)
    if tbl_info.get("style"):
        s.table_style = tbl_info["style"]
    if tbl_info.get("font"):
        s.table_font = tbl_info["font"]
    if tbl_info.get("ascii"):
        s.table_ascii = tbl_info["ascii"]
    if tbl_info.get("size") is not None:
        s.table_font_size = float(tbl_info["size"])
    if tbl_info.get("header_bold") is not None:
        s.table_header_bold = bool(tbl_info["header_bold"])
    if tbl_info.get("bold") is not None:
        s.table_bold = bool(tbl_info["bold"])
    if tbl_info.get("align"):
        s.table_align = tbl_info["align"]
    if tbl_info.get("pos"):
        s.table_pos = tbl_info["pos"]
    if tbl_info.get("line") is not None:
        s.table_line = float(tbl_info["line"])
    if tbl_info.get("line_inner") is not None:
        s.table_line_inner = float(tbl_info["line_inner"])
    if report is not None:
        report["table"] = tbl_info

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
            cur["line_rule"] = setting[0]
            if setting[1] is not None:
                cur["line"] = setting[1]
    page = report.get("page") or {}
    for key in ("top", "bottom", "left", "right", "header", "footer"):
        if page.get(key) is not None:
            s.page[key] = float(page[key])
    hdr = report.get("header")
    if isinstance(hdr, dict):
        if hdr.get("enabled") is not None:
            s.enable_header = bool(hdr["enabled"])
        if s.enable_header:
            if hdr.get("text") is not None:
                s.header_text = str(hdr["text"])
            if hdr.get("font"):
                s.header_font = hdr["font"]
            if hdr.get("ascii"):
                s.header_ascii = hdr["ascii"]
            if hdr.get("size") is not None:
                s.header_size = float(hdr["size"])
            if hdr.get("bold") is not None:
                s.header_bold = bool(hdr["bold"])
            if hdr.get("line") is not None:
                s.header_line = float(hdr["line"])
            if hdr.get("line_style"):
                s.header_line_style = hdr["line_style"]
            if hdr.get("line_color"):
                s.header_line_color = hdr["line_color"]
            if hdr.get("gap") is not None:
                s.header_gap = float(hdr["gap"])
        else:
            s.header_text = ""
    else:
        if page.get("enable_header") is not None:
            s.enable_header = bool(page["enable_header"])
        if page.get("header_text") is not None:
            s.header_text = page["header_text"]
        if not s.enable_header:
            s.header_text = ""
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
    if tbl.get("ascii"):
        s.table_ascii = tbl["ascii"]
    if tbl.get("size") is not None:
        s.table_font_size = float(tbl["size"])
    if tbl.get("header_bold") is not None:
        s.table_header_bold = bool(tbl["header_bold"])
    if tbl.get("bold") is not None:
        s.table_bold = bool(tbl["bold"])
    if tbl.get("align"):
        s.table_align = tbl["align"]
    if tbl.get("pos"):
        s.table_pos = tbl["pos"]
    if tbl.get("line") is not None:
        s.table_line = float(tbl["line"])
    if tbl.get("line_inner") is not None:
        s.table_line_inner = float(tbl["line_inner"])
    return s
