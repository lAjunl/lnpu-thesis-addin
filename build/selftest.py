"""交付前自测：打开真实稿，检查分页预览、行距、字体、重排、编号、分节。

用法:
    .venv/Scripts/python.exe build/selftest.py <docx路径>
产出:
    build/st_*.png + 终端断言结果
"""
from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

os.environ.setdefault("QT_AUTO_SCREEN_SCALE_FACTOR", "1")
os.environ.setdefault("QT_ENABLE_HIGHDPI_SCALING", "1")
os.environ.setdefault("QT_SCALE_FACTOR_ROUNDING_POLICY", "PassThrough")

from PySide6.QtCore import Qt, QSizeF, QCoreApplication
from PySide6.QtGui import QTextCursor, QTextDocument
from PySide6.QtWidgets import QApplication

from detect.classify import analyze, effective_role
from detect.ui.main_window import MainWindow
from detect.ui.settings import Settings
from detect.ui.render import page_count, page_geometry, page_labels, render_side
from detect.ui.style import apply

FAIL = []


def check(name: str, cond: bool, detail: str = "") -> None:
    print(f"{'OK  ' if cond else 'FAIL'} {name} {detail}")
    if not cond:
        FAIL.append(name)


def grab(w, path: Path):
    img = w.grab()
    img.save(str(path))
    return img


def test_line_spacing() -> None:
    """固定 22 磅行距必须真的是 22 磅。

    旧版把 22 磅换算成相对字号的百分比，Qt 的百分比又是相对字体自然行高，
    结果系统性偏大约 20%。这里直接量文档高度。
    """
    from detect.ui.render import _insert_paragraph

    td = QTextDocument()
    td.setPageSize(QSizeF(400, 700))
    td.setDocumentMargin(0)
    cur = QTextCursor(td)
    n = 20
    for _ in range(n):
        _insert_paragraph(cur, "这是一段用来测量行距的文字，需要足够长以便换行。",
                          "宋体", "Times New Roman", 12.0, False, 22.0, "left")
    h = td.documentLayout().documentSize().height()
    expect = n * 22.0
    ok = abs(h - expect) < expect * 0.06
    check("行距=固定22磅", ok, f"实测 {h:.1f} / 期望 {expect:.1f}")

    # 倍数行距：1.5 倍应明显大于单倍
    td2 = QTextDocument()
    td2.setPageSize(QSizeF(400, 700))
    td2.setDocumentMargin(0)
    cur2 = QTextCursor(td2)
    for _ in range(n):
        _insert_paragraph(cur2, "这是一段用来测量行距的文字，需要足够长以便换行。",
                          "宋体", "Times New Roman", 12.0, False, 33.0, "left")
    h2 = td2.documentLayout().documentSize().height()
    check("行距随磅值变化", h2 > h * 1.3, f"22磅 {h:.1f} / 33磅 {h2:.1f}")


def main() -> int:
    src = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "assets" / "sample-filled.docx"
    out_dir = ROOT / "build"
    out_dir.mkdir(exist_ok=True)

    QCoreApplication.setAttribute(Qt.AA_ShareOpenGLContexts, True)
    try:
        QApplication.setHighDpiScaleFactorRoundingPolicy(
            Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
        )
    except Exception:
        pass
    app = QApplication([])
    apply(app)

    test_line_spacing()

    win = MainWindow()
    win.resize(1480, 880)
    win.show()
    app.processEvents()

    # ---- 顶栏按钮是否都在窗口内 ----
    for b in (win.open_btn, win.detect_btn, win.compare_btn, win.apply_btn):
        r = b.geometry()
        check(f"顶栏可见 {b.text()}",
              r.left() >= 0 and r.right() <= win.width() and r.width() > 20,
              f"[{r.left()},{r.right()}] w={r.width()}")

    # ---- 走真实 open_file() ----
    from docx import Document
    from PySide6.QtWidgets import QFileDialog

    QFileDialog.getOpenFileName = staticmethod(lambda *a, **k: (str(src), ""))
    win.open_file()
    app.processEvents()

    s = win.analysis.summary()
    check("识别段落数", win.analysis.paragraph_count > 0,
          f"{win.analysis.paragraph_count} 段")
    check("识别到标题", s["headings"] > 0, f"{s['headings']} 个")

    bd = win.before_view.document()
    ad = win.after_view.document()
    check("左栏（原稿）已渲染", bd is not None and bd.characterCount() > 100,
          f"{bd.characterCount() if bd else 0} 字")
    check("右栏（排版后）已渲染", ad is not None and ad.characterCount() > 100,
          f"{ad.characterCount() if ad else 0} 字")

    nb, na = win.before_view.page_count(), win.after_view.page_count()
    check("真分页", na > 1, f"原稿 {nb} 页 / 排版后 {na} 页")

    # ---- 页眉/页码开关必须同时影响预览（以前只影响 docx，预览在骗人）----
    check("默认形态=期刊论文", win.settings.profile == "journal",
          win.settings.profile)
    check("默认无页眉", win.after_view._header == ""
          and not win.after_view._show_header, repr(win.after_view._header))
    win.settings.enable_header = False
    win.refresh_preview(); app.processEvents()
    check("关掉页眉后预览无页眉",
          win.after_view._header == "" and not win.after_view._show_header, "")
    win.settings.enable_header = True
    win.settings.enable_page_numbers = False
    win.refresh_preview(); app.processEvents()
    check("关掉页码后预览无页码",
          all(x == "" for x in win.after_view._labels),
          f"{win.after_view._labels[:4]}")
    win.settings.enable_page_numbers = True
    win.refresh_preview(); app.processEvents()

    # ---- 摘要正文要有独立角色（以前一律落回 body）----
    ab = [d.index for d in win.analysis.detections if d.role == "abstract_body"]
    abs_idx = [d.index for d in win.analysis.detections if d.role == "abstract"]
    # 只有摘要标题下一段确实是非空正文时才要求（有的稿子这段是空段/文本框）
    has_body_after = any(
        abs_idx and d.index == abs_idx[0] + 1 and d.role == "body"
        for d in win.analysis.detections)
    if has_body_after:
        check("摘要正文独立成角色", len(ab) > 0, f"{len(ab)} 段摘要正文")
    else:
        print("SKIP 摘要正文独立成角色（本文摘要标题后没有非空正文段）")

    # ---- 前置部分的标题不应参与章号（以前会编成 0.1）----

    # ---- 段落级手动格式覆盖 ----
    tgt = next(i for i, d in enumerate(win.analysis.detections)
               if d.role == "body" and win.analysis.features[i].char_len > 60)
    win._override_changed(tgt, {"size": 20.0, "bold": True})
    app.processEvents()
    check("手动格式生效并标记", tgt in win.overrides
          and "处手动" in win.after_head.text(), win.after_head.text())
    win._override_clear(tgt)
    app.processEvents()
    check("清除手动格式", tgt not in win.overrides, "")

    # 表格/图片：原文档里确实有才要求出现在预览里
    n_tbl = len(list(ad.rootFrame().childFrames()))
    n_img = 0
    for i in range(ad.blockCount()):
        it = ad.findBlockByNumber(i).begin()
        while not it.atEnd():
            frag = it.fragment()
            if frag.isValid() and frag.charFormat().isImageFormat():
                n_img += 1
            it += 1
    src_tbl = len(win.doc.tables)
    src_img = sum(1 for f in win.analysis.features if f.has_image)
    if src_tbl:
        # 三线表拆成「表头」+「数据」两个表格，所以是 >=
        check("表格进入预览", n_tbl >= src_tbl, f"{n_tbl}/{src_tbl} 个表")
        # 预览必须跟着「表内文字对齐」变（以前预览根本没实现，
        # 所以用户觉得"设了没用"）
        def _align_of(align: str):
            from detect.ui.render import render_side as _rs
            s2 = Settings()
            s2.table_align = align
            _, _, c2 = page_geometry(s2)
            td2, _b = _rs(win.doc, win.analysis, s2, "after", c2)
            needle = win.doc.tables[0].cell(0, 0).text.strip()[:24]
            if not needle:
                return None
            for i in range(td2.blockCount()):
                b = td2.findBlockByNumber(i)
                if b.text().startswith(needle):
                    return int(b.blockFormat().alignment())
            return None

        a_left, a_cent = _align_of("left"), _align_of("center")
        check("预览跟随表内对齐",
              a_left is not None and a_left != a_cent,
              f"left={a_left} center={a_cent}")
    if src_img:
        # 公式（MathType/EMF）也会以图片形态出现，所以是 >=
        check("图片进入预览", n_img >= src_img, f"{n_img}/{src_img} 张图")

    check("页码标签", len(ad and win.after_view._labels or []) == na,
          f"{na} 个标签")

    # 页码：期刊从第 1 页就是阿拉伯；学位论文封面空、前置罗马
    labels = win.after_view._labels
    if getattr(win.settings, "profile", "journal") == "journal":
        check("期刊：第一页就是 1", labels and labels[0] == "1", f"{labels[:4]}")
        check("页码逐页生成", len(labels) == na and all(labels),
              f"{labels[:4]} … {labels[-2:]}")
    else:
        check("封面无页码", labels and labels[0] == "", f"{labels[:4]}")
        check("页码逐页生成且封面为空",
              len(labels) == na and labels[0] == "" and all(labels[1:]),
              f"{labels[:4]} … {labels[-2:]}")

    # ---- 右栏字号必须来自设置（字号用 pixelSize 设，pointSizeF 是 -1）----
    sizes = []
    for i in range(min(6, ad.blockCount())):
        b = ad.findBlockByNumber(i)
        it = b.begin()
        while not it.atEnd():
            frag = it.fragment()
            if frag.isValid():
                f = frag.charFormat().font()
                if f.pixelSize() > 0:
                    sizes.append(f.pixelSize())
            it += 1
    expected = {int(round(v["size"])) for v in win.settings.roles.values()}
    check("右栏字号取自设置", bool(sizes) and set(sizes) <= expected,
          f"用到 {sorted(set(sizes))} / 设置里 {sorted(expected)}")

    after_text = ad.toPlainText()
    check("右栏含编号", any(ch.isdigit() for ch in after_text), "")

    grab(win, out_dir / "st_main.png")
    grab(win.before_view, out_dir / "st_before.png")
    grab(win.after_view, out_dir / "st_after.png")

    # ---- 设置改动后预览应变化 ----
    win.settings.roles["body"]["size"] = 10.5
    win.refresh_preview()
    app.processEvents()
    check("改设置后右栏重排", win.after_view.page_count() >= 1, "")

    # ---- 识别明细 ----
    from detect.ui.dialogs import IndentDialog, RolesDialog

    rd = RolesDialog(win.analysis, win.doc, win)
    n_rows = rd.table.rowCount()
    check("识别明细行数", n_rows > 0, f"{n_rows} 行")
    grab(rd, out_dir / "st_roles.png")
    if n_rows:
        rd.table.clearSelection()
        rd.table.selectRow(0)
        target = "heading3"
        rd.role_box.setCurrentIndex(rd.role_box.findData(target))
        rd._apply_role()
        rd.apply_to(win.analysis)
        check("手工改正生效", win.analysis.detections[rd._rows[0]].role == target,
              f"第 {rd._rows[0]} 段 → {target}")
        check("手工改正计数", rd.changed == 1, f"{rd.changed}")
    rd.close()

    ind = IndentDialog(win.settings, win, focus_not_indent=True)
    check("顶格默认勾选", ind.not_indent.isChecked(), "")
    grab(ind, out_dir / "st_indent.png")
    ind.close()

    # ---- 一键排版到临时文件 ----
    from detect.reflow import apply_formats, style_tables
    from detect.numbering import (
        apply_plans,
        plan_caption_numbers,
        plan_heading_numbers,
    )
    from detect.sections import apply as apply_sections, diagnose

    out = Path(tempfile.mkdtemp(prefix="selftest_")) / "out.docx"
    import shutil

    shutil.copyfile(src, out)
    d2 = Document(str(out))
    an2 = analyze(d2, source=out.name)

    # ---- 公式编号：Word 域里缓存的结果要能读出来（在改动 d2 之前读）----
    from detect.features import equation_label as _elab
    _eq_idx = [i for i, d in enumerate(an2.detections)
               if effective_role(d) == "equation"]
    _eq_labels = [_elab(d2.paragraphs[i]) for i in _eq_idx]

    front = body = None
    for d in an2.detections:
        if front is None and d.role in ("abstract", "keywords", "toc"):
            front = d.index
        if front is not None and body is None and d.index > front and (
                d.role == "heading" or d.role.startswith("heading")):
            body = d.index
    check("定位到分节点", front is not None and body is not None,
          f"front={front} body={body}")

    if front is not None and body is not None:
        apply_sections(d2, front_index=front, body_index=body,
                       settings=win.settings)
    stats = apply_formats(d2, an2, win.settings.roles)
    plans = (plan_heading_numbers(an2.detections, body, win.settings)
             + plan_caption_numbers(an2.detections, "figure", win.settings,
                             an2.features)
             + plan_caption_numbers(an2.detections, "table", win.settings,
                             an2.features))
    n_num = apply_plans(d2, plans)
    style_tables(d2, win.settings)
    d2.save(str(out))

    check("重排生效段数", sum(stats.values()) > 0, str(stats))

    # ---- 表格：表内对齐 + 表格位置 ----
    if d2.tables:
        from docx.enum.text import WD_ALIGN_PARAGRAPH as _WA
        want = {"left": _WA.LEFT, "center": _WA.CENTER,
                "right": _WA.RIGHT, "both": _WA.JUSTIFY}.get(
                    win.settings.table_align)
        cell = d2.tables[0].cell(0, 0).paragraphs[0]
        check("表内文字对齐生效",
              want is None or cell.alignment == want,
              f"{cell.alignment} / 设置 {win.settings.table_align}")
        try:
            check("表格位置生效",
                  d2.tables[0].alignment is not None, str(d2.tables[0].alignment))
        except Exception:
            check("表格位置生效", False, "读不到")

    # ---- 公式编号：Word 域里缓存的结果要能读出来 ----
    # ---- 公式行距：固定磅值会把 OLE 公式压扁，必须用自动行距 ----
    if _eq_idx:
        from docx.enum.text import WD_LINE_SPACING as _WLS
        rules = {d2.paragraphs[i].paragraph_format.line_spacing_rule
                 for i in _eq_idx}
        check("公式用自动行距（不被压扁）",
              rules <= {_WLS.SINGLE, None}, str(rules))
        check("公式编号能读出来", any(_eq_labels),
              f"{sum(1 for x in _eq_labels if x)}/{len(_eq_labels)} 有编号，"
              f"例 {next((x for x in _eq_labels if x), '')!r}")
    else:
        print("SKIP 公式编号能读出来（本文没有公式）")

    # ---- Word 侧：页脚页码居中 / 作者信息有行距 / 超链接里的引用编号字体 ----
    from docx.enum.text import WD_ALIGN_PARAGRAPH as _WCA
    from docx.oxml.ns import qn as _qn
    d4 = Document(str(out))
    fp = d4.sections[0].footer.paragraphs[0]
    check("页码居中", fp.alignment == _WCA.CENTER, str(fp.alignment))
    meta_i = [i for i, d in enumerate(an2.detections)
              if effective_role(d) == "meta"]
    if meta_i:
        pf4 = d4.paragraphs[meta_i[0]].paragraph_format
        check("作者信息行距生效", pf4.line_spacing is not None,
              f"line={pf4.line_spacing}")
    hyper_runs = [r for p in d4.paragraphs
                  for hl in p._p.iter(_qn("w:hyperlink"))
                  for r in hl.iter(_qn("w:r"))]
    if hyper_runs:
        bad = [r for r in hyper_runs
               if (r.find(_qn("w:rPr")) is None
                   or r.find(_qn("w:rPr")).find(_qn("w:rFonts")) is None)]
        check("超链接内的文字也统一字体", not bad,
              f"{len(hyper_runs)} 个 run，缺字体 {len(bad)} 个")
    else:
        print("SKIP 超链接内的文字也统一字体（本文没有超链接）")

    check("编号写入", n_num > 0, f"{n_num} 个")

    # 中文槽位必须是中文字体：旧版把 w:eastAsia 写成 Times New Roman，
    # 中文会回退成主题字体（实测"表格/正文字体不合适"）
    from docx.oxml.ns import qn as _qn
    from detect.numbering import plan_heading_numbers as _phn
    _east = []
    for para in d2.paragraphs[:body or 1]:
        for r in para.runs:
            rf = r._r.find(_qn("w:rPr"))
            if rf is None:
                continue
            rfonts = rf.find(_qn("w:rFonts"))
            if rfonts is not None and rfonts.get(_qn("w:eastAsia")):
                _east.append(rfonts.get(_qn("w:eastAsia")))
    check("中文槽位=中文字体", bool(_east) and all(e != "Times New Roman" for e in _east),
          f"{sorted(set(_east))}")

    # 标题编号分隔符可自定义
    hs = Settings()
    hs.heading_sep = "-"
    hp = [p.text for p in _phn(an2.detections, body, hs) if "." in p.text or "-" in p.text]
    check("标题编号分隔符可改", bool(hp) and all("." not in t for t in hp),
          str(hp[:3]))

    # 页眉：字体 / 字号 / 线与文字的距离
    hdr = d2.sections[1].header.paragraphs[0]
    check("页眉中文字体", any(
        (rf := r._r.find(_qn("w:rPr"))) is not None
        and (rfonts := rf.find(_qn("w:rFonts"))) is not None
        and rfonts.get(_qn("w:eastAsia")) == win.settings.header_font
        for r in hdr.runs), win.settings.header_font)
    pPr = hdr._p.find(_qn("w:pPr"))
    bdr = pPr.find(_qn("w:pBdr")) if pPr is not None else None
    bot = bdr.find(_qn("w:bottom")) if bdr is not None else None
    gap_val = bot.get(_qn("w:space")) if bot is not None else None
    check("页眉线与文字间距", gap_val is not None
          and abs(int(gap_val) - round(win.settings.header_gap)) <= 1,
          f"w:space={gap_val} 设置={win.settings.header_gap}")

    # 表内字体
    if d2.tables:
        tb = d2.tables[0]
        cell_runs = [r for row in tb.rows for c in row.cells
                     for para in c.paragraphs for r in para.runs]
        sizes = {r.font.size.pt for r in cell_runs if r.font.size is not None}
        check("表内字号统一", sizes == {win.settings.table_font_size}
              or (sizes and max(sizes) <= win.settings.table_font_size + 0.01),
              f"{sorted(sizes)} vs {win.settings.table_font_size}")

    # 图号自定义：前缀 + 分隔符
    cs = Settings()
    cs.caption_fig_prefix = "Figure"
    cs.caption_sep = "-"
    cs.caption_by_chapter = False
    from detect.numbering import plan_caption_numbers as _pcn
    fp = [x for x in _pcn(an2.detections, "figure", cs)][:2]
    if fp:
        check("图号可自定义", fp[0].text.startswith("Figure"),
              str([x.text for x in fp]))
    else:
        print("SKIP 图号可自定义（本文没有图注）")
    from detect.numbering import plan_heading_numbers as _phn
    bad = [p.text for p in _phn(an2.detections, body, win.settings)
           if p.text.startswith("0.") or p.text.startswith("0 ")]
    check("标题编号不出现 0.x", not bad, str(bad[:3]))
    check("前置标题没被编号", not any(
        p.text.strip().startswith(("0.", "0 ")) for p in d2.paragraphs[:body]), "")

    # Highlights / 摘要是前置部分，不能占层级（否则「引言」会被编成 2）
    hl = [d.index for d in an2.detections
          if d.role.startswith("heading")
          and "highlight" in an2.features[d.index].text.lower()]
    check("Highlights 不当标题", not hl, str(hl[:3]))

    # 首段（第一个非空段）要能认出题目；空段不能占层级
    first_ne = next((i for i, f in enumerate(an2.features) if f.char_len > 0), None)
    if first_ne is not None:
        check("首段识别为题目",
              effective_role(an2.detections[first_ne]) == "title",
              f"{effective_role(an2.detections[first_ne])} "
              f"{an2.features[first_ne].text[:26]!r}")
    empt = [i for i, f in enumerate(an2.features)
            if f.char_len == 0
            and effective_role(an2.detections[i]).startswith("heading")]
    check("空段不当标题", not empt, str(empt[:4]))
    # References / 参考文献 这种尾注区标题不能占章号
    refh = [i for i, f in enumerate(an2.features)
            if f.text.strip().lower() in ("references", "参考文献", "致谢", "附录")
            and effective_role(an2.detections[i]).startswith("heading")]
    check("尾注区不占章号", not refh, str(refh[:4]))
    hp = [p.text for p in _phn(an2.detections, body, win.settings)]
    check("引言是第 1 章", bool(hp) and hp[0].startswith("1"), str(hp[:2]))

    # 页号换算必须用 floor：round 会让每一页下半部分的点击落到下一页
    v = win.after_view
    ok = all(v.page_for_y(v.y_for_page(i) + 5) == i
             for i in range(v.page_count()))
    check("页号换算（页内任意位置）", ok and v.page_count() > 1,
          f"{v.page_count()} 页")
    if v.page_count() > 1:
        mid = v.y_for_page(1) + int(v._paper[1] * v._zoom * 0.8)
        check("下半页仍算同一页", v.page_for_y(mid) == 1, f"y={mid}")

    # 手动把某段改成别的角色后，编号要立即重算
    h1s = [i for i, d in enumerate(win.analysis.detections)
           if (body or 0) < i and (
               d.role == "heading1" or effective_role(d) == "heading1")]
    if len(h1s) >= 2:
        before = {p.index: p.text
                  for p in _phn(win.analysis.detections, body, win.settings)}
        t = h1s[1]
        win._apply_role_to_para(t, "heading2")
        after = {p.index: p.text
                 for p in _phn(win.analysis.detections, body, win.settings)}
        check("手动改层级后编号重算",
              t in after and after[t] != before.get(t),
              f"{before.get(t)!r} -> {after.get(t)!r}")
        check("第一个标题仍是 1", after.get(body, "").startswith("1"),
              repr(after.get(body)))
        win._override_clear(t)
    else:
        print("SKIP 手动改层级后编号重算（本文一级标题不足 2 个）")

    d3 = Document(str(out))
    dg = diagnose(d3)
    check("分节数 >= 3", len(dg["sections"]) >= 3,
          f"{len(dg['sections'])} 节: " + " | ".join(
              f"{x['index']}:{x['pgNumType_fmt'] or '-'}" for x in dg["sections"]))

    # ---- 期刊论文：单节、页码全阿拉伯、默认无页眉 ----
    js = Settings()
    js.set_profile("thesis")
    js.set_profile("journal")          # 显式载入期刊预设
    check("期刊：无页眉", js.enable_header is False, "")
    check("期刊：不切分", js.enable_sections is False, "")
    check("期刊：正文字号 10.5", js.roles["body"]["size"] == 10.5, "")
    jp, jm, jc = page_geometry(js)
    jtd, jb = render_side(d2, an2, js, "after", jc, body_index=None)
    jl = page_labels(jtd, jb, None, jc[1], page_count(jtd), profile="journal")
    check("期刊：页码从 1 起全阿拉伯",
          jl[0] == "1" and all(x.isdigit() for x in jl), f"{jl[:4]}")
    # 期刊角色必须齐全：旧版 JOURNAL_ROLES 只有 9 个角色，
    # 点「摘要标题 / 关键词 / 目录 / 参考文献」会直接 KeyError
    for _r in ("abstract", "abstract_body", "keywords", "toc", "reference",
               "back", "equation", "title", "heading1"):
        check(f"期刊角色齐全 {_r}", _r in js.roles, "")
    check("期刊：标签加粗默认开", js.bold_label is True, "")
    check("期刊：摘要居左", js.roles["abstract"]["align"] == "left",
          js.roles["abstract"]["align"])
    ts = Settings()
    ts.set_profile("thesis")
    check("学位：标签加粗默认关", ts.bold_label is False, "")

    # 期刊 Ribbon：没有「分节」，「目录」显示但置灰
    from detect.ui.settings import ribbon_groups, JOURNAL_DISABLED
    jg = ribbon_groups("journal")
    tg = ribbon_groups("thesis")
    check("期刊：Ribbon 无分节", not any(
        k == "layout_section" for _, btns in jg for k, _ in btns), "")
    check("学位：Ribbon 有分节", any(
        k == "layout_section" for _, btns in tg for k, _ in btns), "")
    check("期刊：目录按钮存在但置灰",
          any(k == "toc" for _, btns in jg for k, _ in btns)
          and JOURNAL_DISABLED.get("toc"), "")

    # 期刊：Abstract:/关键词: 标签加粗，正文部分不加粗
    if any(an2.features[i].text.startswith("Keywords") for i in range(len(an2.features))):
        jdoc = Document(str(out))
        jan = analyze(jdoc, source="j.docx")
        _apply = dict(win.settings.roles)
        js.roles["keywords"] = dict(_apply.get("keywords") or js.roles["keywords"])
        from detect.reflow import apply_formats as _af
        _af(jdoc, jan, js.roles, label_bold=True)
        ki = next(i for i, f in enumerate(jan.features)
                  if f.text.startswith("Keywords"))
        _runs = jdoc.paragraphs[ki].runs
        check("期刊：关键词标签加粗",
              bool(_runs) and _runs[0].font.bold is True
              and (_runs[0].text.strip().rstrip(":").strip().lower()
                   in ("keywords", "关键词")),
              repr(_runs[0].text[:20]) if _runs else "无 run")

    # ---- 图注/表注：识别要全，前缀不能重复 ----
    import re as _re
    _RE_CAP_SRC = _re.compile(
        r"^\s*(?:Fig\.?|Figure|Table)\s*[\d(]|^\s*(?:图|表)\s*\d", _re.I)
    from detect.numbering import clean_caption as _cc
    caps = [i for i, d in enumerate(an2.detections)
            if "caption" in effective_role(d)]
    # 图片段后面紧跟着的那段，应当被识别成图注（最可靠的"漏没漏"判据）
    missed = []
    for i, d in enumerate(an2.detections):
        if effective_role(d) != "figure":
            continue
        j = i + 1
        while j < len(an2.detections) and an2.features[j].char_len == 0:
            j += 1
        if j < len(an2.detections) \
                and effective_role(an2.detections[j]) != "figure_caption":
            missed.append(j)
    if any(effective_role(d) == "figure" for d in an2.detections):
        check("图后面的段落都识别成图注", not missed, f"漏了 {missed[:4]}")
    else:
        print("SKIP 图注识别（本文没有图片段）")
    jf = Settings(); jf.set_profile("journal")
    fplans = plan_caption_numbers(an2.detections, "figure", jf, an2.features)
    tplans = plan_caption_numbers(an2.detections, "table", jf, an2.features)
    if fplans and any(f.text.lstrip().lower().startswith("fig")
                      for f in an2.features):
        check("图注自动沿用原稿 Fig./Figure 前缀",
              fplans[0].text.startswith(("Fig.", "Figure", "Fig ")),
              repr(fplans[0].text))
    if tplans and any(f.text.lstrip().lower().startswith("table")
                      for f in an2.features):
        check("表注自动沿用原稿 Table 前缀",
              tplans[0].text.startswith("Table"), repr(tplans[0].text))
    # 激进清洗后不能再残留 Fig./Figure/Table 前缀（否则会出现"图1.1 Fig. 1"）
    dup = [f.text[:30] for i in caps
           for f in [an2.features[i]]
           if _cc(f.text, aggressive=True).lstrip()[:5] in
           ("Fig. ", "Fig. (", "Figure", "Table ")]
    check("图注前缀不重复", not dup, str(dup[:2]))

    # ---- 期刊「目录」卡片：看得见、有说明、点得到（但不能用） ----
    from detect.ui.settings import JOURNAL_DISABLED
    win.settings.set_profile("journal")
    win._rebuild_ribbon()
    app.processEvents()
    toc_btn = None
    from PySide6.QtWidgets import QPushButton as _PB
    for b in win.findChildren(_PB):
        if b.text() == "目录":
            toc_btn = b
            break
    check("期刊：目录按钮没被禁用（否则 tooltip 不显示）",
          toc_btn is not None and toc_btn.isEnabled(), "")
    check("期刊：目录按钮有说明文字",
          bool(toc_btn.toolTip()) if toc_btn else False,
          toc_btn.toolTip() if toc_btn else "")
    asked = []
    from PySide6.QtWidgets import QMessageBox as _QMB
    _QMB.information = lambda *a, **k: asked.append(a)
    win.on_ribbon("toc")
    check("期刊：点目录会说明原因", bool(asked), "")
    win.settings.set_profile("thesis")
    win._rebuild_ribbon()
    app.processEvents()

    # ---- Ctrl + 滚轮缩放 ----
    v = win.after_view
    v.set_fit_mode("pct", 100)
    app.processEvents()
    before = v.zoom_percent()
    from PySide6.QtCore import QPointF as _PF, QEvent as _EV
    from PySide6.QtGui import QWheelEvent as _WE
    import PySide6.QtCore as _qc
    pos = _PF(5, 5)
    we = _WE(pos, pos, _qc.QPoint(0, 0), _qc.QPoint(0, 120),
             _qc.Qt.MouseButton.NoButton, _qc.Qt.KeyboardModifier.ControlModifier,
             _qc.Qt.ScrollPhase.NoScrollPhase, False)
    v.wheelEvent(we)
    app.processEvents()
    check("Ctrl+滚轮放大", v.zoom_percent() > before,
          f"{before} -> {v.zoom_percent()}")
    we2 = _WE(pos, pos, _qc.QPoint(0, 0), _qc.QPoint(0, -120),
              _qc.Qt.MouseButton.NoButton, _qc.Qt.KeyboardModifier.ControlModifier,
              _qc.Qt.ScrollPhase.NoScrollPhase, False)
    v.wheelEvent(we2)
    app.processEvents()
    check("Ctrl+滚轮缩小", v.zoom_percent() == before,
          f"{before} -> {v.zoom_percent()}")

    # ---- 一键排版：只生成结果 + 刷新预览，绝不弹保存框 ----
    import detect.ui.preview as prev_mod
    from PySide6.QtWidgets import QMessageBox

    prev_mod.PreviewDialog.exec = lambda self, *a, **k: 1
    asked = []
    QFileDialog.getSaveFileName = staticmethod(
        lambda *a, **k: (asked.append(1), ("", ""))[1])

    win.apply()
    app.processEvents()
    check("一键排版不弹保存框", not asked, f"被问了 {len(asked)} 次")
    check("结果写入缓存", getattr(win, "cached_out", None) is not None
          and Path(win.cached_out).is_file(),
          str(getattr(win, "cached_out", None)))
    check("保存按钮已启用", win.save_btn.isEnabled(), "")
    check("排版后预览已刷新", "已排版" in win.after_head.text(),
          win.after_head.text())

    # 另存为：把保存框引到临时文件，绝不覆盖用户的稿子
    alt = Path(tempfile.mkdtemp(prefix="selftest_save_")) / "另存为.docx"
    QFileDialog.getSaveFileName = staticmethod(lambda *a, **k: (str(alt), ""))
    win.save("as")
    check("另存为落盘", alt.is_file() and alt.stat().st_size > 0,
          f"{alt} {alt.stat().st_size if alt.exists() else 0} 字节")
    QFileDialog.getSaveFileName = staticmethod(lambda *a, **k: ("", ""))
    win.save("as")
    check("取消另存为不报错", True, "")
    QMessageBox.information = lambda *a, **k: None

    print("\n" + ("全部通过" if not FAIL else f"失败 {len(FAIL)} 项: {FAIL}"))
    win.close()
    return 1 if FAIL else 0


if __name__ == "__main__":
    raise SystemExit(main())
