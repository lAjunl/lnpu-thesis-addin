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
# 隔离 APPDATA：避免受用户本地“上次方案”影响，也绝不把测试方案写进用户目录
os.environ["APPDATA"] = tempfile.mkdtemp(prefix="selftest_appdata_")

from PySide6.QtCore import Qt, QSizeF, QCoreApplication, QMimeData, QUrl, QPointF
from PySide6.QtGui import QTextCursor, QTextDocument, QDropEvent
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
    from detect.ui.app import App
    app = App([])
    apply(app)

    test_line_spacing()

    win = MainWindow()
    app.main_window = win          # 拖拽拦截挂在 App 上，要指过来
    win.resize(1480, 880)
    win.show()
    app.processEvents()

    # ---- 顶栏按钮是否都在窗口内，且 split 按钮两半等高 ----
    for b in (win.open_btn, win.detect_btn, win.compare_btn, win.apply_btn):
        r = b.geometry()
        check(f"顶栏可见 {b.text()}",
              r.left() >= 0 and r.right() <= win.width() and r.width() > 20,
              f"[{r.left()},{r.right()}] w={r.width()}")
    check("识别 split 按钮等高",
          win.detect_btn.height() == win.detect_arrow.height(),
          f"{win.detect_btn.height()} vs {win.detect_arrow.height()}")
    check("标题编号默认关闭", not win.settings.enable_number_heading, "")

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

    # ---- 拖拽打开：只认 .docx ----
    class _DropEvt:
        def __init__(self, urls):
            self._urls = urls

        def mimeData(self):
            from PySide6.QtCore import QMimeData, QUrl
            m = QMimeData()
            m.setUrls([QUrl.fromLocalFile(u) for u in self._urls])
            return m

    p1 = win._drop_docx(_DropEvt([str(Path(src))]))
    check("拖拽识别 .docx", p1 is not None and p1.suffix.lower() == ".docx",
          str(p1))
    check("拖拽拒绝非 docx", win._drop_docx(_DropEvt(["C:/x/a.txt"])) is None, "")
    check("拖拽拒绝空内容", win._drop_docx(_DropEvt([])) is None, "")
    check("拖放已开启", win.acceptDrops(), "")

    # 端到端：把 Drop 事件发给子控件（真实拖拽就是投递给光标下的子控件），
    # 由 App.notify() 拦截后打开文档
    def _send_drop(target, path):
        m = QMimeData()
        m.setUrls([QUrl.fromLocalFile(str(path))] if path else [])
        ev = QDropEvent(QPointF(400, 400), Qt.CopyAction, m,
                        Qt.LeftButton, Qt.NoModifier)
        QApplication.sendEvent(target, ev)
        app.processEvents()

    _send_drop(win.after_view, Path(src))
    check("拖拽打开文档（投递给子控件）",
          win.path is not None and win.path.suffix.lower() == ".docx",
          str(win.path))
    _send_drop(win.after_view, Path("C:/x/not_a_docx.txt"))
    check("拖非 docx 不换文档",
          win.path is not None and win.path.suffix.lower() == ".docx",
          str(win.path))

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

    # 标题编号分隔符可自定义（这里要的是"全部替换"模式，否则已有编号会被保留）
    hs = Settings()
    hs.heading_sep = "-"
    hs.number_keep_existing = False
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

    def _plans_now():
        return {p.index: p.text
                for p in _phn(an2.detections, body, win.settings)}

    def _eff(idx, plans):
        """这段最终的编号：有 plan 用 plan，否则保留原稿已有的编号。"""
        return plans.get(idx) or (an2.detections[idx].number or "")

    _plans0 = _plans_now()
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
    first_h = next((d.index for d in an2.detections
                    if d.index >= (body or 0)
                    and effective_role(d).startswith("heading")), None)
    if first_h is not None:
        eff = _eff(first_h, _plans0)
        check("引言是第 1 章", eff.startswith("1"), repr(eff))
    else:
        print("SKIP 引言是第 1 章（没找到正文标题）")

    # 页号换算必须用 floor：round 会让每一页下半部分的点击落到下一页
    v = win.after_view
    ok = all(v.page_for_y(v.y_for_page(i) + 5) == i
             for i in range(v.page_count()))
    check("页号换算（页内任意位置）", ok and v.page_count() > 1,
          f"{v.page_count()} 页")
    if v.page_count() > 1:
        mid = v.y_for_page(1) + int(v._paper[1] * v._zoom * 0.8)
        check("下半页仍算同一页", v.page_for_y(mid) == 1, f"y={mid}")

    # 编号模式：默认「保留原稿已有编号，只补缺的」；可切成「全部替换」
    h1s = [i for i, d in enumerate(win.analysis.detections)
           if (body or 0) < i and (
               d.role == "heading1" or effective_role(d) == "heading1")]
    if len(h1s) >= 2:
        keep = bool(getattr(win.settings, "number_keep_existing", True))
        check("编号默认保留原稿已有编号", keep, str(keep))

        def _plans():
            return {p.index: p.text
                    for p in _phn(win.analysis.detections, body, win.settings)}

        # 已有编号的段落：保留模式下不该被改
        t = h1s[1]
        has_num = bool(win.analysis.detections[t].number)
        before = _plans()
        win._apply_role_to_para(t, "heading2")
        after = _plans()
        if has_num:
            check("保留模式：已有编号不动",
                  after.get(t, before.get(t)) == before.get(t),
                  f"{before.get(t)!r} -> {after.get(t)!r}")
        else:
            check("保留模式：缺编号的补上", t in after, str(after.get(t)))
        win._override_clear(t)

        # 切成替换模式：同一处必须有新号
        win.settings.number_keep_existing = False
        b2 = _plans()
        win._apply_role_to_para(t, "heading1")
        a2 = _plans()
        check("替换模式：编号被重排",
              t in a2 and a2[t] != b2.get(t),
              f"{b2.get(t)!r} -> {a2.get(t)!r}")
        win.settings.number_keep_existing = True
        win._override_clear(t)
        # 第一个标题：保留模式用原稿已有的号，替换模式用新计划的号
        _p = _plans_now()
        eff = _eff(body, _p) if body is not None else ""
        check("第一个标题仍是 1", eff.startswith("1"), repr(eff))
    else:
        print("SKIP 编号保留/替换（本文一级标题不足 2 个）")

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

    # ---- 缩进 / 公式制表位要真正写进 docx（以前只在预览里改了）----
    from docx.oxml.ns import qn as _qn
    from detect.classify import effective_role as _er

    out_doc = Document(str(win.cached_out))
    out_paras = out_doc.paragraphs
    body_idx = [d.index for d in win.analysis.detections if _er(d) == "body"]
    n_ind = 0
    for i in body_idx:
        pPr = out_paras[i]._p.pPr
        ind = pPr.find(_qn("w:ind")) if pPr is not None else None
        if ind is not None and ind.get(_qn("w:firstLineChars")):
            n_ind += 1
    if body_idx:
        check("正文首行缩进写入 docx", n_ind == len(body_idx),
              f"{n_ind}/{len(body_idx)} 段带 firstLineChars")

    eq_idx = [d.index for d in win.analysis.detections if _er(d) == "equation"]
    if eq_idx:
        tabbed = 0
        for i in eq_idx:
            pPr = out_paras[i]._p.pPr
            tabs = pPr.find(_qn("w:tabs")) if pPr is not None else None
            if tabs is not None:
                tabbed += 1
        check("公式编号有制表位（右对齐）", tabbed > 0, f"{tabbed}/{len(eq_idx)} 段")

    inline_math = [i for i, fd in enumerate(win.analysis.features)
                   if (fd.has_math or fd.has_ole) and fd.text.strip()]
    if inline_math:
        role_by_idx = {d.index: _er(d) for d in win.analysis.detections}
        ok = all(role_by_idx.get(i) != "equation" for i in inline_math)
        check("行内公式段落不算公式段", ok, f"{len(inline_math)} 段")

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

    # ---- 格式方案库：内置两套 + 增删 + 套用 ----
    from detect import schemes

    sc = schemes.Scheme(name="自测方案", kind="journal",
                        settings=win.settings.scheme_payload(),
                        source="selftest")
    schemes.save(sc)
    check("方案保存并能在列表中找到",
          schemes.find("自测方案", "journal") is not None,
          str([s.name for s in schemes.load_user_schemes()]))
    check("内置方案只读",
          schemes.builtin("journal").builtin
          and not schemes.delete(schemes.builtin("journal")), "")
    check("默认两套方案都在",
          bool(schemes.builtin("journal").name and schemes.builtin("thesis").name),
          "")
    win._apply_scheme(schemes.find("自测方案", "journal"))
    check("套用方案后状态栏跟着变", "自测方案" in win.scheme_btn.text(),
          win.scheme_btn.text())
    win._apply_scheme(schemes.builtin("journal"))
    schemes.delete(sc)
    check("方案删除后找不到", schemes.find("自测方案", "journal") is None, "")

    # ---- 格式识别：样式/默认值派生的格式能读出来 ----
    from detect.ui.scheme_dialogs import DetectDialog

    dlg = DetectDialog(win.doc, win.analysis, base_settings=win.settings, parent=win)
    dlg._detect()
    check("识别核对表有内容", dlg.table.rowCount() > 0, f"{dlg.table.rowCount()} 行")
    sizes = [f.size_pt for f in win.analysis.features]
    check("字号没有读空的段落（docDefaults 修复）",
          all(s is not None for s in sizes),
          f"{sum(1 for s in sizes if s is None)} 段为空")
    dlg.close()

    # ---- 本轮回归：手动加粗不乱行距 / 三线表两粗一细 / 网格表预览有线 / 手动保存方案 ----
    from PySide6.QtCore import QRectF as _QRF
    from PySide6.QtGui import QColor as _QC, QImage as _QI, QPainter as _QP
    from detect.ui.dialogs import ParaFormatDialog
    from detect.ui.scheme_dialogs import SaveSchemeDialog
    from detect.reflow import style_tables as _stbl

    pdlg = ParaFormatDialog(0, {"font": "宋体", "ascii": "Times New Roman",
                                "size": 12.0, "bold": False,
                                "line": 1.5, "line_rule": "multiple",
                                "align": "both", "color": ""}, win)
    pdlg.set_value({"font": "宋体", "ascii": "Times New Roman",
                    "size": 12.0, "bold": False,
                    "line": 1.5, "line_rule": "multiple",
                    "align": "both", "color": ""})
    pdlg.bold.setChecked(True)
    pv = pdlg.value()
    check("手动加粗保持原行距方式（不坍成 1.5 磅固定值）",
          pv["bold"] is True and pv["line_rule"] == "multiple", str(pv))
    pdlg.close()

    tdoc = Document()
    tdoc.add_paragraph("表1 自测表")
    tb0 = tdoc.add_table(rows=3, cols=3)
    for r in range(3):
        for c in range(3):
            tb0.cell(r, c).text = f"R{r}C{c}"
    ts = Settings()
    _stbl(tdoc, ts)
    tbl_b = tdoc.tables[0]._tbl.tblPr.find(_qn("w:tblBorders"))
    tc_b = tdoc.tables[0].rows[0].cells[0]._tc.find(_qn("w:tcPr")).find(_qn("w:tcBorders")).find(_qn("w:bottom"))
    check("三线表两粗一细（Word 顶/底 1.5 磅=12，栏目线 0.75 磅=6）",
          tbl_b.find(_qn("w:top")).get(_qn("w:sz")) == "12"
          and tbl_b.find(_qn("w:bottom")).get(_qn("w:sz")) == "12"
          and tc_b.get(_qn("w:sz")) == "6",
          f"top={tbl_b.find(_qn('w:top')).get(_qn('w:sz'))} head={tc_b.get(_qn('w:sz'))}")

    tan = analyze(tdoc)
    for st_name in ("three_line", "grid"):
        ts.table_style = st_name
        td_tbl, _ = render_side(tdoc, tan, ts, "after", (415.0, 680.0))
        im_tbl = _QI(415, 200, _QI.Format.Format_ARGB32)
        im_tbl.fill(_QC("#ffffff"))
        pt_tbl = _QP(im_tbl)
        pt_tbl.setRenderHint(_QP.RenderHint.SmoothPixmapTransform, True)
        td_tbl.drawContents(pt_tbl, _QRF(0, 0, 415, 200))
        pt_tbl.end()
        hlines = 0
        in_l = False
        for y in range(im_tbl.height()):
            dark = sum(1 for x in range(im_tbl.width())
                       if im_tbl.pixelColor(x, y).lightness() < 235)
            if dark > im_tbl.width() * 0.5:
                if not in_l:
                    hlines += 1
                    in_l = True
            else:
                in_l = False
        check(f"表格预览有线（{st_name}）", hlines >= 3, f"{hlines} 条横线")

    sdlg = SaveSchemeDialog(schemes.builtin("journal"), win.settings, win)
    sdlg.name_edit.setText("手动保存自测方案")
    sdlg.kind_cb.setCurrentIndex(sdlg.kind_cb.findData("thesis"))
    sdlg._save()
    check("手动设置可直接存进学位论文方案",
          sdlg.scheme is not None
          and sdlg.scheme.kind == "thesis"
          and sdlg.scheme.settings.get("profile") == "thesis"
          and schemes.find("手动保存自测方案", "thesis") is not None,
          f"{sdlg.scheme.kind if sdlg.scheme else None}")
    if sdlg.scheme is not None:
        check("方案可在学位/期刊之间移动",
              schemes.move_kind(sdlg.scheme, "journal")
              and schemes.find("手动保存自测方案", "journal") is not None
              and schemes.find("手动保存自测方案", "thesis") is None, "")
        schemes.delete(sdlg.scheme)
    sdlg.close()

    # 格式识别弹窗：来源单选与类型单选互不干扰，且能保存到学位论文
    ddlg = DetectDialog(win.doc, win.analysis, base_settings=win.settings, parent=win)
    ddlg.kind_thesis.setChecked(True)
    ddlg.tpl_radio.setChecked(True)
    check("格式识别：选模板来源不会把「学位」单选框取消",
          ddlg.kind_thesis.isChecked() and ddlg.tpl_radio.isChecked(),
          f"thesis={ddlg.kind_thesis.isChecked()} tpl={ddlg.tpl_radio.isChecked()}")
    ddlg.cur_radio.setChecked(True)
    ddlg._detect()
    ddlg.name_edit.setText("识别学位自测")
    ddlg._save_scheme()
    check("格式识别：可正常保存到学位论文",
          ddlg.scheme is not None
          and ddlg.scheme.kind == "thesis"
          and ddlg.scheme.settings.get("profile") == "thesis"
          and schemes.find("识别学位自测", "thesis") is not None,
          f"{ddlg.scheme.kind if ddlg.scheme else None}")
    if ddlg.scheme is not None:
        schemes.delete(ddlg.scheme)
    ddlg.close()

    # ---- 表内字体识别、核对表编辑、表头/表体独立加粗 ----
    from detect.recognize import inspect_table_format as _itf, recognize_from_doc as _rfd, settings_from_report as _sfr
    from detect.reflow import style_tables as _stbl
    from detect.ui.dialogs import TableDialog as _TableDialog
    from docx.enum.text import WD_ALIGN_PARAGRAPH as _WDA

    tdoc = Document()
    tdoc.add_paragraph("1 引言")
    tbl = tdoc.add_table(rows=2, cols=2)
    for r_i, vals in enumerate((("参数", "数值"), ("压力", "10.5 MPa"))):
        for c_i, txt in enumerate(vals):
            p = tbl.cell(r_i, c_i).paragraphs[0]
            p.alignment = _WDA.CENTER
            run = p.add_run(txt)
            from detect.reflow import set_run_font as _srf
            _srf(run, "仿宋", "Arial", 9.5, r_i == 0)

    tan = analyze(tdoc, source="tbl_test.docx")
    trep = {}
    ts_rec = _rfd(tdoc, tan, report=trep)
    ti = trep.get("table") or {}
    check("表内字体识别：中文/西文/字号/对齐",
          ti.get("font") == "仿宋" and ti.get("ascii") == "Arial"
          and ti.get("size") == 9.5 and ti.get("align") == "center",
          str(ti))
    check("表内加粗识别：表头与表体拆开",
          ti.get("header_bold") is True and ti.get("bold") is False
          and ts_rec.table_header_bold is True and ts_rec.table_bold is False,
          f"hb={ti.get('header_bold')} bb={ti.get('bold')}")

    # 无表格文档：report["table"] 必须留 None，不拿默认值冲掉 base
    nodoc = Document()
    nodoc.add_paragraph("正文段落测试内容。")
    noan = analyze(nodoc, source="no_tbl.docx")
    norep = {}
    _rfd(nodoc, noan, report=norep)
    base_keep = Settings()
    base_keep.table_font = "黑体"
    base_keep.table_header_bold = False
    merged = _sfr(base_keep, norep)
    check("无表格文档不拿默认值冒充表内识别",
          norep["table"]["font"] is None and norep["table"]["header_bold"] is None
          and merged.table_font == "黑体" and merged.table_header_bold is False,
          str(norep["table"]))

    # 老版本方案 JSON 兼容：只有 table_bold 时 table_header_bold 自动沿用
    old_s = Settings.from_json({"table_bold": True})
    check("老方案兼容：缺 table_header_bold 时继承 table_bold",
          old_s.table_bold is True and old_s.table_header_bold is True, "")

    # Word 导出：表头与表体加粗分别生效（且关闭时显式写 val="0"）
    s_hb = Settings()
    s_hb.table_font, s_hb.table_ascii, s_hb.table_font_size = "黑体", "Times New Roman", 10.0
    s_hb.table_header_bold, s_hb.table_bold = True, False
    _stbl(tdoc, s_hb)
    r0_b = tdoc.tables[0].rows[0].cells[0].paragraphs[0].runs[0].font.bold
    r1_b = tdoc.tables[0].rows[1].cells[0].paragraphs[0].runs[0].font.bold
    check("Word 表格导出：表头加粗、表体不加粗",
          r0_b is True and r1_b is False, f"r0={r0_b} r1={r1_b}")
    s_hb.table_header_bold, s_hb.table_bold = False, True
    _stbl(tdoc, s_hb)
    r0_b2 = tdoc.tables[0].rows[0].cells[0].paragraphs[0].runs[0].font.bold
    r1_b2 = tdoc.tables[0].rows[1].cells[0].paragraphs[0].runs[0].font.bold
    check("Word 表格导出：切换为表头不加粗、表体加粗",
          r0_b2 is False and r1_b2 is True, f"r0={r0_b2} r1={r1_b2}")

    # TableDialog 双复选框读写与示意图同步
    td_ui = _TableDialog(s_hb, win)
    check("TableDialog 载入表头/表体加粗分离状态",
          td_ui.header_bold.isChecked() is False and td_ui.bold.isChecked() is True
          and td_ui.demo.data.get("header_bold") is False
          and td_ui.demo.data.get("bold") is True, str(td_ui.demo.data))
    td_ui.header_bold.setChecked(True)
    td_ui.bold.setChecked(False)
    td_ui.apply_to(s_hb)
    check("TableDialog 写回表头/表体加粗",
          s_hb.table_header_bold is True and s_hb.table_bold is False, "")
    td_ui.close()

    # DetectDialog 核对表包含「表格内文字」行，且修改后能覆盖设置
    ddlg2 = DetectDialog(tdoc, tan, base_settings=Settings(), parent=win)
    ddlg2._detect()
    last_row = ddlg2.table.rowCount() - 1
    row_name = ddlg2.table.item(last_row, 0).text() if last_row >= 0 else ""
    check("格式识别核对表含「表格内文字」行",
          row_name == "表格内文字" and ddlg2._tbl_cells is not None,
          repr(row_name))
    if ddlg2._tbl_cells is not None:
        ddlg2._tbl_cells["font"].setCurrentText("楷体")
        ddlg2._tbl_cells["ascii"].setCurrentText("Calibri")
        ddlg2._tbl_cells["size"].setValue(9.0)
        ddlg2._tbl_cells["header_bold"].setChecked(True)
        ddlg2._tbl_cells["bold"].setChecked(False)
        ddlg2._tbl_cells["align"].setCurrentIndex(
            ddlg2._tbl_cells["align"].findData("right"))
        ddlg2._apply_now()
        s_out = ddlg2.settings
        check("格式识别核对表修改表内字体与双加粗生效",
              s_out is not None and s_out.table_font == "楷体"
              and s_out.table_ascii == "Calibri" and s_out.table_font_size == 9.0
              and s_out.table_header_bold is True and s_out.table_bold is False
              and s_out.table_align == "right",
              "" if s_out is None else f"{s_out.table_font}/{s_out.table_ascii}/{s_out.table_font_size}/hb={s_out.table_header_bold}/bb={s_out.table_bold}/{s_out.table_align}")
    ddlg2.close()

    # ---- 页眉识别防误判与彻底清理 ----
    from detect.recognize import inspect_header_format as _ihf
    from detect.sections import configure as _cfg_sec, plans_from_settings as _pfs

    # 1) 1 页文档开启「首页不同」，first_page_header 空、sec.header 残留旧校名 → 不得误判为有页眉
    gh_doc = Document()
    gh_doc.sections[0].different_first_page_header_footer = True
    gh_doc.sections[0].header.paragraphs[0].add_run("某模板残留校名页眉")
    gh_doc.add_paragraph("1 引言")
    gh_doc.add_paragraph("简短单页正文，肉眼在 Word 里完全看不到页眉。")
    gh_info = _ihf(gh_doc)
    check("页眉识别：单页首页不同隐藏的残留校名不误判",
          gh_info["enabled"] is False and gh_info["text"] == "", str(gh_info))

    # 2) 多节文档：封面节残留校名、正文节无页眉 → 按正文体量投票判定无页眉
    cv_doc = Document()
    cv_doc.sections[0].header.paragraphs[0].add_run("某模板残留封面校名")
    cv_doc.add_paragraph("封面标题")
    sec_body = cv_doc.add_section()
    sec_body.header.is_linked_to_previous = False
    cv_doc.add_paragraph("1 正文标题")
    cv_doc.add_paragraph("正文内容非常长" * 80)
    cv_info = _ihf(cv_doc)
    check("页眉识别：仅封面残留页眉但正文无页眉时不误判",
          cv_info["enabled"] is False and cv_info["text"] == "", str(cv_info))

    # 3) 纯白字/隐藏字页眉 → 不误判
    hid_doc = Document()
    hr = hid_doc.sections[0].header.paragraphs[0].add_run("隐藏校名页眉")
    from docx.oxml import OxmlElement as _OE
    _rpr = hr._r.get_or_add_rPr()
    _col = _OE("w:color")
    _col.set(_qn("w:val"), "FFFFFF")
    _rpr.append(_col)
    hid_doc.add_paragraph("正文内容" * 30)
    hid_info = _ihf(hid_doc)
    check("页眉识别：纯白隐藏字不误判为页眉",
          hid_info["enabled"] is False and hid_info["text"] == "", str(hid_info))

    # 4) 多页文档开启「首页不同」但第 2 页起确有页眉 → 准确识别文字与字体
    real_hdoc = Document()
    real_hdoc.sections[0].different_first_page_header_footer = True
    rhr = real_hdoc.sections[0].header.paragraphs[0].add_run("真实学术论文页眉")
    _srf(rhr, "楷体", "Times New Roman", 9.5, False)
    real_hdoc.add_paragraph("首页标题")
    real_hdoc.add_page_break()
    real_hdoc.add_paragraph("第二页正文内容" * 40)
    rh_info = _ihf(real_hdoc)
    check("页眉识别：多页文档真实页眉与字号准确识别",
          rh_info["enabled"] is True and rh_info["text"] == "真实学术论文页眉"
          and rh_info["font"] == "楷体" and rh_info["size"] == 9.5, str(rh_info))

    # 5) base 开启了页眉（如学位方案），识别无页眉文档覆盖后必须关闭 enable_header
    base_with_hdr = Settings()
    base_with_hdr.enable_header = True
    base_with_hdr.header_text = "旧方案遗留校名页眉"
    gh_an = analyze(gh_doc, source="无页眉文档.docx")
    gh_rep: dict = {}
    _rfd(gh_doc, gh_an, report=gh_rep)
    merged_nohdr = _sfr(base_with_hdr, gh_rep)
    check("格式识别覆盖：无页眉文档关闭旧方案的 enable_header",
          merged_nohdr.enable_header is False and merged_nohdr.header_text == "",
          f"en={merged_nohdr.enable_header} txt={merged_nohdr.header_text!r}")

    # 6) DetectDialog 页眉核对条显示与开关回写
    ddlg3 = DetectDialog(gh_doc, gh_an, base_settings=base_with_hdr, parent=win)
    ddlg3._detect()
    check("格式识别弹窗：无页眉文档自动取消勾选启用页眉",
          ddlg3.hdr_enable.isChecked() is False and ddlg3.hdr_text.text() == "",
          f"chk={ddlg3.hdr_enable.isChecked()} txt={ddlg3.hdr_text.text()!r}")
    ddlg3._apply_now()
    check("格式识别弹窗：无页眉文档覆盖设置后关闭页眉",
          ddlg3.settings is not None and ddlg3.settings.enable_header is False
          and ddlg3.settings.header_text == "", "")
    ddlg3.close()

    # 7) Word 导出：关闭页眉时彻底清除残留页眉文字、首页页眉与横线
    clr_doc = Document()
    clr_doc.sections[0].different_first_page_header_footer = True
    clr_doc.sections[0].header.paragraphs[0].add_run("残留默认页眉")
    clr_doc.sections[0].first_page_header.paragraphs[0].add_run("残留首页页眉")
    ppr0 = clr_doc.sections[0].header.paragraphs[0]._p.get_or_add_pPr()
    pbdr = _OE("w:pBdr")
    bot = _OE("w:bottom")
    bot.set(_qn("w:val"), "single")
    bot.set(_qn("w:sz"), "6")
    pbdr.append(bot)
    ppr0.append(pbdr)
    clr_doc.add_paragraph("正文内容")
    s_nohdr = Settings()
    s_nohdr.profile = "journal"
    s_nohdr.enable_header = False
    s_nohdr.header_text = ""
    _cfg_sec(clr_doc, _pfs(s_nohdr), s_nohdr)
    rem_def = "".join(p.text for p in clr_doc.sections[0].header.paragraphs).strip()
    rem_fst = "".join(p.text for p in clr_doc.sections[0].first_page_header.paragraphs).strip()
    rem_bdr = clr_doc.sections[0].header.paragraphs[0]._p.find(".//" + _qn("w:pBdr"))
    check("Word 导出：无页眉时彻底清除默认/首页页眉与残留横线",
          rem_def == "" and rem_fst == "" and rem_bdr is None,
          f"def={rem_def!r} fst={rem_fst!r} bdr={rem_bdr}")

    # ---- 表格底纹清除、单元格首行缩进清零、原稿侧预览缓存与幂等 apply ----
    shd_doc = Document()
    shd_tbl = shd_doc.add_table(rows=2, cols=2)
    # 模拟蓝色表头底纹（tcPr/w:shd + tblStyle）和单元格首行缩进（firstLine=422）
    _ts = _OE("w:tblStyle")
    _ts.set(_qn("w:val"), "GridTable4-Accent1")
    shd_tbl._tbl.tblPr.append(_ts)
    for c_i, txt in enumerate(("序号", "内容")):
        c = shd_tbl.rows[0].cells[c_i]
        tcPr = c._tc.get_or_add_tcPr()
        _sh = _OE("w:shd")
        _sh.set(_qn("w:val"), "clear")
        _sh.set(_qn("w:fill"), "B5C7EA")
        tcPr.append(_sh)
        p = c.paragraphs[0]
        p.add_run(txt)
        _ind = p._p.get_or_add_pPr().get_or_add_ind()
        _ind.set(_qn("w:firstLine"), "422")
    shd_tbl.rows[1].cells[0].paragraphs[0].add_run("1")
    shd_tbl.rows[1].cells[0].paragraphs[0]._p.get_or_add_pPr().get_or_add_ind().set(_qn("w:firstLine"), "420")
    shd_tbl.rows[1].cells[1].paragraphs[0].add_run("测试正文")

    s_shd = Settings()
    s_shd.table_style = "three_line"
    s_shd.table_clear_shading = True
    _stbl(shd_doc, s_shd)
    rem_shd = list(shd_tbl._tbl.iter(_qn("w:shd")))
    rem_ts = list(shd_tbl._tbl.iter(_qn("w:tblStyle")))
    ind_r0 = shd_tbl.rows[0].cells[0].paragraphs[0]._p.find(_qn("w:pPr")).find(_qn("w:ind"))
    ind_r1 = shd_tbl.rows[1].cells[0].paragraphs[0]._p.find(_qn("w:pPr")).find(_qn("w:ind"))
    check("Word 表格导出：默认清除表头蓝色底纹与内置着色表格样式",
          len(rem_shd) == 0 and len(rem_ts) == 0,
          f"shd={len(rem_shd)} tblStyle={len(rem_ts)}")
    check("Word 表格导出：表内单元格首行缩进显式清零（防窄列序号断行）",
          ind_r0 is not None and ind_r0.get(_qn("w:firstLine")) == "0"
          and ind_r0.get(_qn("w:firstLineChars")) == "0"
          and ind_r1 is not None and ind_r1.get(_qn("w:firstLine")) == "0",
          f"r0={dict(ind_r0.attrib) if ind_r0 is not None else None}")

    # 取消勾选清除底纹时保留原底纹
    keep_doc = Document()
    keep_tbl = keep_doc.add_table(rows=1, cols=1)
    _sh2 = _OE("w:shd")
    _sh2.set(_qn("w:val"), "clear")
    _sh2.set(_qn("w:fill"), "B5C7EA")
    keep_tbl.rows[0].cells[0]._tc.get_or_add_tcPr().append(_sh2)
    keep_tbl.rows[0].cells[0].paragraphs[0].add_run("表头")
    s_keep = Settings()
    s_keep.table_clear_shading = False
    _stbl(keep_doc, s_keep)
    check("Word 表格导出：关闭清除底纹选项时保留原表底纹",
          len(list(keep_tbl._tbl.iter(_qn("w:shd")))) == 1, "")

    # TableDialog 载入与写回 table_clear_shading
    s_td = Settings()
    s_td.table_clear_shading = False
    tdlg_shd = _TableDialog(s_td, win)
    check("TableDialog 载入清除底纹开关",
          tdlg_shd.clear_shading.isChecked() is False, "")
    tdlg_shd.clear_shading.setChecked(True)
    tdlg_shd.apply_to(s_td)
    check("TableDialog 写回清除底纹开关",
          s_td.table_clear_shading is True, "")
    tdlg_shd.close()

    # 原稿侧预览缓存命中：调非页边距设置时直接复用同一份 before QTextDocument
    win.refresh_preview()
    bef_doc_1 = win.before_view._doc
    win.settings.table_header_bold = not win.settings.table_header_bold
    win.refresh_preview()
    bef_doc_2 = win.before_view._doc
    check("长文档优化：调非页边距设置时复用原稿侧 QTextDocument 缓存",
          bef_doc_1 is not None and bef_doc_1 is bef_doc_2, "")

    from detect import __repo_url__ as _repo_url
    from detect.ui.dialogs import AboutDialog as _AboutDialog
    adlg = _AboutDialog(win)
    check("关于弹窗显示可点击 GitHub 地址",
          _repo_url in adlg.repo_label.text() and adlg.repo_label.openExternalLinks(),
          adlg.repo_label.text())
    adlg.close()

    # ---- 拖放浮层：只盖住下方预览区，不染蓝顶部卡片 ----
    win._drop_hl(True, "demo.docx")
    app.processEvents()
    check("拖放悬停：浮层覆盖整个下方预览区（且不挤占分栏）",
          win._drop_overlay.isVisible()
          and win.preview_split.count() == 2
          and win._drop_overlay.geometry() == win.preview_host.rect()
          and win._drop_overlay.width() == win.width(),
          f"vis={win._drop_overlay.isVisible()} cols={win.preview_split.count()} geo={win._drop_overlay.geometry()}")
    check("拖放悬停：不改 centralWidget 样式（顶部卡片不变蓝）",
          win.centralWidget().styleSheet() == "",
          repr(win.centralWidget().styleSheet()))
    win._drop_hl(False)
    app.processEvents()
    check("拖放离开：隐藏预览区浮层", not win._drop_overlay.isVisible(), "")

    # ---- 预览空白行与分页符/分节符渲染（含文末空段裁剪与 val=0 防误判） ----
    blk_doc = Document()
    blk_doc.add_paragraph("第一行")
    for _ in range(4):
        blk_doc.add_paragraph("")
    p_after_blank = blk_doc.add_paragraph("空行之后的第二行")
    # 模拟 Word 常见的 w:pageBreakBefore w:val="0"（不得误分页）
    _pbb0 = _OE("w:pageBreakBefore")
    _pbb0.set(_qn("w:val"), "0")
    p_after_blank._p.get_or_add_pPr().append(_pbb0)
    # 显式手动分页符
    p_br = blk_doc.add_paragraph("")
    _r_br = p_br.add_run()
    _br_el = _OE("w:br")
    _br_el.set(_qn("w:type"), "page")
    _r_br._r.append(_br_el)
    blk_doc.add_paragraph("分页后的第二页正文")
    # 文末 10 个连续空段：应被裁掉，不产生空白尾页
    for _ in range(10):
        blk_doc.add_paragraph("")

    blk_ana = analyze(blk_doc)
    blk_set = Settings()
    _, _, blk_content = page_geometry(blk_set)
    for sd in ("before", "after"):
        td_b, bmap_b = render_side(blk_doc, blk_ana, blk_set, sd, blk_content)
        lay_b = td_b.documentLayout()
        y0 = lay_b.blockBoundingRect(bmap_b[0]).y()
        y5 = lay_b.blockBoundingRect(bmap_b[5]).y()
        y7 = lay_b.blockBoundingRect(bmap_b[7]).y()
        pc_b = page_count(td_b)
        check(f"预览空白行渲染({sd})：中间4个空行拉开段间距",
              (y5 - y0) > 60.0,
              f"y0={y0:.1f} y5={y5:.1f} delta={y5 - y0:.1f}")
        check(f"预览分页符与文末空段裁剪({sd})：val=0不误分页、br分页生效且总页数=2",
              y5 < blk_content[1] and y7 >= blk_content[1] - 0.5 and pc_b == 2,
              f"y5={y5:.1f} y7={y7:.1f} pages={pc_b} ch={blk_content[1]:.1f}")

    # ---- 兼容旧方案中 line=None（如从模板识别单倍行距公式后存下的方案）----
    null_set = Settings()
    null_set.roles["body"]["line"] = None
    null_set.roles["body"]["line_rule"] = "single"
    null_set.roles["equation"]["line"] = None
    null_set.roles["equation"]["line_rule"] = "single"
    td_null, _ = render_side(blk_doc, blk_ana, null_set, "after", blk_content)
    check("预览兼容方案中 line=None（单倍行距角色不崩）",
          page_count(td_null) >= 1, f"pages={page_count(td_null)}")
    from detect.recognize import settings_from_report as _apply_rep
    _rep_s = _apply_rep(Settings(), {"roles": {"equation": {"line": ("single", None)}}})
    check("格式识别写回单倍行距时不把 role.line 覆盖成 None",
          _rep_s.roles["equation"]["line_rule"] == "single"
          and _rep_s.roles["equation"]["line"] is not None,
          repr(_rep_s.roles["equation"]))

    # ---- 表格边框识别（Table Grid 改三线表、无线表、识别弹窗核对条）----
    from docx.oxml import OxmlElement as _OxmlElement
    from docx.oxml.ns import qn as _qn
    from detect.recognize import inspect_table_format as _inspect_tbl
    from detect.ui.scheme_dialogs import DetectDialog as RecognizeDialog
    from detect.ui.dialogs import ParaFormatDialog, ROLE_CHOICES

    _bdr_doc = Document()
    _t3 = _bdr_doc.add_table(rows=2, cols=2)
    _t3.style = "Table Grid"
    _t3.cell(0, 0).text = "列A"
    _t3.cell(1, 0).text = "1"
    _tp3 = _t3._tbl.find(_qn("w:tblPr"))
    _tb3 = _OxmlElement("w:tblBorders")
    for _side, _val, _sz in (("top", "single", "12"), ("bottom", "single", "12"),
                             ("left", "none", "0"), ("right", "none", "0"),
                             ("insideH", "none", "0"), ("insideV", "none", "0")):
        _e = _OxmlElement(f"w:{_side}")
        _e.set(_qn("w:val"), _val)
        _e.set(_qn("w:sz"), _sz)
        _tb3.append(_e)
    _tp3.append(_tb3)
    for _tc in _t3._tbl.tr_lst[0].tc_lst:
        _tcPr = _tc.get_or_add_tcPr()
        _tcb = _OxmlElement("w:tcBorders")
        _bot = _OxmlElement("w:bottom")
        _bot.set(_qn("w:val"), "single")
        _bot.set(_qn("w:sz"), "4")
        _tcb.append(_bot)
        _tcPr.append(_tcb)
    _res3 = _inspect_tbl(_bdr_doc)
    check("表格边框识别：Table Grid 显式去竖线识别为三线表且读出线宽",
          _res3["style"] == "three_line" and _res3["line"] == 1.5 and _res3["line_inner"] == 0.5,
          repr(_res3))

    _none_doc = Document()
    _tn = _none_doc.add_table(rows=2, cols=2)
    _tn.cell(0, 0).text = "无线"
    _tpn = _tn._tbl.find(_qn("w:tblPr"))
    _tbn = _OxmlElement("w:tblBorders")
    for _side in ("top", "bottom", "left", "right", "insideH", "insideV"):
        _e = _OxmlElement(f"w:{_side}")
        _e.set(_qn("w:val"), "none")
        _e.set(_qn("w:sz"), "0")
        _tbn.append(_e)
    _tpn.append(_tbn)
    _resn = _inspect_tbl(_none_doc)
    check("表格边框识别：全无框线识别为 none",
          _resn["style"] == "none", repr(_resn))

    _rdlg = RecognizeDialog(doc=_bdr_doc, analysis=analyze(_bdr_doc, "bdr.docx"),
                            base_settings=Settings())
    _rdlg._detect()
    check("格式识别器弹窗正确回填表格边框样式与线宽",
          _rdlg.tbl_style.currentData() == "three_line"
          and abs(_rdlg.tbl_line.value() - 1.5) < 1e-6
          and abs(_rdlg.tbl_line_inner.value() - 0.5) < 1e-6,
          f"st={_rdlg.tbl_style.currentData()} o={_rdlg.tbl_line.value()} i={_rdlg.tbl_line_inner.value()}")
    _rdlg.tbl_style.setCurrentIndex(_rdlg.tbl_style.findData("grid"))
    _rdlg.tbl_line.setValue(1.0)
    _rdlg.tbl_line_inner.setValue(0.75)
    _cre = _rdlg._collect()
    _cs = _apply_rep(Settings(), _cre)
    check("格式识别器弹窗修改表格边框后正确写回方案",
          _cs.table_style == "grid" and _cs.table_line == 1.0 and _cs.table_line_inner == 0.75,
          f"st={_cs.table_style} o={_cs.table_line} i={_cs.table_line_inner}")
    _rdlg.close()

    # ---- 悬浮菜单「设为」包含「作者信息」(meta) 且排在前排可设 ----
    _meta_idx = [k for k, _ in ROLE_CHOICES].index("meta") if any(k == "meta" for k, _ in ROLE_CHOICES) else -1
    _meta_label = dict(ROLE_CHOICES).get("meta", "")
    check("悬浮菜单角色列表包含「作者信息」(meta) 且排在前列",
          0 <= _meta_idx <= 2 and _meta_label == "作者信息",
          f"idx={_meta_idx} label={_meta_label!r}")
    _pdlg = ParaFormatDialog(0, Settings().roles["body"])
    _pdlg.set_role("meta")
    check("悬浮菜单下拉框可选中「作者信息」",
          _pdlg.role.currentData() == "meta" and _pdlg.role.currentText() == "作者信息",
          f"data={_pdlg.role.currentData()!r} text={_pdlg.role.currentText()!r}")
    _pdlg.close()

    # ---- OMML 公式数组 eqArr / 空矩阵 / fName 防闪退回归 ----
    from lxml import etree as _etree
    from detect.omml import render_element as _render_omml
    _MNS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
    _eq_xml = (
        f'<m:oMath xmlns:m="{_MNS}">'
        '<m:eqArr><m:e><m:r><m:t>a=1</m:t></m:r></m:e>'
        '<m:e><m:r><m:t>b=2</m:t></m:r></m:e></m:eqArr>'
        '<m:r><m:t>)</m:t></m:r></m:oMath>'
    )
    _eq_im = _render_omml(_etree.fromstring(_eq_xml), 10.5, "Times New Roman", False)
    check("OMML eqArr（含 m:e 子项）正常渲染不闪退",
          _eq_im is not None and _eq_im.width() > 0 and _eq_im.height() > 0,
          f"im={_eq_im.width() if _eq_im else None}x{_eq_im.height() if _eq_im else None}")

    _empty_eq_xml = (
        f'<m:oMath xmlns:m="{_MNS}"><m:eqArr/><m:r><m:t>)</m:t></m:r></m:oMath>'
    )
    _empty_eq_im = _render_omml(_etree.fromstring(_empty_eq_xml), 10.5, "Times New Roman", False)
    check("OMML 空 eqArr 搭配 m:r 不触发 Qt SIGABRT",
          _empty_eq_im is not None and _empty_eq_im.width() > 0,
          f"im={_empty_eq_im}")

    _empty_mat_xml = (
        f'<m:oMath xmlns:m="{_MNS}"><m:m><m:mr/></m:m><m:r><m:t>x</m:t></m:r></m:oMath>'
    )
    _empty_mat_im = _render_omml(_etree.fromstring(_empty_mat_xml), 10.5, "Times New Roman", False)
    check("OMML 空行矩阵搭配 m:r 不抛错不闪退",
          _empty_mat_im is not None and _empty_mat_im.width() > 0,
          f"im={_empty_mat_im}")

    _func_xml = (
        f'<m:oMath xmlns:m="{_MNS}"><m:func>'
        '<m:fName><m:r><m:t>sin</m:t></m:r></m:fName>'
        '<m:e><m:r><m:t>x</m:t></m:r></m:e></m:func></m:oMath>'
    )
    _func_im = _render_omml(_etree.fromstring(_func_xml), 10.5, "Times New Roman", False)
    check("OMML func(m:fName) 正常测量与渲染",
          _func_im is not None and _func_im.width() > 0,
          f"im={_func_im.width() if _func_im else None}")

    # ---- 摘要/关键词仅标签加粗（前缀识别不误判整段、跨 run 拆分、不加粗正文与关键词本体）----
    from detect.recognize import recognize_from_doc as _rec_doc
    from detect.reflow import apply_formats as _apply_fmts
    _ak_doc = Document()
    _p_abs = _ak_doc.add_paragraph()
    _r1 = _p_abs.add_run("摘要"); _r1.bold = True
    _r2 = _p_abs.add_run("："); _r2.bold = True
    _p_abs.add_run("本文系统研究了深海柔性立管在热力耦合作用下的力学行为与蠕变特性。")
    _p_kw = _ak_doc.add_paragraph()
    _rk = _p_kw.add_run("关键词："); _rk.bold = True
    _p_kw.add_run("柔性立管；热力耦合；蠕变；数值模拟")
    _p_en_abs = _ak_doc.add_paragraph()
    _re1 = _p_en_abs.add_run("Abstract"); _re1.bold = True
    _re2 = _p_en_abs.add_run(": "); _re2.bold = True
    _p_en_abs.add_run("This paper investigates the thermo-mechanical coupling behavior.")
    _p_en_kw = _ak_doc.add_paragraph()
    _rek = _p_en_kw.add_run("Keywords: "); _rek.bold = True
    _p_en_kw.add_run("flexible riser; coupling; creep")
    _ak_ana = analyze(_ak_doc, "ak.docx")
    _ak_rec = _rec_doc(_ak_doc, _ak_ana)
    check("摘要/关键词仅前缀加粗时，格式识别不误判为整段加粗",
          _ak_rec.roles["abstract"]["bold"] is False and _ak_rec.roles["keywords"]["bold"] is False,
          f"abs={_ak_rec.roles['abstract']['bold']} kw={_ak_rec.roles['keywords']['bold']}")

    _ak_s = Settings()
    _ak_s.roles["abstract"]["bold"] = True
    _ak_s.roles["keywords"]["bold"] = True
    _apply_fmts(_ak_doc, _ak_ana, _ak_s.roles, label_bold=True)
    _abs_runs = [(r.text, r.bold) for r in _ak_doc.paragraphs[0].runs if r.text]
    _kw_runs = [(r.text, r.bold) for r in _ak_doc.paragraphs[1].runs if r.text]
    _eabs_runs = [(r.text, r.bold) for r in _ak_doc.paragraphs[2].runs if r.text]
    _ekw_runs = [(r.text, r.bold) for r in _ak_doc.paragraphs[3].runs if r.text]
    check("摘要/关键词同行段落即使角色勾选加粗也只加粗中英文标签、不加粗正文与关键词",
          _abs_runs[-1][1] is False and _abs_runs[0][1] is True
          and _kw_runs[-1][1] is False and _kw_runs[0][1] is True
          and _eabs_runs[-1][1] is False and _eabs_runs[0][1] is True and _eabs_runs[1][1] is True
          and _ekw_runs[-1][1] is False and _ekw_runs[0][1] is True,
          f"abs={_abs_runs} kw={_kw_runs} eabs={_eabs_runs} ekw={_ekw_runs}")

    # 预览侧同步验证：同行摘要/关键词只加粗标签，后面正文 charFormat.fontWeight 非加粗
    _ak_qdoc, _ = render_side(_ak_doc, _ak_ana, _ak_s, "after", (415.0, 680.0))
    _b0 = _ak_qdoc.findBlockByNumber(0)
    _frags0 = []
    _it = _b0.begin()
    while not _it.atEnd():
        _fr = _it.fragment()
        if _fr.isValid():
            _frags0.append((_fr.text(), _fr.charFormat().fontWeight() >= 700))
        _it += 1
    check("预览渲染同行摘要只加粗标签、正文不加粗",
          len(_frags0) >= 2 and _frags0[0][1] is True and _frags0[-1][1] is False,
          repr(_frags0))

    print("\n" + ("全部通过" if not FAIL else f"失败 {len(FAIL)} 项: {FAIL}"))
    win.close()
    return 1 if FAIL else 0


if __name__ == "__main__":
    raise SystemExit(main())
