"""确认预览里表格和图片存在，并抓出它们所在的页。"""
import io
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

from PySide6.QtCore import Qt, QCoreApplication, QRectF
from PySide6.QtWidgets import QApplication, QFileDialog

from detect.ui.main_window import MainWindow
from detect.ui.style import apply

QCoreApplication.setAttribute(Qt.AA_ShareOpenGLContexts, True)
app = QApplication([])
apply(app)

src = Path(sys.argv[1])
win = MainWindow()
win.resize(1500, 900)
QFileDialog.getOpenFileName = staticmethod(lambda *a, **k: (str(src), ""))
win.open_file()
app.processEvents()

v = win.after_view
td = v.document()
v.refit(900)

frames = list(td.rootFrame().childFrames())
print("frames(tables):", len(frames))

imgs = 0
first_img_page = None
for i in range(td.blockCount()):
    b = td.findBlockByNumber(i)
    r = td.documentLayout().blockBoundingRect(b)
    it = b.begin()
    while not it.atEnd():
        frag = it.fragment()
        if frag.isValid() and frag.charFormat().isImageFormat():
            imgs += 1
            if first_img_page is None:
                first_img_page = int(r.y() / v._content_h)
        it += 1
print("images:", imgs, "first image page:", first_img_page)

table_page = 0
if frames:
    fr = td.documentLayout().frameBoundingRect(frames[0])
    print("first table rect:", round(fr.x(), 1), round(fr.y(), 1),
          round(fr.width(), 1), round(fr.height(), 1))
    table_page = int(fr.y() / v._content_h)
    print("first table page:", table_page)

z = v._zoom
ph = v._paper[1] * z
for name, page in (("p17", 17), ("p18", 18), ("p19", 19)):
    y = v.PAD + page * (ph + v.GAP)
    v.grab(QRectF(0, y, v.width(), min(ph, 1500)).toRect()).save(
        f"build/v2_{name}.png")
    print("saved", f"build/v2_{name}.png")
win.close()
