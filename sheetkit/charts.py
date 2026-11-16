"""Charts that float over the grid (column, bar, line, area, pie, scatter) and follow the data."""
import math

from PySide6.QtCore import QPointF, QRect, QRectF, QSize, Qt, Signal
from PySide6.QtGui import QAction, QBrush, QColor, QFont, QFontMetricsF, QImage, QPainter, QPainterPath, QPen, QPolygonF, QRegion
from PySide6.QtWidgets import QApplication, QFileDialog, QMenu, QWidget

from . import formula as F
from .cfmt import parse_range

PALETTE = ["#C2395C", "#E8710A", "#1A73E8", "#D93025", "#F9AB00", "#9334E6", "#12B5CB", "#7CB342", "#E91E63", "#5F6368"]
KINDS = [("column", "Column"), ("bar", "Bar"), ("line", "Line"), ("area", "Area"), ("pie", "Pie"), ("scatter", "Scatter")]


def nice_ticks(lo, hi, n=5):
    if hi == lo:
        hi = lo + 1
    span = hi - lo
    step = 10 ** math.floor(math.log10(span / n))
    for m in (1, 2, 2.5, 5, 10):
        if span / (step * m) <= n:
            step *= m
            break
    a = math.floor(lo / step) * step
    b = math.ceil(hi / step) * step
    ticks, v = [], a
    while v <= b + step * 1e-6:
        ticks.append(round(v, 10))
        v += step
    return ticks


def fmt_tick(v):
    if abs(v) >= 1e6:
        return f"{v / 1e6:g}M"
    if abs(v) >= 1e4:
        return f"{v / 1e3:g}k"
    return f"{v:g}"


def read_data(wb, ch, refs=None):
    """-> (categories, [(name, values)]) from the chart's range. refs: an empty dict that is filled with cell references for export"""
    rng = parse_range(ch["range"].split("!")[-1]) if ch.get("range") else None
    sheet = ch["range"].split("!")[0].strip("'") if "!" in ch.get("range", "") else None
    if not rng or not (sheet and wb.sheet(sheet)):
        return [], []
    r1, c1, r2, c2 = rng
    ur, uc = wb.used_extent(sheet)
    r2, c2 = min(r2, max(ur, r1)), min(c2, max(uc, c1))
    m = [[wb.cell_value(sheet, r, c) for c in range(c1, c2 + 1)] for r in range(r1, r2 + 1)]
    if not m or not m[0]:
        return [], []
    if ch.get("by_rows"):
        m = [list(x) for x in zip(*m)]

    def isnum(v):
        return isinstance(v, (int, float)) and not isinstance(v, bool)
    header = len(m) > 1 and any(isinstance(v, str) for v in m[0][1:]) and any(isnum(v) for row in m[1:] for v in row[1:])
    if len(m[0]) == 1:
        header = len(m) > 1 and isinstance(m[0][0], str) and any(isnum(v) for v in [row[0] for row in m[1:]])
    labels_col = len(m[0]) > 1 and all(not isnum(row[0]) for row in m[1 if header else 0:] if row[0] is not None) and any(row[0] is not None for row in m[1 if header else 0:])
    body = m[1:] if header else m
    cats = [("" if row[0] is None else (F.text(row[0]) if not isinstance(row[0], F.XErr) else "#")) for row in body] if labels_col else [str(i + 1) for i in range(len(body))]
    first = 1 if labels_col else 0
    series = []
    if refs is not None:
        by_rows = bool(ch.get("by_rows"))
        shq = F.quote_sheet(sheet)
        b0 = 1 if header else 0
        last = len(m) - 1

        def co(i, j):
            return (r1 + j, c1 + i) if by_rows else (r1 + i, c1 + j)

        def rg(i1, j1, i2, j2):
            (ra, ca), (rb, cb) = co(i1, j1), co(i2, j2)
            return f"{shq}!{F.cell_name(ra, ca, True, True)}:{F.cell_name(rb, cb, True, True)}"
        refs["cat"] = rg(b0, 0, last, 0) if labels_col else None
        refs["series"] = []
    for j in range(first, len(m[0])):
        if refs is not None:
            refs["series"].append({"name": (lambda a: f"{shq}!{F.cell_name(a[0], a[1], True, True)}")(co(0, j)) if header and m[0][j] is not None else None,
                                   "val": rg(b0, j, last, j)})
        name = F.text(m[0][j]) if header and m[0][j] is not None else f"Series {j - first + 1}"
        vals = [row[j] if isnum(row[j]) else (0.0 if row[j] in (None, "") else None) for row in body]
        series.append((name, vals))
    return cats, series


class ChartWidget(QWidget):
    changed = Signal()
    edit_requested = Signal(object)
    delete_requested = Signal(object)

    def __init__(self, wb, ch):
        super().__init__()
        self.wb, self.ch = wb, ch
        self.setMouseTracking(True)
        self.setFocusPolicy(Qt.ClickFocus)
        self._drag = None
        self.selected = False

    def place(self, grid):
        z = grid.zoom
        ch = self.ch
        x = grid.hdr_w + (ch["x"] - grid.sx) * z
        y = grid.hdr_h + (ch["y"] - grid.sy) * z
        w, h = ch["w"] * z, ch["h"] * z
        self.setGeometry(int(x), int(y), int(w), int(h))
        vis = QRect(grid.hdr_w, grid.hdr_h, grid.viewport().width(), grid.viewport().height()).translated(-int(x), -int(y))
        self.setMask(QRegion(vis.intersected(QRect(0, 0, int(w), int(h)))))
        self.raise_()

    # -------------------------------------------------------- drawing
    def paintEvent(self, ev):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing, True)
        p.setRenderHint(QPainter.TextAntialiasing, True)
        self.draw(p, QRectF(0, 0, self.width(), self.height()), self.height() / max(1, self.ch["h"]))
        if self.selected:
            p.setPen(QPen(QColor("#C2395C"), 2))
            p.setBrush(Qt.NoBrush)
            p.drawRect(QRectF(1, 1, self.width() - 2, self.height() - 2))
            p.setBrush(QColor("#C2395C"))
            for cx, cy in ((0, 0), (self.width(), 0), (0, self.height()), (self.width(), self.height())):
                p.drawRect(QRectF(cx - 4 if cx else 0, cy - 4 if cy else 0, 8, 8))
        p.end()

    def draw(self, p, rect, k=1.0):
        ch = self.ch
        p.fillRect(rect, QColor("#FFFFFF"))
        p.setPen(QPen(QColor("#CBD5CF"), 1))
        p.drawRect(rect.adjusted(0.5, 0.5, -0.5, -0.5))
        cats, series = read_data(self.wb, ch)
        f = QFont(QApplication.font())
        f.setPixelSize(max(7, int(12 * k)))
        p.setFont(f)
        fm = QFontMetricsF(f)
        top = 8 * k
        title = ch.get("title") or ""
        if title:
            tf = QFont(f)
            tf.setPixelSize(max(8, int(15 * k)))
            tf.setBold(True)
            p.setFont(tf)
            p.setPen(QColor("#1B1040"))
            p.drawText(QRectF(rect.left(), rect.top() + 6 * k, rect.width(), 22 * k), Qt.AlignCenter, title)
            p.setFont(f)
            top = 32 * k
        if not series:
            p.setPen(QColor("#8A7FA6"))
            p.drawText(rect, Qt.AlignCenter, "No data in " + ch.get("range", ""))
            return
        kind = ch["kind"]
        # legend
        show_legend = ch.get("legend", True) and (len(series) > 1 or kind == "pie")
        names = cats if kind == "pie" else [s[0] for s in series]
        leg_h = 0
        if show_legend:
            x = rect.left() + 10 * k
            y = rect.bottom() - 20 * k
            leg_h = 24 * k
            for i, n in enumerate(names):
                w = fm.horizontalAdvance(n) + 26 * k
                if x + w > rect.right() - 6:
                    break
                p.setBrush(QColor(PALETTE[i % len(PALETTE)]))
                p.setPen(Qt.NoPen)
                p.drawRect(QRectF(x, y + 3 * k, 10 * k, 10 * k))
                p.setPen(QColor("#33443B"))
                p.drawText(QRectF(x + 14 * k, y, w, 16 * k), Qt.AlignVCenter | Qt.AlignLeft, n)
                x += w
        area = QRectF(rect.left() + 8 * k, top, rect.width() - 16 * k, rect.height() - top - leg_h - 6 * k)
        if kind == "pie":
            self._pie(p, area, cats, series[0][1], k)
        elif kind == "scatter":
            self._scatter(p, area, cats, series, fm, k)
        else:
            self._axes_chart(p, area, kind, cats, series, fm, k)

    def _pie(self, p, area, cats, vals, k):
        vals = [max(0.0, v or 0.0) for v in vals]
        tot = sum(vals) or 1
        d = min(area.width(), area.height()) - 8 * k
        box = QRectF(area.center().x() - d / 2, area.center().y() - d / 2, d, d)
        ang = 90 * 16
        for i, v in enumerate(vals):
            span = -int(v / tot * 360 * 16)
            p.setBrush(QColor(PALETTE[i % len(PALETTE)]))
            p.setPen(QPen(QColor("#FFFFFF"), 1.5))
            p.drawPie(box, ang, span)
            if v / tot > 0.05:
                mid = math.radians(ang / 16 + span / 32)
                tx = box.center().x() + math.cos(mid) * d * 0.33
                ty = box.center().y() - math.sin(mid) * d * 0.33
                p.setPen(QColor("#FFFFFF"))
                p.drawText(QRectF(tx - 30 * k, ty - 8 * k, 60 * k, 16 * k), Qt.AlignCenter, f"{v / tot * 100:.0f}%")
            ang += span

    def _axes_chart(self, p, area, kind, cats, series, fm, k):
        nums = [v for s in series for v in s[1] if v is not None]
        if not nums:
            return
        stacked = kind in ("column", "bar", "area") and self.ch.get("stacked")
        if stacked:
            n = len(cats)
            tot = [sum(max(0, (s[1][i] or 0)) for s in series) for i in range(n)]
            lo, hi = min(0, min(nums)), max(tot)
        else:
            lo, hi = min(0, min(nums)), max(0, max(nums))
        ticks = nice_ticks(lo, hi)
        lo, hi = ticks[0], ticks[-1]
        horizontal = kind == "bar"
        labw = max(fm.horizontalAdvance(fmt_tick(t)) for t in ticks) + 8 * k
        catw = max([fm.horizontalAdvance(c[:14]) for c in cats] + [10]) + 8 * k if horizontal else 0
        left = area.left() + (catw if horizontal else labw)
        bottom = area.bottom() - (fm.height() + 6 * k if not horizontal else labw * 0 + fm.height() + 6 * k)
        plot = QRectF(left, area.top() + 4 * k, area.right() - left - 4 * k, bottom - area.top() - 4 * k)
        if plot.width() < 20 or plot.height() < 20:
            return
        p.setPen(QPen(QColor("#E3E8E5"), 1))
        n = max(1, len(cats))

        def vpos(v):
            t = (v - lo) / ((hi - lo) or 1)
            return plot.bottom() - t * plot.height() if not horizontal else plot.left() + t * plot.width()
        for t in ticks:
            pos = vpos(t)
            if horizontal:
                p.drawLine(QPointF(pos, plot.top()), QPointF(pos, plot.bottom()))
                p.setPen(QColor("#5B6B63"))
                p.drawText(QRectF(pos - 30 * k, plot.bottom() + 2 * k, 60 * k, fm.height()), Qt.AlignCenter, fmt_tick(t))
            else:
                p.drawLine(QPointF(plot.left(), pos), QPointF(plot.right(), pos))
                p.setPen(QColor("#5B6B63"))
                p.drawText(QRectF(area.left(), pos - fm.height() / 2, labw - 4 * k, fm.height()), Qt.AlignRight | Qt.AlignVCenter, fmt_tick(t))
            p.setPen(QPen(QColor("#E3E8E5"), 1))
        slot = (plot.height() if horizontal else plot.width()) / n
        step = max(1, int(math.ceil(fm.horizontalAdvance("MMMMMM") / max(slot, 1)))) if not horizontal else 1
        p.setPen(QColor("#5B6B63"))
        for i, c in enumerate(cats):
            if i % step:
                continue
            if horizontal:
                p.drawText(QRectF(area.left(), plot.top() + i * slot, catw - 4 * k, slot), Qt.AlignRight | Qt.AlignVCenter, c[:14])
            else:
                p.drawText(QRectF(plot.left() + i * slot - 10 * k, plot.bottom() + 3 * k, slot + 20 * k, fm.height()), Qt.AlignCenter, c[:14])
        zero = vpos(0)
        p.setPen(QPen(QColor("#8A7FA6"), 1))
        if horizontal:
            p.drawLine(QPointF(zero, plot.top()), QPointF(zero, plot.bottom()))
        else:
            p.drawLine(QPointF(plot.left(), zero), QPointF(plot.right(), zero))
        ns = len(series)
        if kind in ("column", "bar"):
            gap = slot * 0.2
            bw = (slot - gap) / (1 if stacked else ns)
            for si, (name, vals) in enumerate(series):
                p.setBrush(QColor(PALETTE[si % len(PALETTE)]))
                p.setPen(Qt.NoPen)
                for i, v in enumerate(vals):
                    if v is None:
                        continue
                    base = sum(max(0, (series[q][1][i] or 0)) for q in range(si)) if stacked else 0
                    a, b = vpos(base), vpos(base + v)
                    off = i * slot + gap / 2 + (0 if stacked else si * bw)
                    if horizontal:
                        p.drawRect(QRectF(min(a, b), plot.top() + off, abs(b - a), bw))
                    else:
                        p.drawRect(QRectF(plot.left() + off, min(a, b), bw, abs(b - a)))
        else:
            for si, (name, vals) in enumerate(series):
                col = QColor(PALETTE[si % len(PALETTE)])
                pts = []
                for i, v in enumerate(vals):
                    if v is None:
                        continue
                    base = sum(max(0, (series[q][1][i] or 0)) for q in range(si)) if stacked else 0
                    pts.append(QPointF(plot.left() + (i + 0.5) * slot, vpos(base + v)))
                if not pts:
                    continue
                if kind == "area":
                    poly = QPolygonF([QPointF(pts[0].x(), zero)] + pts + [QPointF(pts[-1].x(), zero)])
                    c2 = QColor(col)
                    c2.setAlpha(110)
                    p.setBrush(c2)
                    p.setPen(Qt.NoPen)
                    p.drawPolygon(poly)
                p.setPen(QPen(col, 2.2 * k, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin))
                p.setBrush(Qt.NoBrush)
                p.drawPolyline(QPolygonF(pts))
                p.setBrush(col)
                if kind == "line":
                    for pt in pts:
                        p.drawEllipse(pt, 3 * k, 3 * k)

    def _scatter(self, p, area, cats, series, fm, k):
        xs = []
        for c in cats:
            try:
                xs.append(float(c.replace(",", "")))
            except ValueError:
                xs.append(float(len(xs) + 1))
        ys = [v for s in series for v in s[1] if v is not None]
        if not ys or not xs:
            return
        tx, ty = nice_ticks(min(xs), max(xs)), nice_ticks(min(0, min(ys)), max(ys))
        labw = max(fm.horizontalAdvance(fmt_tick(t)) for t in ty) + 8 * k
        plot = QRectF(area.left() + labw, area.top() + 4 * k, area.width() - labw - 6 * k, area.height() - fm.height() - 12 * k)
        p.setPen(QPen(QColor("#E3E8E5"), 1))
        for t in ty:
            y = plot.bottom() - (t - ty[0]) / ((ty[-1] - ty[0]) or 1) * plot.height()
            p.drawLine(QPointF(plot.left(), y), QPointF(plot.right(), y))
            p.setPen(QColor("#5B6B63"))
            p.drawText(QRectF(area.left(), y - fm.height() / 2, labw - 4 * k, fm.height()), Qt.AlignRight | Qt.AlignVCenter, fmt_tick(t))
            p.setPen(QPen(QColor("#E3E8E5"), 1))
        for t in tx:
            x = plot.left() + (t - tx[0]) / ((tx[-1] - tx[0]) or 1) * plot.width()
            p.drawLine(QPointF(x, plot.top()), QPointF(x, plot.bottom()))
            p.setPen(QColor("#5B6B63"))
            p.drawText(QRectF(x - 25 * k, plot.bottom() + 3 * k, 50 * k, fm.height()), Qt.AlignCenter, fmt_tick(t))
            p.setPen(QPen(QColor("#E3E8E5"), 1))
        for si, (name, vals) in enumerate(series):
            p.setBrush(QColor(PALETTE[si % len(PALETTE)]))
            p.setPen(Qt.NoPen)
            for x, v in zip(xs, vals):
                if v is None:
                    continue
                p.drawEllipse(QPointF(plot.left() + (x - tx[0]) / ((tx[-1] - tx[0]) or 1) * plot.width(),
                                      plot.bottom() - (v - ty[0]) / ((ty[-1] - ty[0]) or 1) * plot.height()), 3.5 * k, 3.5 * k)

    def to_image(self, scale=2):
        img = QImage(int(self.ch["w"] * scale), int(self.ch["h"] * scale), QImage.Format_ARGB32)
        img.fill(Qt.white)
        p = QPainter(img)
        p.setRenderHint(QPainter.Antialiasing, True)
        p.setRenderHint(QPainter.TextAntialiasing, True)
        self.draw(p, QRectF(0, 0, img.width(), img.height()), scale)
        p.end()
        return img

    # -------------------------------------------------------- mouse
    def _edge(self, pos):
        m = 7
        l, r, t, b = pos.x() <= m, pos.x() >= self.width() - m, pos.y() <= m, pos.y() >= self.height() - m
        return l, r, t, b

    def mouseMoveEvent(self, e):
        pos = e.position()
        if self._drag is None:
            l, r, t, b = self._edge(pos)
            if (l and t) or (r and b):
                self.setCursor(Qt.SizeFDiagCursor)
            elif (r and t) or (l and b):
                self.setCursor(Qt.SizeBDiagCursor)
            elif l or r:
                self.setCursor(Qt.SizeHorCursor)
            elif t or b:
                self.setCursor(Qt.SizeVerCursor)
            else:
                self.setCursor(Qt.SizeAllCursor)
            return
        g = self.parent().parent()
        z = g.zoom
        dx, dy = (e.globalPosition().x() - self._drag[1].x()) / z, (e.globalPosition().y() - self._drag[1].y()) / z
        mode, _, orig = self._drag
        ch = self.ch
        if mode == "move":
            ch["x"], ch["y"] = orig["x"] + dx, orig["y"] + dy
        else:
            l, r, t, b = mode
            x, y, w, h = orig["x"], orig["y"], orig["w"], orig["h"]
            if r:
                w = max(120, orig["w"] + dx)
            if b:
                h = max(90, orig["h"] + dy)
            if l:
                w = max(120, orig["w"] - dx)
                x = orig["x"] + orig["w"] - w
            if t:
                h = max(90, orig["h"] - dy)
                y = orig["y"] + orig["h"] - h
            ch.update(x=x, y=y, w=w, h=h)
        self.place(g)
        self.update()

    def mousePressEvent(self, e):
        self.setFocus()
        for w in self.parent().findChildren(ChartWidget):
            w.selected = w is self
            w.update()
        if e.button() == Qt.RightButton:
            self.menu(e.globalPosition().toPoint())
            return
        l, r, t, b = self._edge(e.position())
        mode = (l, r, t, b) if (l or r or t or b) else "move"
        self._drag = (mode, e.globalPosition(), dict(self.ch))

    def mouseReleaseEvent(self, e):
        if self._drag:
            self._drag = None
            self.changed.emit()

    def mouseDoubleClickEvent(self, e):
        self.edit_requested.emit(self)

    def keyPressEvent(self, e):
        if e.key() in (Qt.Key_Delete, Qt.Key_Backspace):
            self.delete_requested.emit(self)
        else:
            super().keyPressEvent(e)

    def menu(self, pos):
        m = QMenu(self)
        m.addAction("Edit chart ...", lambda: self.edit_requested.emit(self))
        tm = m.addMenu("Chart type")
        for k, label in KINDS:
            tm.addAction(label, lambda k=k: (self.ch.update(kind=k), self.update(), self.changed.emit()))
        m.addAction("Save as picture ...", self.save_png)
        m.addAction("Copy picture", lambda: QApplication.clipboard().setImage(self.to_image()))
        m.addSeparator()
        m.addAction("Delete chart", lambda: self.delete_requested.emit(self))
        m.exec(pos)

    def save_png(self):
        path, _ = QFileDialog.getSaveFileName(self, "Save chart", "chart.png", "PNG image (*.png)")
        if path:
            self.to_image(3).save(path)
