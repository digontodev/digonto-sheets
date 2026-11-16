"""The spreadsheet grid: a scroll area that paints its own cells (fast, frozen panes, merged cells, zoom, overflow text)."""
import json
import math
import re

from PySide6.QtCore import QEvent, QMimeData, QPoint, QPointF, QRect, QRectF, QSize, Qt, QTimer, Signal
from PySide6.QtGui import (QBrush, QClipboard, QColor, QCursor, QFont, QFontMetricsF, QGuiApplication, QKeySequence, QPainter, QPen,
                           QPolygonF, QTextOption)
from PySide6.QtWidgets import QAbstractScrollArea, QApplication, QLineEdit, QMenu, QToolTip

from . import formula as F
from . import validation as V
from .axis import Axis
from .cfmt import CF, parse_range, range_text
from .model import DEFAULT, DEFAULT_H, DEFAULT_W, Style
from .numfmt import format_value

ROWS, COLS = F.MAXROW, F.MAXCOL
GREEN = QColor("#C2395C")
GRID = QColor("#D4DAD7")
HDR_BG, HDR_SEL, HDR_FULL = QColor("#F1F4F2"), QColor("#E0D4E9"), QColor("#CEBCDD")
HDR_TXT, HDR_TXT_SEL = QColor("#534A5A"), QColor("#8A2150")
SEL_TINT = QColor(0, 128, 90, 36)
OPS = "=+-*/^&,;(<>:%"


def norm(r1, c1, r2, c2):
    return min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)


class CellEditor(QLineEdit):
    """in-cell editor. Enter / Tab / Esc and arrow keys are handled by the grid"""
    key = Signal(object)

    def __init__(self, grid):
        super().__init__(grid.viewport())
        self.grid = grid
        self.setFrame(False)
        self.mode = "edit"       # "enter" = typed over the cell (arrows move away), "edit" = F2 / double click
        self.setStyleSheet("QLineEdit{background:white; border:2px solid #C2395C; padding:0 2px; selection-background-color:#C2395C; selection-color:white;}")

    def keyPressEvent(self, e):
        k = e.key()
        if k in (Qt.Key_Return, Qt.Key_Enter, Qt.Key_Tab, Qt.Key_Backtab, Qt.Key_Escape):
            self.key.emit(e)
            return
        if k in (Qt.Key_Up, Qt.Key_Down, Qt.Key_PageUp, Qt.Key_PageDown) or (self.mode == "enter" and k in (Qt.Key_Left, Qt.Key_Right)):
            self.key.emit(e)
            return
        if k == Qt.Key_F4 and self.text().startswith("="):
            self.key.emit(e)
            return
        super().keyPressEvent(e)
        if self.mode == "enter" and k in (Qt.Key_Left, Qt.Key_Right):
            pass

    def event(self, e):
        if e.type() == QEvent.KeyPress and e.key() in (Qt.Key_Tab, Qt.Key_Backtab):
            self.key.emit(e)
            return True
        return super().event(e)


class Grid(QAbstractScrollArea):
    selection_changed = Signal()
    edit_text_changed = Signal(str)
    edit_state = Signal(bool)
    zoom_changed = Signal(float)
    filter_clicked = Signal(int, QPoint)
    context_menu = Signal(QPoint, str)
    message = Signal(str)
    axes_changed = Signal()
    undo_requested = Signal()
    redo_requested = Signal()
    key_filter = None

    def __init__(self, wb, parent=None):
        super().__init__(parent)
        self.wb = wb
        self.sheet = wb.sheets[0]
        self.zoom = 1.0
        self.ax_r = Axis(ROWS, DEFAULT_H)
        self.ax_c = Axis(COLS, DEFAULT_W)
        self.cur = (0, 0)
        self.anchor = (0, 0)
        self.sel = [(0, 0, 0, 0)]
        self.sx = self.sy = 0.0
        self.cf = CF()
        self.editor = None
        self.edit_origin = None
        self.copy_rect = None
        self._cut = None
        self._fonts = {}
        self._mode = None
        self._pt = None          # point mode (start, end) of an inserted reference in the editor
        self._pt_cell = None
        self.show_formulas = False
        self.base_font = QFont(QApplication.font())
        self.setFocusPolicy(Qt.StrongFocus)
        self.viewport().setMouseTracking(True)
        self.viewport().setAttribute(Qt.WA_OpaquePaintEvent, True)
        self.horizontalScrollBar().valueChanged.connect(self._scrolled)
        self.verticalScrollBar().valueChanged.connect(self._scrolled)
        self.horizontalScrollBar().setSingleStep(60)
        self.verticalScrollBar().setSingleStep(DEFAULT_H)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.auto = QTimer(self, interval=40, timeout=self._autoscroll)
        self._floating = []      # chart widgets
        self.set_sheet(self.sheet.name)

    # ------------------------------------------------------------ sheet / geometry
    def set_sheet(self, name):
        self.cancel_edit()
        self.sheet = self.wb.sheet(name)
        self.reload_axes()
        self.cur = self.anchor = (0, 0)
        self.sel = [(0, 0, 0, 0)]
        self.sx = self.sy = 0
        self.copy_rect = None
        self._update_scrollbars()
        self.viewport().update()
        self.selection_changed.emit()

    def reload_axes(self):
        s = self.sheet
        self.ax_r.set(s.rowh, s.hidden_rows)
        self.ax_c.set(s.colw, s.hidden_cols)
        self._merge_map = {}
        self._covered = set()
        for (r1, c1, r2, c2) in s.merges:
            self._merge_map[(r1, c1)] = (r1, c1, r2, c2)
            if (r2 - r1 + 1) * (c2 - c1 + 1) < 100000:
                for r in range(r1, r2 + 1):
                    for c in range(c1, c2 + 1):
                        if (r, c) != (r1, c1):
                            self._covered.add((r, c))
        self._update_scrollbars()

    def merge_at(self, r, c):
        """merged rectangle containing (r, c) or None"""
        if (r, c) in self._merge_map:
            return self._merge_map[(r, c)]
        if (r, c) in self._covered:
            for m in self.sheet.merges:
                if m[0] <= r <= m[2] and m[1] <= c <= m[3]:
                    return m
        return None

    @property
    def hdr_w(self):
        digits = max(2, len(str(self._max_visible_row() + 1)))
        return int((12 + digits * 8) * self.zoom)

    @property
    def hdr_h(self):
        return int(24 * self.zoom)

    def _max_visible_row(self):
        vh = self.viewport().height() / self.zoom
        return self.ax_r.index_at(self.ax_r.pos(self.sheet.freeze[0]) + self.sy + vh)

    def _frozen_px(self):
        fr, fc = self.sheet.freeze
        return self.ax_c.pos(fc) * self.zoom, self.ax_r.pos(fr) * self.zoom

    def col_x(self, c):
        fr, fc = self.sheet.freeze
        a = self.ax_c
        if c < fc:
            return self.hdr_w + a.pos(c) * self.zoom
        return self.hdr_w + (a.pos(c) - self.sx) * self.zoom

    def row_y(self, r):
        fr, fc = self.sheet.freeze
        a = self.ax_r
        if r < fr:
            return self.hdr_h + a.pos(r) * self.zoom
        return self.hdr_h + (a.pos(r) - self.sy) * self.zoom

    def col_w(self, c):
        return self.ax_c.size(c) * self.zoom

    def row_h(self, r):
        return self.ax_r.size(r) * self.zoom

    def cell_rect(self, r, c):
        m = self.merge_at(r, c)
        if m:
            r, c, r2, c2 = m
            x, y = self.col_x(c), self.row_y(r)
            return QRectF(x, y, self.col_x(c2) + self.col_w(c2) - x, self.row_y(r2) + self.row_h(r2) - y)
        return QRectF(self.col_x(c), self.row_y(r), self.col_w(c), self.row_h(r))

    def col_at(self, x):
        fr, fc = self.sheet.freeze
        hw = self.hdr_w
        fw = self.ax_c.pos(fc) * self.zoom
        if x < hw + fw:
            return self.ax_c.index_at((x - hw) / self.zoom)
        return self.ax_c.index_at((x - hw) / self.zoom + self.sx)

    def row_at(self, y):
        fr, fc = self.sheet.freeze
        fh = self.ax_r.pos(fr) * self.zoom
        if y < self.hdr_h + fh:
            return self.ax_r.index_at((y - self.hdr_h) / self.zoom)
        return self.ax_r.index_at((y - self.hdr_h) / self.zoom + self.sy)

    def _update_scrollbars(self):
        fr, fc = self.sheet.freeze
        vw = (self.viewport().width() - self.hdr_w) / self.zoom
        vh = (self.viewport().height() - self.hdr_h) / self.zoom
        tw = self.ax_c.total() - self.ax_c.pos(fc)
        th = self.ax_r.total() - self.ax_r.pos(fr)
        fw = self.ax_c.pos(fc)
        fh = self.ax_r.pos(fr)
        hb, vb = self.horizontalScrollBar(), self.verticalScrollBar()
        hb.blockSignals(True)
        vb.blockSignals(True)
        hb.setRange(0, int(max(0, tw - (vw - fw) + 40)))
        hb.setPageStep(int(max(10, vw - fw)))
        vb.setRange(0, int(max(0, th - (vh - fh) + 40)))
        vb.setPageStep(int(max(10, vh - fh)))
        hb.setValue(int(self.sx))
        vb.setValue(int(self.sy))
        hb.blockSignals(False)
        vb.blockSignals(False)
        self.sx, self.sy = float(hb.value()), float(vb.value())

    def _scrolled(self, *_):
        self.sx = float(self.horizontalScrollBar().value())
        self.sy = float(self.verticalScrollBar().value())
        self._place_editor()
        self._place_floating()
        self.viewport().update()

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self._update_scrollbars()
        self._place_editor()
        self._place_floating()

    def set_zoom(self, z):
        z = max(0.4, min(3.0, z))
        if abs(z - self.zoom) < 1e-3:
            return
        self.zoom = z
        self._fonts.clear()
        self._update_scrollbars()
        self._place_editor()
        self._place_floating()
        self.viewport().update()
        self.zoom_changed.emit(z)

    def wheelEvent(self, e):
        if e.modifiers() & Qt.ControlModifier:
            self.set_zoom(self.zoom * (1.1 if e.angleDelta().y() > 0 else 1 / 1.1))
            e.accept()
            return
        if e.modifiers() & Qt.ShiftModifier:
            dx = e.angleDelta().y() or e.angleDelta().x()
            hb = self.horizontalScrollBar()
            hb.setValue(hb.value() - int(dx / 120 * 90))
            e.accept()
            return
        dy, dx = e.angleDelta().y(), e.angleDelta().x()
        if dy:
            vb = self.verticalScrollBar()
            vb.setValue(vb.value() - int(dy / 120 * 3 * DEFAULT_H))
        if dx:
            hb = self.horizontalScrollBar()
            hb.setValue(hb.value() - int(dx / 120 * 90))
        e.accept()

    def ensure_visible(self, r, c):
        fr, fc = self.sheet.freeze
        vw = (self.viewport().width() - self.hdr_w) / self.zoom
        vh = (self.viewport().height() - self.hdr_h) / self.zoom
        if c >= fc:
            left = self.ax_c.pos(c) - self.ax_c.pos(fc)
            right = left + self.ax_c.size(c)
            fw = self.ax_c.pos(fc)
            avail = vw - fw
            if left < self.sx:
                self.sx = left
            elif right > self.sx + avail:
                self.sx = min(left, right - avail)
        if r >= fr:
            top = self.ax_r.pos(r) - self.ax_r.pos(fr)
            bot = top + self.ax_r.size(r)
            fh = self.ax_r.pos(fr)
            avail = vh - fh
            if top < self.sy:
                self.sy = top
            elif bot > self.sy + avail:
                self.sy = min(top, bot - avail)
        self.sx, self.sy = max(0.0, self.sx), max(0.0, self.sy)
        hb, vb = self.horizontalScrollBar(), self.verticalScrollBar()
        hb.blockSignals(True)
        vb.blockSignals(True)
        hb.setValue(int(self.sx))
        vb.setValue(int(self.sy))
        hb.blockSignals(False)
        vb.blockSignals(False)
        self.sx, self.sy = float(hb.value()), float(vb.value())
        self._place_editor()
        self._place_floating()

    # ------------------------------------------------------------ fonts
    def font_for(self, st):
        key = (st.font, st.size, st.bold, st.italic, st.underline, st.strike)
        f = self._fonts.get(key)
        if f is None:
            f = QFont(self.base_font)
            if st.font:
                f.setFamily(st.font)
            f.setPixelSize(max(4, int(round((st.size or 11) * 96 / 72 * self.zoom))))
            f.setBold(st.bold)
            f.setItalic(st.italic)
            f.setUnderline(st.underline)
            f.setStrikeOut(st.strike)
            self._fonts[key] = f
        return f

    # ------------------------------------------------------------ painting
    def paintEvent(self, ev):
        p = QPainter(self.viewport())
        p.setRenderHint(QPainter.TextAntialiasing, True)
        vw, vh = self.viewport().width(), self.viewport().height()
        p.fillRect(ev.rect(), QColor("#FFFFFF"))
        sh = self.sheet
        fr, fc = sh.freeze
        hw, hh = self.hdr_w, self.hdr_h
        fw, fh = self._frozen_px()
        self.cf.prepare(self.wb, sh)
        self._merged_painted = set()
        # column / row ranges per region
        c_scroll0 = self.ax_c.index_at(self.ax_c.pos(fc) + self.sx)
        c_scroll1 = self.ax_c.index_at(self.sx + (vw - hw) / self.zoom + 1)
        r_scroll0 = self.ax_r.index_at(self.ax_r.pos(fr) + self.sy)
        r_scroll1 = self.ax_r.index_at(self.sy + (vh - hh) / self.zoom + 1)
        c_scroll1 = min(COLS - 1, c_scroll1 + 1)
        r_scroll1 = min(ROWS - 1, r_scroll1 + 1)
        cols_f, rows_f = range(0, fc), range(0, fr)
        cols_s, rows_s = range(c_scroll0, c_scroll1 + 1), range(r_scroll0, r_scroll1 + 1)
        regions = [
            (QRect(hw, hh, int(fw), int(fh)), rows_f, cols_f),
            (QRect(hw + int(fw), hh, vw, int(fh)), rows_f, cols_s),
            (QRect(hw, hh + int(fh), int(fw), vh), rows_s, cols_f),
            (QRect(hw + int(fw), hh + int(fh), vw, vh), rows_s, cols_s),
        ]
        for clip, rows, cols in regions:
            if clip.width() <= 0 or clip.height() <= 0 or not len(rows) or not len(cols):
                continue
            c = clip.intersected(ev.rect())
            if c.isEmpty():
                continue
            p.save()
            p.setClipRect(c)
            self._paint_region(p, clip, rows, cols)
            p.restore()
        # frozen separators
        p.setPen(QPen(QColor("#998DA3"), 1.5))
        if fc:
            p.drawLine(int(hw + fw), hh, int(hw + fw), vh)
        if fr:
            p.drawLine(hw, int(hh + fh), vw, int(hh + fh))
        self._paint_headers(p, ev.rect(), cols_f, cols_s, rows_f, rows_s)
        self._paint_selection(p, regions)
        p.end()

    def _paint_region(self, p, clip, rows, cols):
        sh, wb = self.sheet, self.wb
        z = self.zoom
        # gridlines
        if sh.grid:
            p.setPen(QPen(GRID, 1))
            for c in cols:
                if self.ax_c.hidden(c):
                    continue
                x = int(self.col_x(c) + self.col_w(c)) - 0
                p.drawLine(x, clip.top(), x, clip.bottom())
            for r in rows:
                if self.ax_r.hidden(r):
                    continue
                y = int(self.row_y(r) + self.row_h(r))
                p.drawLine(clip.left(), y, clip.right(), y)
        has_cf = self.cf.has_rules()
        cells = sh.cells
        covered = self._covered
        pad = 3 * z
        for r in rows:
            h = self.row_h(r)
            if h <= 0:
                continue
            y = self.row_y(r)
            for c in cols:
                w = self.col_w(c)
                if w <= 0:
                    continue
                if (r, c) in covered:
                    continue
                cell = cells.get((r, c))
                if cell is None and not has_cf:
                    if (r, c) not in self._merge_map:
                        continue
                x = self.col_x(c)
                self._paint_cell(p, r, c, QRectF(x, y, w, h), cell, clip, cols, pad)
        # merged cells whose anchor is off screen but body is visible
        for (r1, c1, r2, c2) in sh.merges:
            if (r1, c1) in self._merged_painted:
                continue
            if r1 > rows[-1] or r2 < rows[0] or c1 > cols[-1] or c2 < cols[0]:
                continue
            rect = self.cell_rect(r1, c1)
            if rect.intersects(QRectF(clip)):
                self._paint_cell(p, r1, c1, rect, sh.cells.get((r1, c1)), clip, cols, pad, force=True)

    def _paint_cell(self, p, r, c, rect, cell, clip, cols, pad, force=False):
        sh, wb = self.sheet, self.wb
        merged = self.merge_at(r, c)
        if merged:
            if (r, c) != (merged[0], merged[1]) and not force:
                return
            rect = self.cell_rect(r, c)
            self._merged_painted.add((merged[0], merged[1]))
        if cell is None:
            st, v, text, sid = DEFAULT, None, "", 0
        else:
            raw, sid = cell
            st = wb.styles[sid]
            if self.show_formulas and isinstance(raw, str) and raw.startswith("="):
                text, v = raw, raw
            else:
                text, v, st = wb.display(sh, r, c)
        eff = None
        if self.cf.has_rules():
            eff = self.cf.apply(r, c, v if cell is not None else None)
        bg = st.bg
        fg = st.color
        if eff:
            bg = eff.get("bg", bg)
            fg = eff.get("fg", fg)
        z = self.zoom
        r_ = rect.adjusted(0, 0, 0, 0)
        if bg or merged:
            p.fillRect(r_.adjusted(0, 0, -1, -1) if bg else r_.adjusted(0, 0, -1, -1), QColor(bg) if bg else QColor("#FFFFFF"))
        if eff and eff.get("bar"):
            frac, col = eff["bar"]
            bc = QColor(col)
            bc.setAlpha(150)
            p.fillRect(QRectF(r_.left() + 2, r_.top() + 3 * z, max(0, (r_.width() - 4) * max(0, min(1, frac))), r_.height() - 6 * z), bc)
        if text != "":
            p.save()
            font = self.font_for(st)
            if eff and eff.get("bold") and not st.bold:
                font = QFont(font)
                font.setBold(True)
            p.setFont(font)
            p.setPen(QColor(fg) if fg else (QColor("#C0392B") if isinstance(v, F.XErr) else QColor("#1B1040")))
            h = st.h
            if h is None:
                h = "right" if isinstance(v, (int, float)) and not isinstance(v, bool) else ("center" if isinstance(v, (bool, F.XErr)) else "left")
            if self.show_formulas and isinstance(v, str) and str(v).startswith("="):
                h = "left"
            tr = r_.adjusted(pad + st.indent * 9 * z if h == "left" else pad, 0, -pad - (st.indent * 9 * z if h == "right" else 0), -1)
            fm = QFontMetricsF(font)
            va = {"top": Qt.AlignTop, "middle": Qt.AlignVCenter}.get(st.v, Qt.AlignBottom)
            if st.wrap and not isinstance(v, (int, float)):
                opt = QTextOption()
                opt.setWrapMode(QTextOption.WrapAtWordBoundaryOrAnywhere)
                opt.setAlignment({"left": Qt.AlignLeft, "center": Qt.AlignHCenter, "right": Qt.AlignRight}[h] | va)
                p.setClipRect(r_.adjusted(0, 0, -1, -1).toRect().intersected(p.clipBoundingRect().toRect()))
                p.drawText(tr, text, opt)
            else:
                tw = fm.horizontalAdvance(text)
                avail = tr.width()
                is_num = isinstance(v, (int, float)) and not isinstance(v, bool)
                if tw > avail:
                    if is_num and not merged:
                        text = "#" * max(1, int(avail / max(1.0, fm.horizontalAdvance("#"))))
                        tw = fm.horizontalAdvance(text)
                    elif h == "left" and not merged:
                        # overflow into empty cells on the right
                        extra, cc = 0.0, c + 1
                        last_col = cols[-1] if len(cols) else c
                        while cc <= last_col and tw > avail + extra and cc < COLS:
                            if (r, cc) in sh.cells and sh.cells[(r, cc)][0] not in (None, ""):
                                break
                            if (r, cc) in self._covered:
                                break
                            extra += self.col_w(cc)
                            cc += 1
                        if extra:
                            ext = QRectF(r_.left() + 1, r_.top() + 1, r_.width() + min(extra, tw - avail + pad), r_.height() - 2)
                            p.fillRect(QRectF(r_.right(), r_.top() + 1, ext.right() - r_.right(), r_.height() - 2), QColor(bg) if bg else QColor("#FFFFFF"))
                            tr = QRectF(tr.left(), tr.top(), ext.right() - tr.left(), tr.height())
                            p.setClipRect(ext.toRect().united(p.clipBoundingRect().toRect()).intersected(QRect(0, 0, 10 ** 6, 10 ** 6)))
                            avail = tr.width()
                    elif h == "right" and not merged:
                        pass
                align = {"left": Qt.AlignLeft, "center": Qt.AlignHCenter, "right": Qt.AlignRight}[h] | va | Qt.AlignVCenter if st.v == "middle" else \
                    {"left": Qt.AlignLeft, "center": Qt.AlignHCenter, "right": Qt.AlignRight}[h] | va
                if tw > avail and h != "left":
                    p.setClipRect(r_.adjusted(0, 0, -1, -1).toRect().intersected(p.clipBoundingRect().toRect()))
                elif tw > avail and h == "left" and "ext" not in locals():
                    p.setClipRect(r_.adjusted(0, 0, -1, -1).toRect().intersected(p.clipBoundingRect().toRect()))
                p.drawText(tr, int(align) | Qt.TextSingleLine, text)
            p.restore()
        # borders
        if st.bl or st.bt or st.br or st.bb:
            for col, a, b in ((st.bl, r_.topLeft(), r_.bottomLeft()), (st.bt, r_.topLeft(), r_.topRight()),
                              (st.br, r_.topRight(), r_.bottomRight()), (st.bb, r_.bottomLeft(), r_.bottomRight())):
                if col:
                    p.setPen(QPen(QColor(col), max(1, round(z))))
                    p.drawLine(a, b)
        if (r, c) in sh.comments:
            p.setBrush(QColor("#D64545"))
            p.setPen(Qt.NoPen)
            p.drawPolygon(QPolygonF([QPointF(r_.right() - 7 * z, r_.top()), QPointF(r_.right(), r_.top()), QPointF(r_.right(), r_.top() + 7 * z)]))
        f = sh.filter
        if f and f["range"][0] == r and f["range"][1] <= c <= f["range"][3]:
            bs = min(16 * z, r_.height() - 4)
            bx = QRectF(r_.right() - bs - 2, r_.bottom() - bs - 2, bs, bs)
            active = bool(f["hidden"].get(c))
            p.setPen(QPen(QColor("#C2395C") if active else QColor("#8A7FA6"), 1))
            p.setBrush(QColor("#E0D4E9") if active else QColor("#FFFFFF"))
            p.drawRect(bx)
            p.setBrush(QColor("#8A2150") if active else QColor("#534A5A"))
            p.setPen(Qt.NoPen)
            cx, cy = bx.center().x(), bx.center().y()
            p.drawPolygon(QPolygonF([QPointF(cx - 3.5 * z, cy - 1.5 * z), QPointF(cx + 3.5 * z, cy - 1.5 * z), QPointF(cx, cy + 3 * z)]))

    def _paint_headers(self, p, rect, cols_f, cols_s, rows_f, rows_s):
        hw, hh = self.hdr_w, self.hdr_h
        vw, vh = self.viewport().width(), self.viewport().height()
        z = self.zoom
        fr, fc = self.sheet.freeze
        fw, fh = self._frozen_px()
        # selected sets
        sel_cols, sel_rows, full_cols, full_rows = set(), set(), set(), set()
        for (r1, c1, r2, c2) in self.sel:
            if c2 - c1 < 400:
                sel_cols.update(range(c1, c2 + 1))
            if r2 - r1 < 400:
                sel_rows.update(range(r1, r2 + 1))
            if r1 == 0 and r2 >= ROWS - 1:
                full_cols.update(range(c1, c2 + 1))
            if c1 == 0 and c2 >= COLS - 1:
                full_rows.update(range(r1, r2 + 1))
        font = QFont(self.base_font)
        font.setPixelSize(max(6, int(12 * z)))
        p.setFont(font)
        p.save()
        p.setClipRect(QRect(hw, 0, vw, hh))
        p.fillRect(QRect(hw, 0, vw, hh), HDR_BG)
        for rng, scroll in ((cols_f, False), (cols_s, True)):
            for c in rng:
                w = self.col_w(c)
                if w <= 0:
                    continue
                x = self.col_x(c)
                if scroll and x + w < hw + fw:
                    continue
                if scroll:
                    p.setClipRect(QRect(int(hw + fw), 0, vw, hh))
                else:
                    p.setClipRect(QRect(hw, 0, int(fw), hh))
                rr = QRectF(x, 0, w, hh)
                if c in full_cols:
                    p.fillRect(rr, HDR_FULL)
                elif c in sel_cols:
                    p.fillRect(rr, HDR_SEL)
                p.setPen(QPen(QColor("#CEC8D2"), 1))
                p.drawLine(int(x + w), 0, int(x + w), hh)
                p.setPen(HDR_TXT_SEL if c in sel_cols else HDR_TXT)
                p.drawText(rr, Qt.AlignCenter, F.col_name(c))
        p.restore()
        p.save()
        p.setClipRect(QRect(0, hh, hw, vh))
        p.fillRect(QRect(0, hh, hw, vh), HDR_BG)
        for rng, scroll in ((rows_f, False), (rows_s, True)):
            for r in rng:
                h = self.row_h(r)
                if h <= 0:
                    continue
                y = self.row_y(r)
                if scroll and y + h < hh + fh:
                    continue
                if scroll:
                    p.setClipRect(QRect(0, int(hh + fh), hw, vh))
                else:
                    p.setClipRect(QRect(0, hh, hw, int(fh)))
                rr = QRectF(0, y, hw, h)
                if r in full_rows:
                    p.fillRect(rr, HDR_FULL)
                elif r in sel_rows:
                    p.fillRect(rr, HDR_SEL)
                p.setPen(QPen(QColor("#CEC8D2"), 1))
                p.drawLine(0, int(y + h), hw, int(y + h))
                p.setPen(HDR_TXT_SEL if r in sel_rows else HDR_TXT)
                p.drawText(rr, Qt.AlignCenter, str(r + 1))
        p.restore()
        p.fillRect(QRect(0, 0, hw, hh), HDR_BG)
        p.setPen(QPen(QColor("#CEC8D2"), 1))
        p.drawLine(0, hh - 1, vw, hh - 1)
        p.drawLine(hw - 1, 0, hw - 1, vh)
        p.setBrush(QColor("#BFB8C4"))
        p.setPen(Qt.NoPen)
        p.drawPolygon(QPolygonF([QPointF(hw - 4, hh - 4), QPointF(hw - 4, hh - 4 - 10 * z), QPointF(hw - 4 - 10 * z, hh - 4)]))

    def _clip_regions(self, regions):
        return [c for c, rows, cols in regions if c.width() > 0 and c.height() > 0]

    def _paint_selection(self, p, regions):
        clips = self._clip_regions(regions)
        z = self.zoom
        for ri, (r1, c1, r2, c2) in enumerate(self.sel):
            rect = self._range_rect(r1, c1, r2, c2)
            for clip in clips:
                p.save()
                p.setClipRect(clip)
                p.fillRect(rect, SEL_TINT)
                # the active cell stays clear
                if ri == len(self.sel) - 1 and not (r1 == r2 and c1 == c2) and not self.merge_at(*self.cur):
                    cr = self.cell_rect(*self.cur)
                    p.setCompositionMode(QPainter.CompositionMode_Source)
                    p.restore()
                    p.save()
                    p.setClipRect(clip)
                p.restore()
        # redraw the active cell without the tint
        for clip in clips:
            p.save()
            p.setClipRect(clip)
            if len(self.sel) and (self.sel[-1][0], self.sel[-1][1]) != (self.sel[-1][2], self.sel[-1][3]):
                cr = self.cell_rect(*self.cur)
                if cr.intersects(QRectF(clip)):
                    p.save()
                    p.setClipRect(QRectF(clip).intersected(cr).toRect())
                    r, c = self.cur
                    self._paint_cell(p, r, c, cr, self.sheet.cells.get((r, c)), clip, range(c, c + 1), 3 * z, force=True) if (r, c) in self.sheet.cells or self.cf.has_rules() else p.fillRect(cr.adjusted(0, 0, -1, -1), QColor("#FFFFFF"))
                    p.restore()
            p.restore()
        # outline of the last range and the fill handle
        lr = self.sel[-1]
        rect = self._range_rect(*lr)
        for clip in clips:
            p.save()
            p.setClipRect(clip)
            p.setPen(QPen(GREEN, max(2, 2 * z)))
            p.setBrush(Qt.NoBrush)
            p.drawRect(rect.adjusted(0, 0, -1, -1))
            p.restore()
        hx, hy = rect.right() - 1, rect.bottom() - 1
        p.setPen(QPen(QColor("#FFFFFF"), 1))
        p.setBrush(GREEN)
        p.drawRect(QRectF(hx - 3, hy - 3, 7, 7))
        dd = self.dropdown_rect()
        if dd is not None and self.editor is None:
            p.setPen(QPen(QColor("#AC9CB8"), 1))
            p.setBrush(QColor("#F3F8F5"))
            p.drawRect(dd)
            p.setPen(Qt.NoPen)
            p.setBrush(QColor("#342C3B"))
            cx, cy, a = dd.center().x(), dd.center().y(), min(dd.width(), dd.height()) * 0.22
            from PySide6.QtGui import QPolygonF
            p.drawPolygon(QPolygonF([QPointF(cx - a * 1.5, cy - a), QPointF(cx + a * 1.5, cy - a), QPointF(cx, cy + a)]))
        if self.copy_rect:
            rr = self._range_rect(*self.copy_rect)
            p.setBrush(Qt.NoBrush)
            p.setPen(QPen(QColor("#C2395C"), 1.5, Qt.DashLine))
            for clip in clips:
                p.save()
                p.setClipRect(clip)
                p.drawRect(rr.adjusted(0, 0, -1, -1))
                p.restore()
        if self._mode and self._mode[0] == "fill":
            rr = self._range_rect(*self._mode[2])
            p.setPen(QPen(QColor("#1B1040"), 1, Qt.DashLine))
            p.setBrush(Qt.NoBrush)
            p.drawRect(rr)

    def _range_rect(self, r1, c1, r2, c2):
        x1, y1 = self.col_x(c1), self.row_y(r1)
        if c2 >= COLS - 1 or r2 >= ROWS - 1:
            x2 = self.viewport().width() + 10 if c2 >= COLS - 1 else self.col_x(c2) + self.col_w(c2)
            y2 = self.viewport().height() + 10 if r2 >= ROWS - 1 else self.row_y(r2) + self.row_h(r2)
            if c2 >= COLS - 1 and c1 == 0:
                pass
            return QRectF(x1, y1, x2 - x1, y2 - y1)
        x2 = self.col_x(c2) + self.col_w(c2)
        y2 = self.row_y(r2) + self.row_h(r2)
        return QRectF(x1, y1, x2 - x1, y2 - y1)

    # ------------------------------------------------------------ selection
    def selected_range(self):
        return self.sel[-1]

    def set_selection(self, r1, c1, r2, c2, cur=None, add=False, emit=True):
        r1, c1, r2, c2 = norm(r1, c1, r2, c2)
        # grow over merged cells
        changed = True
        while changed:
            changed = False
            for m in self.sheet.merges:
                if m[0] <= r2 and m[2] >= r1 and m[1] <= c2 and m[3] >= c1:
                    n = (min(r1, m[0]), min(c1, m[1]), max(r2, m[2]), max(c2, m[3]))
                    if n != (r1, c1, r2, c2):
                        r1, c1, r2, c2 = n
                        changed = True
        if add:
            self.sel.append((r1, c1, r2, c2))
        else:
            self.sel = [(r1, c1, r2, c2)]
        if cur is not None:
            self.cur = cur
        self.viewport().update()
        if emit:
            self.selection_changed.emit()
            if self.sheet.dv:
                self.show_input_message()

    def go_to(self, r, c, extend=False, scroll=True):
        r = max(0, min(ROWS - 1, r))
        c = max(0, min(COLS - 1, c))
        if extend:
            ar, ac = self.anchor
            self.set_selection(ar, ac, r, c, cur=self.cur)
            self._ext = (r, c)
        else:
            m = self.merge_at(r, c)
            if m:
                r, c = m[0], m[1]
            self.anchor = (r, c)
            self.set_selection(r, c, r, c, cur=(r, c))
            self._ext = (r, c)
        if scroll:
            self.ensure_visible(*(self._ext if extend else (r, c)))
        self.viewport().update()

    def select_all(self):
        self.set_selection(0, 0, ROWS - 1, COLS - 1, cur=self.cur)

    def select_columns(self, c1, c2):
        self.anchor = (0, c1)
        self.set_selection(0, c1, ROWS - 1, c2, cur=(0, c1))

    def select_rows(self, r1, r2):
        self.anchor = (r1, 0)
        self.set_selection(r1, 0, r2, COLS - 1, cur=(r1, 0))

    def cells_in_selection(self, limit=2_000_000, fill_empty=True):
        """list of (r, c) in all selected ranges; huge selections are cut to the used area"""
        out = []
        ur, uc = self.wb.used_extent(self.sheet.name)
        for (r1, c1, r2, c2) in self.sel:
            if (r2 - r1 + 1) * (c2 - c1 + 1) > 20000:
                # whole rows / columns: only touch the area in use plus a screenful
                rr2 = min(r2, max(ur, r1) + (0 if (r2 - r1 > 5000 and r1 == 0) else 0))
                cc2 = min(c2, max(uc, c1))
                if r2 >= ROWS - 1 and r1 == 0 or c2 >= COLS - 1:
                    rr2 = min(r2, max(ur, r1 + 100))
                    cc2 = min(c2, max(uc, c1 + 25))
                r2, c2 = rr2, cc2
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    if (r, c) in self._covered:
                        continue
                    out.append((r, c))
                    if len(out) >= limit:
                        return out
        return out

    # ------------------------------------------------------------ mouse
    def _hit_header(self, pos):
        hw, hh = self.hdr_w, self.hdr_h
        if pos.x() < hw and pos.y() < hh:
            return "corner", None
        if pos.y() < hh:
            c = self.col_at(pos.x())
            x = self.col_x(c)
            if abs(pos.x() - (x + self.col_w(c))) <= 4:
                return "col-edge", c
            if abs(pos.x() - x) <= 4 and c > 0:
                nc = c - 1
                while nc > 0 and self.ax_c.hidden(nc):
                    nc -= 1
                return "col-edge", nc
            return "col", c
        if pos.x() < hw:
            r = self.row_at(pos.y())
            y = self.row_y(r)
            if abs(pos.y() - (y + self.row_h(r))) <= 4:
                return "row-edge", r
            if abs(pos.y() - y) <= 4 and r > 0:
                return "row-edge", r - 1
            return "row", r
        return "cell", None

    def _handle_rect(self):
        rect = self._range_rect(*self.sel[-1])
        return QRectF(rect.right() - 6, rect.bottom() - 6, 12, 12)

    def mousePressEvent(self, e):
        if self.editor is None:
            self.setFocus()
        pos = e.position().toPoint()
        if e.button() == Qt.RightButton:
            kind, idx = self._hit_header(pos)
            if kind == "cell":
                r, c = self.row_at(pos.y()), self.col_at(pos.x())
                if not any(a <= r <= c2_ and b <= c <= d for (a, b, c2_, d) in self.sel):
                    self.go_to(r, c)
                self.context_menu.emit(e.globalPosition().toPoint(), "cell")
            elif kind in ("col", "row"):
                self.context_menu.emit(e.globalPosition().toPoint(), kind)
            return
        if e.button() != Qt.LeftButton:
            return
        dd = self.dropdown_rect() if self.editor is None else None
        if dd is not None and dd.contains(QPointF(pos)):
            self.open_dropdown()
            return
        kind, idx = self._hit_header(pos)
        if self.editor is not None:
            if not self._can_point() or kind != "cell":
                self.setFocus()
            if self._can_point() and kind == "cell":
                r, c = self.row_at(pos.y()), self.col_at(pos.x())
                self._point_set(r, c, r, c)
                self._mode = ("point", r, c)
                self.editor.setFocus()
                return
            self.commit_edit()
        mod = e.modifiers()
        if kind == "corner":
            self.select_all()
        elif kind == "col-edge":
            self._mode = ("size-col", idx, pos.x(), self.ax_c.size(idx), self._size_targets("col", idx))
        elif kind == "row-edge":
            self._mode = ("size-row", idx, pos.y(), self.ax_r.size(idx), self._size_targets("row", idx))
        elif kind == "col":
            if mod & Qt.ShiftModifier:
                self.set_selection(0, self.anchor[1], ROWS - 1, idx, cur=self.cur)
            else:
                self.select_columns(idx, idx)
            self._mode = ("cols", idx)
        elif kind == "row":
            if mod & Qt.ShiftModifier:
                self.set_selection(self.anchor[0], 0, idx, COLS - 1, cur=self.cur)
            else:
                self.select_rows(idx, idx)
            self._mode = ("rows", idx)
        else:
            r, c = self.row_at(pos.y()), self.col_at(pos.x())
            if self._handle_rect().contains(QPointF(pos)) and len(self.sel) == 1:
                self._mode = ("fill", self.sel[-1], self.sel[-1])
                return
            f = self.sheet.filter
            if f and f["range"][0] == r and f["range"][1] <= c <= f["range"][3]:
                rect = self.cell_rect(r, c)
                if pos.x() >= rect.right() - 20 * self.zoom and pos.y() >= rect.bottom() - 20 * self.zoom:
                    self.filter_clicked.emit(c, self.viewport().mapToGlobal(QPoint(int(rect.left()), int(rect.bottom()))))
                    return
            if mod & Qt.ShiftModifier:
                self.go_to(r, c, extend=True)
            elif mod & Qt.ControlModifier:
                m = self.merge_at(r, c)
                if m:
                    r, c = m[0], m[1]
                self.anchor = (r, c)
                self.set_selection(r, c, r, c, cur=(r, c), add=True)
            else:
                self.go_to(r, c, scroll=False)
            self._mode = ("cells", r, c)
            self.auto.start()
        self.viewport().update()

    def _size_targets(self, kind, idx):
        if kind == "col":
            for (r1, c1, r2, c2) in self.sel:
                if r1 == 0 and r2 >= ROWS - 1 and c1 <= idx <= c2:
                    return [c for c in range(c1, min(c2, c1 + 2000) + 1)]
        else:
            for (r1, c1, r2, c2) in self.sel:
                if c1 == 0 and c2 >= COLS - 1 and r1 <= idx <= r2:
                    return [r for r in range(r1, min(r2, r1 + 5000) + 1)]
        return [idx]

    def _can_point(self):
        ed = self.editor
        if ed is None or not ed.text().startswith("="):
            return False
        if self._pt is not None:
            return True
        t = ed.text()[:ed.cursorPosition()].rstrip()
        return bool(t) and t[-1] in OPS

    def mouseMoveEvent(self, e):
        pos = e.position().toPoint()
        m = self._mode
        if m is None:
            kind, idx = self._hit_header(pos)
            if kind == "col-edge":
                self.viewport().setCursor(Qt.SplitHCursor)
            elif kind == "row-edge":
                self.viewport().setCursor(Qt.SplitVCursor)
            elif kind == "cell" and self._handle_rect().contains(QPointF(pos)):
                self.viewport().setCursor(Qt.CrossCursor)
            elif kind in ("col", "row"):
                self.viewport().setCursor(Qt.PointingHandCursor)
            else:
                self.viewport().setCursor(Qt.ArrowCursor)
            if kind == "cell":
                r, c = self.row_at(pos.y()), self.col_at(pos.x())
                t = self.sheet.comments.get((r, c))
                if t:
                    QToolTip.showText(e.globalPosition().toPoint(), t, self.viewport())
            return
        if not (e.buttons() & Qt.LeftButton):
            return
        self._last_pos = pos
        if m[0] == "size-col":
            _, idx, x0, s0, targets = m
            ns = max(6, s0 + (pos.x() - x0) / self.zoom)
            for t in targets:
                self.ax_c.custom[t] = ns
            self.ax_c._rebuild()
            self._after_resize()
        elif m[0] == "size-row":
            _, idx, y0, s0, targets = m
            ns = max(6, s0 + (pos.y() - y0) / self.zoom)
            for t in targets:
                self.ax_r.custom[t] = ns
            self.ax_r._rebuild()
            self._after_resize()
        elif m[0] == "cols":
            self.set_selection(0, self.anchor[1], ROWS - 1, self.col_at(pos.x()), cur=self.cur)
        elif m[0] == "rows":
            self.set_selection(self.anchor[0], 0, self.row_at(pos.y()), COLS - 1, cur=self.cur)
        elif m[0] == "cells":
            r, c = self.row_at(pos.y()), self.col_at(pos.x())
            self.go_to(r, c, extend=True, scroll=False)
        elif m[0] == "point":
            r, c = self.row_at(pos.y()), self.col_at(pos.x())
            self._point_set(m[1], m[2], r, c)
        elif m[0] == "fill":
            r, c = self.row_at(pos.y()), self.col_at(pos.x())
            r1, c1, r2, c2 = m[1]
            if abs(r - (r2 if r > r2 else r1)) >= abs(c - (c2 if c > c2 else c1)) * 0 and (r > r2 or r < r1 or (c1 <= c <= c2)):
                tgt = (min(r1, r), c1, max(r2, r), c2)
            else:
                tgt = (r1, min(c1, c), r2, max(c2, c))
            # choose direction by distance outside the source
            dv = (r - r2) if r > r2 else (r1 - r) if r < r1 else 0
            dh = (c - c2) if c > c2 else (c1 - c) if c < c1 else 0
            tgt = (min(r1, r), c1, max(r2, r), c2) if dv >= dh else (r1, min(c1, c), r2, max(c2, c))
            self._mode = ("fill", m[1], tgt)
        self.viewport().update()

    def _after_resize(self):
        self._update_scrollbars()
        self._place_editor()
        self._place_floating()

    def mouseReleaseEvent(self, e):
        m = self._mode
        self._mode = None
        self.auto.stop()
        if m is None:
            return
        if m[0] in ("size-col", "size-row"):
            sh = self.sheet
            if m[0] == "size-col":
                new = {k: v for k, v in self.ax_c.custom.items()}
                # keep hidden flags separate from sizes
                sizes = {k: v for k, v in new.items() if v > 0 or k not in sh.hidden_cols}
                self.wb.set_prop(sh.name, "colw", {k: v for k, v in sizes.items() if v > 0}, "Column width")
            else:
                new = {k: v for k, v in self.ax_r.custom.items() if v > 0}
                self.wb.set_prop(sh.name, "rowh", new, "Row height")
            self.axes_changed.emit()
        elif m[0] == "fill":
            src, tgt = m[1], m[2]
            if tgt != src:
                self.fill_range(src, tgt)
        elif m[0] == "point":
            pass
        self.viewport().update()

    def mouseDoubleClickEvent(self, e):
        pos = e.position().toPoint()
        kind, idx = self._hit_header(pos)
        if kind == "col-edge":
            self.autofit_columns(self._size_targets("col", idx))
            return
        if kind == "row-edge":
            self.autofit_rows(self._size_targets("row", idx))
            return
        if kind == "cell":
            if self._handle_rect().contains(QPointF(pos)) and len(self.sel) == 1:
                self.autofill_down()
                return
            self.begin_edit()

    def _autoscroll(self):
        if self._mode and self._mode[0] in ("cells", "cols", "rows", "fill"):
            pos = self.viewport().mapFromGlobal(QCursor.pos())
            dx = dy = 0
            if pos.x() > self.viewport().width():
                dx = 40
            elif pos.x() < self.hdr_w:
                dx = -40
            if pos.y() > self.viewport().height():
                dy = 40
            elif pos.y() < self.hdr_h:
                dy = -40
            if dx or dy:
                hb, vb = self.horizontalScrollBar(), self.verticalScrollBar()
                hb.setValue(hb.value() + dx)
                vb.setValue(vb.value() + dy)
                ev = type("E", (), {})()
                self.mouseMoveEvent(_FakeMove(self, pos))

    def contextMenuEvent(self, e):
        e.accept()

    # ------------------------------------------------------------ keyboard
    def keyPressEvent(self, e):
        k, mod = e.key(), e.modifiers()
        ctrl = bool(mod & Qt.ControlModifier)
        shift = bool(mod & Qt.ShiftModifier)
        r, c = self._ext if (shift and hasattr(self, "_ext")) else self.cur
        if not shift:
            r, c = self.cur
        nav = {Qt.Key_Up: (-1, 0), Qt.Key_Down: (1, 0), Qt.Key_Left: (0, -1), Qt.Key_Right: (0, 1)}
        if k == Qt.Key_Down and (mod & Qt.AltModifier) and self.dropdown_rect() is not None:
            self.open_dropdown()
            return
        if k in nav:
            dr, dc = nav[k]
            if ctrl:
                r, c = self._jump(r, c, dr, dc)
            else:
                m = self.merge_at(r, c) if not shift else None
                if m:
                    r2, c2 = (m[2] if dr > 0 else m[0], m[3] if dc > 0 else m[1])
                    r, c = r2, c2
                r, c = self._step(r, c, dr, dc)
            self.go_to(r, c, extend=shift)
            return
        if k in (Qt.Key_Tab, Qt.Key_Backtab):
            self.move_in_selection(0, -1 if (shift or k == Qt.Key_Backtab) else 1)
            return
        if k in (Qt.Key_Return, Qt.Key_Enter):
            self.move_in_selection(-1 if shift else 1, 0)
            return
        if k == Qt.Key_PageDown or k == Qt.Key_PageUp:
            n = max(1, int((self.viewport().height() - self.hdr_h) / (DEFAULT_H * self.zoom)) - 1) * (1 if k == Qt.Key_PageDown else -1)
            if mod & Qt.AltModifier:
                self.go_to(self.cur[0], self.cur[1] + n // 2, extend=shift)
            else:
                self.go_to(self.cur[0] + n, self.cur[1], extend=shift)
            return
        if k == Qt.Key_Home:
            self.go_to(0 if ctrl else self.cur[0], 0, extend=shift)
            return
        if k == Qt.Key_End:
            if ctrl:
                ur, uc = self.wb.used_extent(self.sheet.name)
                self.go_to(max(0, ur), max(0, uc), extend=shift)
            else:
                ur, uc = self.wb.used_extent(self.sheet.name)
                self.go_to(self.cur[0], max(0, uc), extend=shift)
            return
        if k == Qt.Key_Space and ctrl:
            self.select_columns(self.cur[1], self.cur[1])
            return
        if k == Qt.Key_Space and shift:
            self.select_rows(self.cur[0], self.cur[0])
            return
        if k == Qt.Key_F2:
            self.begin_edit()
            return
        if k in (Qt.Key_Delete, Qt.Key_Backspace):
            if k == Qt.Key_Backspace and not ctrl:
                self.begin_edit(initial="")
                return
            self.clear_contents()
            return
        if k == Qt.Key_Escape:
            self.copy_rect = None
            self._cut = None
            self.viewport().update()
            return
        if ctrl:
            if k == Qt.Key_Z:
                (self.redo_requested if shift else self.undo_requested).emit()
            elif k == Qt.Key_Y:
                self.redo_requested.emit()
            elif k == Qt.Key_A:
                self.select_all()
            elif k == Qt.Key_C:
                self.copy()
            elif k == Qt.Key_X:
                self.copy(cut=True)
            elif k == Qt.Key_V:
                self.paste("values" if shift else "all")
            elif k == Qt.Key_D:
                self.fill_edge("down")
            elif k == Qt.Key_R:
                self.fill_edge("right")
            elif k == Qt.Key_Semicolon:
                import datetime
                from .numfmt import dt_to_serial
                self.wb.set_values(self.sheet.name, [(self.cur[0], self.cur[1], datetime.date.today().isoformat())])
            else:
                super().keyPressEvent(e)
            return
        t = e.text()
        if t and t.isprintable() and not (mod & (Qt.ControlModifier | Qt.AltModifier | Qt.MetaModifier)):
            self.begin_edit(initial=t)
            return
        super().keyPressEvent(e)

    def _step(self, r, c, dr, dc):
        r += dr
        c += dc
        if dr:
            r = self.ax_r.nearest_visible(r, dr)
        if dc:
            c = self.ax_c.nearest_visible(c, dc)
        return max(0, min(ROWS - 1, r)), max(0, min(COLS - 1, c))

    def _jump(self, r, c, dr, dc):
        cells = self.sheet.cells

        def filled(rr, cc):
            cell = cells.get((rr, cc))
            return cell is not None and cell[0] not in (None, "")
        ur, uc = self.wb.used_extent(self.sheet.name)
        nr, nc = r + dr, c + dc
        if not (0 <= nr < ROWS and 0 <= nc < COLS):
            return r, c
        if filled(r, c) and filled(nr, nc):
            while 0 <= nr + dr < ROWS and 0 <= nc + dc < COLS and filled(nr + dr, nc + dc) and (nr <= ur + 1 and nc <= uc + 1):
                nr, nc = nr + dr, nc + dc
            return nr, nc
        while 0 <= nr < ROWS and 0 <= nc < COLS and not filled(nr, nc):
            if (dr > 0 and nr > ur) or (dc > 0 and nc > uc):
                return (ROWS - 1 if dr > 0 else nr, COLS - 1 if dc > 0 else nc)
            nr, nc = nr + dr, nc + dc
        return max(0, min(ROWS - 1, nr)), max(0, min(COLS - 1, nc))

    def move_in_selection(self, dr, dc):
        """Enter / Tab: move inside the selected block (wraps), or step when only one cell"""
        r1, c1, r2, c2 = self.sel[-1]
        r, c = self.cur
        if (r1, c1) == (r2, c2) or len(self.sel) > 1 and False:
            nr, nc = self._step(r, c, dr, dc)
            self.go_to(nr, nc)
            return
        if dr:
            r += dr
            if r > r2:
                r, c = r1, c + 1
            elif r < r1:
                r, c = r2, c - 1
            if c > c2:
                c = c1
            elif c < c1:
                c = c2
        else:
            c += dc
            if c > c2:
                c, r = c1, r + 1
            elif c < c1:
                c, r = c2, r - 1
            if r > r2:
                r = r1
            elif r < r1:
                r = r2
        self.cur = (r, c)
        self.ensure_visible(r, c)
        self.viewport().update()
        self.selection_changed.emit()

    # ------------------------------------------------------------ editing
    def begin_edit(self, initial=None):
        if self.editor is not None:
            return
        r, c = self.cur
        m = self.merge_at(r, c)
        if m:
            r, c = m[0], m[1]
        self.ensure_visible(r, c)
        ed = CellEditor(self)
        st = self.wb.style_of(self.sheet, r, c)
        ed.setFont(self.font_for(st))
        ed.mode = "edit" if initial is None else "enter"
        ed.setText(self.wb.edit_text(self.sheet.name, r, c) if initial is None else initial)
        ed.key.connect(self._editor_key)
        ed.textChanged.connect(self._editor_text)
        self.editor = ed
        self.edit_origin = (r, c)
        self._pt = None
        self._pt_cell = None
        self._place_editor()
        ed.show()
        ed.setFocus()
        if initial is None:
            ed.deselect()
            ed.setCursorPosition(len(ed.text()))
        self.edit_state.emit(True)
        self.edit_text_changed.emit(ed.text())

    def _editor_text(self, t):
        self._grow_editor()
        self.edit_text_changed.emit(t)

    def set_editor_text(self, t):
        if self.editor is not None and self.editor.text() != t:
            self.editor.blockSignals(True)
            self.editor.setText(t)
            self.editor.blockSignals(False)
            self._grow_editor()

    def _grow_editor(self):
        if self.editor is None:
            return
        self._place_editor()

    def _place_editor(self):
        ed = self.editor
        if ed is None:
            return
        r, c = self.edit_origin
        rect = self.cell_rect(r, c)
        fm = ed.fontMetrics()
        want = fm.horizontalAdvance(ed.text()) + 24
        w = max(rect.width(), min(want, self.viewport().width() - rect.left() - 4))
        ed.setGeometry(int(rect.left()) - 1, int(rect.top()) - 1, int(w) + 2, int(rect.height()) + 2)

    def commit_edit(self, move=None):
        ed = self.editor
        if ed is None:
            return
        text = ed.text()
        r, c = self.edit_origin
        self._close_editor()
        sh = self.sheet.name
        cur_raw = self.wb.edit_text(sh, r, c)
        if text != cur_raw:
            if not self._validated(r, c, text):
                return
            self.wb.set_values(sh, [(r, c, text)], "Edit cell")
            self._maybe_fit_wrap(r, c)
        if move:
            self.move_in_selection(*move)
        self.viewport().update()

    def _validated(self, r, c, text):
        """data validation on typed input; False = value refused (the cell is left as it was or reopened for a retry)"""
        rule = V.check(self.wb, self.sheet, r, c, text)
        if rule is None:
            return True
        from PySide6.QtWidgets import QMessageBox
        title = rule.get("title") or "Invalid value"
        msg = rule.get("msg") or V.default_message(rule)
        style = rule.get("style", "stop")
        box = QMessageBox(self)
        box.setWindowTitle(title)
        box.setText(msg)
        if style == "stop":
            box.setIcon(QMessageBox.Critical)
            retry = box.addButton("Retry", QMessageBox.AcceptRole)
            box.addButton("Cancel", QMessageBox.RejectRole)
            box.exec()
            if box.clickedButton() is retry:
                self.go_to(r, c)
                self.begin_edit(initial=text)
            return False
        box.setIcon(QMessageBox.Warning if style == "warning" else QMessageBox.Information)
        box.setInformativeText("Continue anyway?" if style == "warning" else "")
        yes = box.addButton("Yes" if style == "warning" else "OK", QMessageBox.AcceptRole)
        box.addButton("No" if style == "warning" else "Cancel", QMessageBox.RejectRole)
        box.exec()
        return box.clickedButton() is yes

    def dropdown_rect(self):
        """small arrow button at the right edge of the active cell when it has a list rule"""
        r, c = self.cur
        rule = V.rule_at(self.sheet, r, c)
        if rule is None or rule["kind"] != "list" or not rule.get("dropdown", True):
            return None
        cr = self.cell_rect(r, c)
        h = min(cr.height() - 2, 20 * self.zoom)
        w = min(cr.width() / 2, 18 * self.zoom)
        return QRectF(cr.right() - w - 1, cr.bottom() - h - 1, w, h)

    def open_dropdown(self):
        r, c = self.cur
        rule = V.rule_at(self.sheet, r, c)
        if rule is None or rule["kind"] != "list":
            return
        items = V.list_items(self.wb, self.sheet, rule, r, c)
        menu = QMenu(self)
        cur = self.wb.display(self.sheet.name, r, c)
        cur = cur[0] if isinstance(cur, tuple) else cur
        for it in items[:500]:
            a = menu.addAction(it)
            a.setCheckable(True)
            a.setChecked(it == cur)
            a.triggered.connect(lambda checked=False, it=it: self.wb.set_values(self.sheet.name, [(r, c, it)], "Choose from list"))
        if not items:
            menu.addAction("(the list is empty)").setEnabled(False)
        cr = self.cell_rect(r, c)
        menu.exec(self.viewport().mapToGlobal(QPoint(int(cr.left()), int(cr.bottom()))))
        self.viewport().update()

    def show_input_message(self):
        r, c = self.cur
        rule = V.rule_at(self.sheet, r, c)
        if rule and (rule.get("in_title") or rule.get("in_msg")):
            cr = self.cell_rect(r, c)
            txt = (f"<b>{rule['in_title']}</b><br>" if rule.get("in_title") else "") + rule.get("in_msg", "").replace("\n", "<br>")
            QToolTip.showText(self.viewport().mapToGlobal(QPoint(int(cr.left()), int(cr.bottom()) + 4)), txt, self.viewport())

    def _maybe_fit_wrap(self, r, c):
        st = self.wb.style_of(self.sheet, r, c)
        if st.wrap and r not in self.sheet.rowh:
            self.autofit_rows([r])

    def cancel_edit(self):
        if self.editor is not None:
            self._close_editor()
            self.viewport().update()

    def _close_editor(self):
        ed = self.editor
        self.editor = None
        self._pt = None
        ed.hide()
        ed.deleteLater()
        self.edit_state.emit(False)
        self.setFocus()

    def _editor_key(self, e):
        k = e.key()
        shift = bool(e.modifiers() & Qt.ShiftModifier)
        ed = self.editor
        if self.key_filter is not None and self.key_filter(e):
            return
        if k == Qt.Key_Escape:
            self.cancel_edit()
        elif k in (Qt.Key_Return, Qt.Key_Enter):
            if e.modifiers() & Qt.AltModifier:
                ed.insert("\n")
                return
            self.commit_edit(move=(-1 if shift else 1, 0))
        elif k in (Qt.Key_Tab, Qt.Key_Backtab):
            self.commit_edit(move=(0, -1 if (shift or k == Qt.Key_Backtab) else 1))
        elif k == Qt.Key_F4:
            self._toggle_abs()
        elif k in (Qt.Key_Up, Qt.Key_Down, Qt.Key_Left, Qt.Key_Right):
            dr = {Qt.Key_Up: -1, Qt.Key_Down: 1}.get(k, 0)
            dc = {Qt.Key_Left: -1, Qt.Key_Right: 1}.get(k, 0)
            if self._can_point() and ed.mode == "enter":
                base = self._pt_cell or self.edit_origin
                nr, nc = self._step(base[0], base[1], dr, dc)
                self._point_set(nr, nc, nr, nc)
                self._pt_cell = (nr, nc)
                self.ensure_visible(nr, nc)
            else:
                self.commit_edit()
                r, c = self.cur
                nr, nc = self._step(r, c, dr, dc)
                self.go_to(nr, nc, extend=False)
        elif k in (Qt.Key_PageUp, Qt.Key_PageDown):
            self.commit_edit()
            self.keyPressEvent(e)

    def _toggle_abs(self):
        ed = self.editor
        t = ed.text()
        pos = ed.cursorPosition()
        for m in re.finditer(r"(\$?)([A-Za-z]{1,3})(\$?)(\d+)", t):
            if m.start() <= pos <= m.end():
                a, b = bool(m.group(1)), bool(m.group(3))
                states = [(False, False), (True, True), (False, True), (True, False)]
                nxt = states[(states.index((a, b)) + 1) % 4] if (a, b) in states else (True, True)
                new = ("$" if nxt[0] else "") + m.group(2) + ("$" if nxt[1] else "") + m.group(4)
                ed.setText(t[:m.start()] + new + t[m.end():])
                ed.setCursorPosition(m.start() + len(new))
                return

    def _point_set(self, r1, c1, r2, c2):
        ed = self.editor
        if ed is None:
            return
        r1, c1, r2, c2 = norm(r1, c1, r2, c2)
        txt = F.cell_name(r1, c1) if (r1, c1) == (r2, c2) else F.cell_name(r1, c1) + ":" + F.cell_name(r2, c2)
        t = ed.text()
        if self._pt is None:
            pos = ed.cursorPosition()
            self._pt = (pos, pos)
        s, en = self._pt
        ed.setText(t[:s] + txt + t[en:])
        self._pt = (s, s + len(txt))
        ed.setCursorPosition(s + len(txt))
        self._pt_cell = (r2, c2)

    # ------------------------------------------------------------ data operations
    def clear_contents(self):
        cells = self.cells_in_selection()
        ch = {}
        for (r, c) in cells:
            cur = self.sheet.cells.get((r, c))
            if cur is not None and cur[0] is not None:
                ch[(self.sheet.name, r, c)] = (None, cur[1]) if cur[1] else None
        if ch:
            self.wb.apply(ch, "Clear")

    def clear_all(self):
        ch = {(self.sheet.name, r, c): None for (r, c) in self.cells_in_selection() if (r, c) in self.sheet.cells}
        if ch:
            self.wb.apply(ch, "Clear all")

    def clear_formats(self):
        ch = {}
        for (r, c) in self.cells_in_selection():
            cur = self.sheet.cells.get((r, c))
            if cur and cur[1]:
                ch[(self.sheet.name, r, c)] = (cur[0], 0) if cur[0] is not None else None
        if ch:
            self.wb.apply(ch, "Clear formats")

    def _style_dict(self, sid):
        return list(self.wb.styles[sid])

    def copy(self, cut=False):
        r1, c1, r2, c2 = self.sel[-1]
        ur, uc = self.wb.used_extent(self.sheet.name)
        r2, c2 = min(r2, max(ur, r1)), min(c2, max(uc, c1))
        if (r2 - r1 + 1) * (c2 - c1 + 1) > 2_000_000:
            self.message.emit("Selection is too big to copy")
            return
        rows, text_rows, html = [], [], ["<table border='0' cellspacing='0' cellpadding='3'>"]
        for r in range(r1, r2 + 1):
            row, trow = [], []
            html.append("<tr>")
            for c in range(c1, c2 + 1):
                cell = self.sheet.cells.get((r, c))
                text, v, st = self.wb.display(self.sheet, r, c)
                if cell:
                    row.append([cell[0], self._style_dict(cell[1])])
                else:
                    row.append(None)
                trow.append(text.replace("\t", " ").replace("\n", " "))
                css = ("font-weight:bold;" if st.bold else "") + ("font-style:italic;" if st.italic else "") + (f"background:{st.bg};" if st.bg else "") + (f"color:{st.color};" if st.color else "")
                html.append(f"<td style='{css}'>{_esc(text)}</td>")
            html.append("</tr>")
            rows.append(row)
            text_rows.append("\t".join(trow))
        html.append("</table>")
        text = "\n".join(text_rows)
        md = QMimeData()
        md.setText(text)
        md.setHtml("".join(html))
        payload = {"sheet": self.sheet.name, "origin": [r1, c1], "rows": rows, "text": text}
        md.setData("application/x-digonto-cells", json.dumps(payload).encode("utf-8"))
        QGuiApplication.clipboard().setMimeData(md)
        self.copy_rect = (r1, c1, r2, c2)
        self._cut = (self.sheet.name, (r1, c1, r2, c2)) if cut else None
        self.viewport().update()

    def paste(self, mode="all"):
        md = QGuiApplication.clipboard().mimeData()
        if md is None:
            return
        r0, c0, r1x, c1x = self.sel[-1]
        sh = self.sheet.name
        ch = {}
        payload = None
        if md.hasFormat("application/x-digonto-cells"):
            try:
                payload = json.loads(bytes(md.data("application/x-digonto-cells")).decode("utf-8"))
                if md.hasText() and md.text() != payload.get("text"):
                    payload = None
            except ValueError:
                payload = None
        if payload:
            rows = payload["rows"]
            h, w = len(rows), len(rows[0]) if rows else 0
            if mode == "transpose":
                rows = [list(x) for x in zip(*rows)]
                h, w = len(rows), len(rows[0]) if rows else 0
            if not h or not w:
                return
            sr0, sc0 = payload["origin"]
            tr = (r1x - r0 + 1) // h if (r1x - r0 + 1) % h == 0 and (r1x - r0 + 1) >= h and r1x < ROWS - 1 else 1
            tc = (c1x - c0 + 1) // w if (c1x - c0 + 1) % w == 0 and (c1x - c0 + 1) >= w and c1x < COLS - 1 else 1
            if self._cut and self._cut[0] == sh:
                a, b, c_, d = self._cut[1]
                for r in range(a, c_ + 1):
                    for cc in range(b, d + 1):
                        if (r, cc) in self.sheet.cells:
                            ch[(sh, r, cc)] = None
            for ty in range(tr):
                for tx in range(tc):
                    for i in range(h):
                        for j in range(w):
                            r, c = r0 + ty * h + i, c0 + tx * w + j
                            if r >= ROWS or c >= COLS:
                                continue
                            item = rows[i][j]
                            cur = self.sheet.cells.get((r, c))
                            raw = item[0] if item else None
                            if mode == "transpose":
                                src_r, src_c = sr0 + j, sc0 + i
                            else:
                                src_r, src_c = sr0 + i, sc0 + j
                            if isinstance(raw, str) and raw.startswith("="):
                                if self._cut:
                                    pass
                                else:
                                    raw = F.shift_formula(raw, r - src_r, c - src_c)
                            if mode == "values":
                                if isinstance(raw, str) and raw.startswith("="):
                                    raw = self.wb.cell_value(payload["sheet"], src_r, src_c) if self.wb.sheet(payload["sheet"]) else raw
                                    raw = None if isinstance(raw, F.XErr) and False else (raw if not isinstance(raw, F.XErr) else raw.code)
                                sid = cur[1] if cur else 0
                            elif mode == "formats":
                                raw = cur[0] if cur else None
                                sid = self.wb.style_id(Style(*item[1])) if item else 0
                            else:
                                sid = self.wb.style_id(Style(*item[1])) if item else 0
                            if mode == "formulas" and not (isinstance(raw, str) and str(raw).startswith("=")) and item is None:
                                pass
                            if mode == "formulas":
                                sid = cur[1] if cur else 0
                            ch[(sh, r, c)] = (raw, sid) if (raw is not None or sid) else None
            end = (r0 + h * tr - 1, c0 + w * tc - 1)
        else:
            text = md.text() if md.hasText() else ""
            if not text:
                return
            lines = text.replace("\r\n", "\n").replace("\r", "\n").rstrip("\n").split("\n")
            from .model import parse_input
            end = (r0, c0)
            for i, line in enumerate(lines):
                for j, val in enumerate(line.split("\t")):
                    r, c = r0 + i, c0 + j
                    if r >= ROWS or c >= COLS:
                        continue
                    raw, fmt = parse_input(val)
                    cur = self.sheet.cells.get((r, c))
                    st = self.wb.styles[cur[1]] if cur else DEFAULT
                    if fmt and st.fmt == "General":
                        st = st._replace(fmt=fmt)
                    ch[(sh, r, c)] = (raw, self.wb.style_id(st)) if (raw is not None or st != DEFAULT) else None
                    end = (max(end[0], r), max(end[1], c))
        if ch:
            self.wb.apply(ch, "Paste")
            self.set_selection(r0, c0, end[0], end[1], cur=(r0, c0))
            if self._cut:
                self._cut = None
                self.copy_rect = None
        self.viewport().update()

    # fill ------------------------------------------------------------
    def fill_edge(self, direction):
        r1, c1, r2, c2 = self.sel[-1]
        if direction == "down":
            if r1 == r2:
                return
            self.fill_range((r1, c1, r1, c2), (r1, c1, r2, c2), single_copy=True)
        else:
            if c1 == c2:
                return
            self.fill_range((r1, c1, r2, c1), (r1, c1, r2, c2), single_copy=True)

    def autofill_down(self):
        r1, c1, r2, c2 = self.sel[-1]
        cells = self.sheet.cells
        # follow the neighbouring column to the left (or right)
        for nc in (c1 - 1, c2 + 1):
            if nc < 0:
                continue
            n = r2
            while (n + 1, nc) in cells and cells[(n + 1, nc)][0] not in (None, ""):
                n += 1
            if n > r2:
                self.fill_range((r1, c1, r2, c2), (r1, c1, n, c2))
                return

    def fill_range(self, src, tgt, single_copy=False):
        """extend the source block over the target block (copy / series)"""
        sr1, sc1, sr2, sc2 = src
        tr1, tc1, tr2, tc2 = tgt
        sh = self.sheet.name
        ch = {}
        vertical = (tc1 == sc1 and tc2 == sc2) and (tr1 != sr1 or tr2 != sr2)
        if vertical:
            for c in range(sc1, sc2 + 1):
                col = [self.sheet.cells.get((r, c)) for r in range(sr1, sr2 + 1)]
                fwd = tr2 > sr2
                targets = list(range(sr2 + 1, tr2 + 1)) if fwd else list(range(sr1 - 1, tr1 - 1, -1))
                vals = _fill_series(col, len(targets), fwd, single_copy, [sr1 + i for i in range(len(col))], targets, True, self.wb)
                for t, cell in zip(targets, vals):
                    ch[(sh, t, c)] = cell
        else:
            for r in range(sr1, sr2 + 1):
                row = [self.sheet.cells.get((r, c)) for c in range(sc1, sc2 + 1)]
                fwd = tc2 > sc2
                targets = list(range(sc2 + 1, tc2 + 1)) if fwd else list(range(sc1 - 1, tc1 - 1, -1))
                vals = _fill_series(row, len(targets), fwd, single_copy, [sc1 + i for i in range(len(row))], targets, False, self.wb)
                for t, cell in zip(targets, vals):
                    ch[(sh, r, t)] = cell
        if ch:
            self.wb.apply(ch, "Fill")
            self.set_selection(*norm(*tgt), cur=self.cur)

    # sizing ------------------------------------------------------------
    def autofit_columns(self, cols):
        sh = self.sheet
        new = dict(sh.colw)
        cset = set(cols)
        best = {c: 0.0 for c in cols}
        count = 0
        for (r, c), cell in sh.cells.items():
            if c in cset and cell[0] is not None:
                text, v, st = self.wb.display(sh, r, c)
                if not text:
                    continue
                fm = QFontMetricsF(self.font_for(st._replace(size=st.size)) if self.zoom == 1 else self._unzoomed_font(st))
                w = fm.horizontalAdvance(text.split("\n")[0]) + 14 + st.indent * 9
                best[c] = max(best[c], w)
                count += 1
                if count > 200000:
                    break
        for c in cols:
            new[c] = max(30, min(600, best[c] or DEFAULT_W))
        self.wb.set_prop(sh.name, "colw", new, "Auto-fit column")
        self.reload_axes()
        self._place_editor()
        self.viewport().update()
        self.axes_changed.emit()

    def _unzoomed_font(self, st):
        f = QFont(self.base_font)
        if st.font:
            f.setFamily(st.font)
        f.setPixelSize(max(4, int(round((st.size or 11) * 96 / 72))))
        f.setBold(st.bold)
        f.setItalic(st.italic)
        return f

    def autofit_rows(self, rows):
        sh = self.sheet
        new = dict(sh.rowh)
        for r in rows:
            best = DEFAULT_H
            for (rr, c), cell in sh.cells.items() if len(rows) > 50 else ((k, v) for k, v in sh.cells.items() if k[0] == r):
                if rr != r:
                    continue
                text, v, st = self.wb.display(sh, r, c)
                px = (st.size or 11) * 96 / 72
                lines = 1
                if st.wrap and text:
                    fm = QFontMetricsF(self._unzoomed_font(st))
                    wd = max(10, self.ax_c.size(c) - 6)
                    br = fm.boundingRect(QRectF(0, 0, wd, 10 ** 6), int(Qt.TextWordWrap), text)
                    best = max(best, br.height() + 8)
                else:
                    best = max(best, px * 1.6 + 4)
            if best == DEFAULT_H:
                new.pop(r, None)
            else:
                new[r] = best
        self.wb.set_prop(sh.name, "rowh", new, "Auto-fit row")
        self.reload_axes()
        self._place_editor()
        self.viewport().update()
        self.axes_changed.emit()

    # floating objects (charts) ------------------------------------------
    def add_floating(self, w):
        self._floating.append(w)
        w.setParent(self.viewport())
        self._place_floating()
        w.show()

    def remove_floating(self, w):
        if w in self._floating:
            self._floating.remove(w)
        w.hide()
        w.deleteLater()

    def clear_floating(self):
        for w in list(self._floating):
            self.remove_floating(w)

    def _place_floating(self):
        for w in self._floating:
            w.place(self)


class _FakeMove:
    def __init__(self, grid, pos):
        self._p = pos

    def position(self):
        return QPointF(self._p)

    def buttons(self):
        return Qt.LeftButton


def _esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


MONTHS = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"]
DAYS = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]


def _list_step(text):
    """names of months / days -> (list, index) else None"""
    t = text.strip().lower()
    for names in (MONTHS, DAYS, [m[:3] for m in MONTHS], [d[:3] for d in DAYS]):
        if t in names:
            return names, names.index(t)
    return None


def _case_like(src, s):
    if src.isupper():
        return s.upper()
    if src[:1].isupper():
        return s.capitalize()
    return s


def _fill_series(col, n, fwd, single_copy, src_idx, targets, vertical, wb):
    """new cells for n targets from the source cells"""
    k = len(col)
    out = []
    vals = [c[0] if c else None for c in col]
    numeric = all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in vals) and k >= 1
    formula = any(isinstance(v, str) and v.startswith("=") for v in vals)

    def style_of(i):
        return col[i % k][1] if col[i % k] else 0
    step = None
    if numeric and not single_copy and not formula:
        if k >= 2:
            step = (vals[-1] - vals[0]) / (k - 1)
        else:
            step = 0 if True else 1
    text_pat = None
    if k == 1 and not single_copy and isinstance(vals[0], str) and not vals[0].startswith("="):
        ls = _list_step(vals[0])
        if ls:
            text_pat = ("list", ls)
        else:
            m = re.match(r"^(.*?)(\d+)(\D*)$", vals[0])
            if m:
                text_pat = ("num", m)
    elif k >= 2 and not single_copy and all(isinstance(v, str) for v in vals):
        ms = [re.match(r"^(.*?)(\d+)(\D*)$", v) for v in vals]
        if all(ms) and len({(m.group(1), m.group(3)) for m in ms}) == 1:
            text_pat = ("num2", ms)
    for i, t in enumerate(targets):
        j = i + 1 if fwd else -(i + 1)
        si = (j - 1) % k if fwd else (k + (j % k) - 0) % k
        if not fwd:
            si = (-(i + 1)) % k
        else:
            si = i % k
        if numeric and step is not None and not formula:
            base = vals[0] if not fwd else vals[-1]
            v = vals[-1] + step * (i + 1) if fwd else vals[0] - step * (i + 1)
            if k == 1:
                v = vals[0]
            out.append((v if not float(v).is_integer() else (int(v) if isinstance(vals[0], int) else v), style_of(si)))
            continue
        src = vals[si]
        if text_pat and text_pat[0] == "list":
            names, idx = text_pat[1]
            nm = names[(idx + (i + 1 if fwd else -(i + 1))) % len(names)]
            out.append((_case_like(src, nm), style_of(si)))
        elif text_pat and text_pat[0] == "num":
            m = text_pat[1]
            n_ = int(m.group(2)) + (i + 1 if fwd else -(i + 1))
            out.append((m.group(1) + str(max(0, n_)).zfill(len(m.group(2))) + m.group(3), style_of(si)))
        elif text_pat and text_pat[0] == "num2":
            ms = text_pat[1]
            d = int(ms[-1].group(2)) - int(ms[0].group(2))
            d = d // (len(ms) - 1) if len(ms) > 1 else 1
            base = int(ms[-1].group(2)) if fwd else int(ms[0].group(2))
            n_ = base + d * (i + 1) * (1 if fwd else -1)
            out.append((ms[0].group(1) + str(max(0, n_)).zfill(len(ms[0].group(2))) + ms[0].group(3), style_of(si)))
        elif isinstance(src, str) and src.startswith("="):
            off = (targets[i] - src_idx[si])
            raw = F.shift_formula(src, off if vertical else 0, 0 if vertical else off)
            out.append((raw, style_of(si)))
        else:
            out.append((src, style_of(si)) if (src is not None or style_of(si)) else None)
    return out
