#!/usr/bin/env python3
"""Digonto Sheets - spreadsheets with formulas, charts, XLSX and CSV. PySide6 only."""
import copy
import os
import re
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.setrecursionlimit(6000)

from PySide6.QtCore import QEvent, QPoint, QRectF, QSize, Qt, QTimer
from PySide6.QtGui import QAction, QActionGroup, QColor, QFont, QFontDatabase, QIcon, QKeySequence, QPageLayout, QPageSize, QPdfWriter, QTextDocument
from PySide6.QtWidgets import (QDialog, QApplication, QCheckBox, QComboBox, QFileDialog, QFrame, QHBoxLayout, QInputDialog, QLabel, QLineEdit, QListWidget,
                               QListWidgetItem, QMainWindow, QMenu, QMessageBox, QPushButton, QSlider, QTabBar, QToolBar, QToolButton, QVBoxLayout,
                               QWidget, QWidgetAction, QStackedWidget, QDialog)

import digonto_ui as ui
import digonto_office as office
from sheetkit import __version__ as VERSION
from sheetkit import formula as F
from sheetkit import dsht, xlsxio
from sheetkit.cfmt import parse_range, range_text
from sheetkit.charts import ChartWidget
from sheetkit.dialogs import (ValidationDialog, CFDialog, ChartDialog, ColorButton, FindReplace, FormatCells, InsertFunction, NamesDialog, SortDialog, TextToColumns)
from sheetkit.grid import COLS, ROWS, Grid
from sheetkit.icons import icon
from sheetkit.model import DEFAULT, DEFAULT_H, DEFAULT_W, Workbook, parse_input
from sheetkit.numfmt import FORMATS, format_value

APP_NAME = "Digonto Sheets"
SLUG = "digonto-sheets"
FILTERS = "Spreadsheets (*.xlsx *.dsht *.csv *.tsv *.txt);;Excel workbook (*.xlsx);;Digonto spreadsheet (*.dsht);;CSV text (*.csv *.tsv *.txt);;All files (*)"


def act(window, text, fn, shortcut=None, tip=None, icon_name=None, checkable=False):
    a = QAction(text, window)
    if shortcut:
        a.setShortcut(shortcut)
    if tip:
        a.setToolTip(tip)
        a.setStatusTip(tip)
    if icon_name:
        a.setIcon(icon(icon_name))
    a.setCheckable(checkable)
    a.triggered.connect(lambda checked=False: fn())
    window.addAction(a)
    return a


class Suggest(QListWidget):
    """function name suggestions under the cell editor"""

    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowFlags(Qt.ToolTip)
        self.setFocusPolicy(Qt.NoFocus)
        self.setStyleSheet("QListWidget{background:white; border:1px solid #AC9CB8; border-radius:0;} QListWidget::item{padding:3px 8px;} QListWidget::item:selected{background:#FBDCE5; color:#1B1040;}")
        self.itemClicked.connect(lambda it: self.chosen(it.text()))
        self.chosen = lambda name: None


class SheetsWindow(QMainWindow):
    def __init__(self, path=None):
        super().__init__()
        self.wb = Workbook()
        self.setWindowTitle(APP_NAME)
        self.setWindowIcon(ui.app_icon(HERE, SLUG))
        self.find_dlg = None
        self.format_painter = None
        self._filling = False
        self._charts = []
        self._last_search = None
        self.grid = Grid(self.wb)
        self.build_actions()
        self.build_menu()
        header = ui.make_header(APP_NAME, "Spreadsheets with formulas, charts, XLSX and CSV", ui.app_icon(HERE, SLUG), self.about)
        central = QWidget()
        cl = QVBoxLayout(central)
        cl.setContentsMargins(0, 0, 0, 0)
        cl.setSpacing(0)
        cl.addWidget(header)
        self.tb1 = self.make_toolbar()
        self.tb2 = self.make_toolbar()
        self.work = QWidget()
        wl = QVBoxLayout(self.work)
        wl.setContentsMargins(0, 0, 0, 0)
        wl.setSpacing(0)
        wl.addWidget(self.tb1)
        wl.addWidget(self.tb2)
        self.build_toolbars()
        wl.addWidget(self.build_formula_bar())
        wl.addWidget(self.grid, 1)
        wl.addWidget(self.build_tabs())
        self.start = office.StartScreen(SLUG, "Digonto Sheets", "Spreadsheets with formulas, charts, XLSX and CSV.", [
            ("+", "Blank spreadsheet", "Start with an empty sheet.", self.start_new),
            ("\u2197", "Open ...", "Open a file from this computer, a recent file or a folder shortcut.", self.open_dialog)])
        self.start.open_recent.connect(self.open_path)
        self.stack = QStackedWidget()
        self.stack.addWidget(self.start)
        self.stack.addWidget(self.work)
        cl.addWidget(self.stack, 1)
        self.setCentralWidget(central)
        self.build_status()
        self.suggest = Suggest(self)
        self.suggest.chosen = self.accept_suggestion
        self.suggest.hide()
        self.connect_signals()
        self.attach(self.wb)
        ui.fit_window(self, 1280, 860)
        self.show_start()
        if path and os.path.isfile(path):
            QTimer.singleShot(0, lambda: self.open_path(path))

    # ------------------------------------------------------------------ build
    def make_toolbar(self):
        tb = QToolBar()
        tb.setMovable(False)
        tb.setIconSize(QSize(20, 20))
        tb.setStyleSheet("QToolBar{background:#F7FAF8; border:0; border-bottom:1px solid #E1DBE6; spacing:1px; padding:2px 6px;}"
                         "QToolButton{border:1px solid transparent; border-radius:5px; padding:2px; min-width:22px; min-height:22px;}"
                         "QToolButton:hover{background:#E8E0EF; border-color:#D0C4E1;} QToolButton:checked{background:#FBDCE5; border-color:#A98AC1;}"
                         "QToolButton::menu-button{width:12px; border-left:0;} QToolButton::menu-indicator{image:none;}")
        return tb

    def build_actions(self):
        A = self
        A.a_new = act(A, "New", self.new_file, "Ctrl+N", "New workbook", "new")
        A.a_open = act(A, "Open ...", self.open_dialog, "Ctrl+O", "Open a spreadsheet", "open")
        A.a_save = act(A, "Save", self.save, "Ctrl+S", "Save (XLSX)", "save")
        A.a_saveas = act(A, "Save as ...", lambda: self.save(True), "Ctrl+Shift+S")
        A.a_pdf = act(A, "Export PDF ...", self.export_pdf, "Ctrl+Shift+E", "Save as PDF", "save")
        A.a_print = act(A, "Print ...", self.print_sheet, "Ctrl+P", "Print the sheet", "print")
        A.a_csv = act(A, "Export this sheet as CSV ...", self.export_csv)
        A.a_undo = act(A, "Undo", self.undo, None, "Undo (Ctrl+Z)", "undo")
        A.a_redo = act(A, "Redo", self.redo, None, "Redo (Ctrl+Y)", "redo")
        A.a_bold = act(A, "Bold", lambda: self.toggle("bold"), "Ctrl+B", "Bold", "bold", True)
        A.a_italic = act(A, "Italic", lambda: self.toggle("italic"), "Ctrl+I", "Italic", "italic", True)
        A.a_under = act(A, "Underline", lambda: self.toggle("underline"), "Ctrl+U", "Underline", "underline", True)
        A.a_strike = act(A, "Strikethrough", lambda: self.toggle("strike"), "Ctrl+Shift+X", "Strikethrough", "strike", True)
        A.a_wrap = act(A, "Wrap text", lambda: self.toggle("wrap"), None, "Wrap text", "wrap", True)
        A.a_merge = act(A, "Merge cells", self.merge_toggle, None, "Merge / unmerge cells", "merge")
        A.a_find = act(A, "Find ...", lambda: self.show_find(False), "Ctrl+F", "Find", "find")
        A.a_replace = act(A, "Find and replace ...", lambda: self.show_find(True), "Ctrl+H")
        A.a_goto = act(A, "Go to ...", self.goto_dialog, "Ctrl+G")
        A.a_fx = act(A, "Insert function ...", self.insert_function, "Shift+F3", "Insert function", "fx")
        A.a_autosum = act(A, "Sum", lambda: self.autosum("SUM"), "Alt+=", "AutoSum", "sum")
        A.a_filter = act(A, "Filter", self.toggle_filter, "Ctrl+Shift+L", "Turn the filter on / off", "filter")
        A.a_chart = act(A, "Chart ...", self.insert_chart, None, "Insert chart from selected cells", "chart")
        A.a_fmt = act(A, "Format cells ...", self.format_dialog, "Ctrl+1")
        A.a_paintfmt = act(A, "Paint format", self.paint_format, None, "Copy the look of this cell to others", "paintformat", True)

    def build_menu(self):
        mb = self.menuBar()
        fm = mb.addMenu("File")
        for a in (self.a_new, self.a_open, self.a_save, self.a_saveas):
            fm.addAction(a)
        self.a_close = act(self, "Close project", self.close_document, "Ctrl+W")
        fm.addSeparator()
        fm.addAction(self.a_print)
        fm.addAction(self.a_pdf)
        fm.addAction(self.a_csv)
        fm.addAction("Export as HTML table ...", self.export_html)
        fm.addSeparator()
        fm.addAction(self.a_close)
        fm.addSeparator()
        fm.addAction(QAction("Exit\tCtrl+Q", self, triggered=self.close))
        q = QAction("Quit (force)", self, triggered=lambda: os._exit(0))
        fm.addAction(q)
        QAction("", self)
        em = mb.addMenu("Edit")
        em.addAction("Undo\tCtrl+Z", self.undo)
        em.addAction("Redo\tCtrl+Y", self.redo)
        em.addSeparator()
        em.addAction("Cut\tCtrl+X", lambda: self.grid.copy(cut=True))
        em.addAction("Copy\tCtrl+C", self.grid.copy)
        em.addAction("Paste\tCtrl+V", lambda: self.grid.paste("all"))
        ps = em.addMenu("Paste special")
        for t, m in (("Values only", "values"), ("Formulas only", "formulas"), ("Formats only", "formats"), ("Transposed", "transpose")):
            ps.addAction(t, lambda m=m: self.grid.paste(m))
        em.addSeparator()
        em.addAction("Fill down\tCtrl+D", lambda: self.grid.fill_edge("down"))
        em.addAction("Fill right\tCtrl+R", lambda: self.grid.fill_edge("right"))
        cm = em.addMenu("Clear")
        cm.addAction("Clear contents\tDel", self.grid.clear_contents)
        cm.addAction("Clear formats", self.grid.clear_formats)
        cm.addAction("Clear all", self.grid.clear_all)
        em.addSeparator()
        em.addAction(self.a_find)
        em.addAction(self.a_replace)
        em.addAction(self.a_goto)
        em.addAction("Select all\tCtrl+A", self.grid.select_all)
        vm = mb.addMenu("View")
        self.a_grid = vm.addAction("Show gridlines", self.toggle_gridlines)
        self.a_grid.setCheckable(True)
        self.a_grid.setChecked(True)
        self.a_formulas = vm.addAction("Show formulas", self.toggle_show_formulas)
        self.a_formulas.setCheckable(True)
        fz = vm.addMenu("Freeze")
        fz.addAction("Top row", lambda: self.freeze(1, 0))
        fz.addAction("First column", lambda: self.freeze(0, 1))
        fz.addAction("Top row and first column", lambda: self.freeze(1, 1))
        fz.addAction("Up to the selected cell", lambda: self.freeze(self.grid.cur[0], self.grid.cur[1]))
        fz.addAction("Unfreeze", lambda: self.freeze(0, 0))
        zm = vm.addMenu("Zoom")
        for z in (50, 75, 90, 100, 125, 150, 200):
            zm.addAction(f"{z} %", lambda z=z: self.grid.set_zoom(z / 100))
        im = mb.addMenu("Insert")
        im.addAction("Row above", lambda: self.insert_rows(False))
        im.addAction("Row below", lambda: self.insert_rows(True))
        im.addAction("Column left", lambda: self.insert_cols(False))
        im.addAction("Column right", lambda: self.insert_cols(True))
        im.addSeparator()
        im.addAction(self.a_fx)
        im.addAction(self.a_chart)
        im.addAction("Note on this cell ...", self.edit_note)
        im.addAction("New sheet", self.add_sheet)
        fo = mb.addMenu("Format")
        fo.addAction(self.a_fmt)
        for a in (self.a_bold, self.a_italic, self.a_under, self.a_strike):
            fo.addAction(a)
        fo.addSeparator()
        nf = fo.addMenu("Number")
        for label, code in FORMATS:
            nf.addAction(label, lambda code=code: self.apply(lambda s: s._replace(fmt=code), "Number format"))
        fo.addAction(self.a_wrap)
        fo.addAction(self.a_merge)
        fo.addSeparator()
        fo.addAction("Column width ...", self.col_width_dialog)
        fo.addAction("Row height ...", self.row_height_dialog)
        fo.addAction("Auto-fit column width", lambda: self.grid.autofit_columns(self.sel_columns()))
        fo.addAction("Auto-fit row height", lambda: self.grid.autofit_rows(self.sel_rows()))
        fo.addAction("Hide rows", lambda: self.hide("row", True))
        fo.addAction("Hide columns", lambda: self.hide("col", True))
        fo.addAction("Unhide rows and columns", self.unhide_all)
        fo.addSeparator()
        fo.addAction("Conditional formatting ...", self.cf_dialog)
        dm = mb.addMenu("Data")
        dm.addAction("Sort A to Z", lambda: self.quick_sort(False))
        dm.addAction("Sort Z to A", lambda: self.quick_sort(True))
        dm.addAction("Sort range ...", self.sort_dialog)
        dm.addAction(self.a_filter)
        dm.addSeparator()
        dm.addAction("Data validation ...", self.validation_dialog)
        dm.addAction("Remove duplicate rows", self.remove_duplicates)
        dm.addAction("Split text into columns ...", self.text_to_columns)
        dm.addAction("Named ranges ...", self.names_dialog)
        hm = mb.addMenu("Help")
        hm.addAction("Function list ...", self.insert_function)
        hm.addAction("Keyboard shortcuts", self.shortcuts_help)
        hm.addAction(f"About {APP_NAME}", self.about)

    def build_toolbars(self):
        t1, t2 = self.tb1, self.tb2
        for a in (self.a_new, self.a_open, self.a_save, self.a_pdf):
            t1.addAction(a)
        t1.addSeparator()
        t1.addAction(self.a_undo)
        t1.addAction(self.a_redo)
        t1.addAction(self.a_paintfmt)
        t1.addSeparator()
        self.font_box = QComboBox()
        self.font_box.setMinimumWidth(130)
        self.font_box.addItem("Default")
        self.font_box.addItems(sorted(set(QFontDatabase.families())))
        self.font_box.setToolTip("Font")
        self.font_box.activated.connect(lambda i: self.apply(lambda s: s._replace(font=None if i == 0 else self.font_box.currentText()), "Font"))
        t1.addWidget(self.font_box)
        self.size_box = QComboBox()
        self.size_box.setEditable(True)
        self.size_box.addItems(["8", "9", "10", "11", "12", "14", "16", "18", "20", "24", "28", "36", "48", "72"])
        self.size_box.setCurrentText("11")
        self.size_box.setFixedWidth(64)
        self.size_box.setToolTip("Font size")
        self.size_box.activated.connect(lambda i: self.set_size(self.size_box.currentText()))
        self.size_box.lineEdit().editingFinished.connect(lambda: self.set_size(self.size_box.currentText()))
        t1.addWidget(self.size_box)
        t1.addSeparator()
        for a in (self.a_bold, self.a_italic, self.a_under, self.a_strike):
            t1.addAction(a)
        self.cb_text = ColorButton("textcolor", "Text colour", None, "Automatic")
        self.cb_text.picked.connect(lambda c: self.apply(lambda s: s._replace(color=c), "Text colour"))
        self.cb_fill = ColorButton("fill", "Fill colour", None, "No fill")
        self.cb_fill.picked.connect(lambda c: self.apply(lambda s: s._replace(bg=c), "Fill colour"))
        t1.addWidget(self.cb_text)
        t1.addWidget(self.cb_fill)
        bm = QToolButton()
        bm.setIcon(icon("borders"))
        bm.setToolTip("Borders")
        bm.setPopupMode(QToolButton.InstantPopup)
        m = QMenu(bm)
        for label, ic, mode in (("All borders", "b_all", "all"), ("Outside borders", "borders", "outer"), ("Bottom border", "b_bottom", "bottom"), ("Top border", "b_top", "top"),
                                ("Left border", "b_left", "left"), ("Right border", "b_right", "right"), ("No border", "b_none", "none")):
            m.addAction(icon(ic), label, lambda mode=mode: self.borders(mode))
        bm.setMenu(m)
        t1.addWidget(bm)
        t1.addAction(self.a_merge)
        t1.addSeparator()
        grp = QActionGroup(self)
        for name, ic, key in (("Align left", "alignl", "left"), ("Align centre", "alignc", "center"), ("Align right", "alignr", "right")):
            a = act(self, name, lambda key=key: self.set_h(key), None, name, ic, True)
            grp.addAction(a)
            t1.addAction(a)
            setattr(self, "a_h_" + key, a)
        grp.setExclusive(False)
        vm = QToolButton()
        vm.setIcon(icon("valignmid"))
        vm.setToolTip("Vertical alignment")
        vm.setPopupMode(QToolButton.InstantPopup)
        m = QMenu(vm)
        for label, ic, key in (("Top", "valigntop", "top"), ("Middle", "valignmid", "middle"), ("Bottom", "valignbot", "bottom")):
            m.addAction(icon(ic), label, lambda key=key: self.apply(lambda s: s._replace(v=key), "Vertical alignment"))
        vm.setMenu(m)
        t1.addWidget(vm)
        t1.addAction(self.a_wrap)
        t1.addAction(icon("outdent"), "Decrease indent", lambda: self.apply(lambda s: s._replace(indent=max(0, s.indent - 1)), "Indent"))
        t1.addAction(icon("indent"), "Increase indent", lambda: self.apply(lambda s: s._replace(indent=min(15, s.indent + 1)), "Indent"))
        # row 2
        self.fmt_box = QComboBox()
        self.fmt_box.setMinimumWidth(170)
        for label, code in FORMATS:
            self.fmt_box.addItem(label.split("  ")[0], code)
        self.fmt_box.addItem("More formats ...", "__more__")
        self.fmt_box.setToolTip("Number format")
        self.fmt_box.activated.connect(self.fmt_chosen)
        t2.addWidget(self.fmt_box)
        for ic, tip, code in (("currency", "Currency", "$#,##0.00"), ("percent", "Percent", "0%"), ("comma", "Thousands separator", "#,##0.00")):
            t2.addAction(icon(ic), tip, lambda code=code: self.apply(lambda s: s._replace(fmt=code), "Number format"))
        t2.addAction(icon("decless"), "Fewer decimals", lambda: self.decimals(-1))
        t2.addAction(icon("decmore"), "More decimals", lambda: self.decimals(1))
        t2.addSeparator()
        sm = QToolButton()
        sm.setIcon(icon("sum"))
        sm.setToolTip("AutoSum")
        sm.setPopupMode(QToolButton.MenuButtonPopup)
        sm.clicked.connect(lambda: self.autosum("SUM"))
        m = QMenu(sm)
        for n in ("SUM", "AVERAGE", "COUNT", "MAX", "MIN"):
            m.addAction({"SUM": "Sum", "AVERAGE": "Average", "COUNT": "Count numbers", "MAX": "Max", "MIN": "Min"}[n], lambda n=n: self.autosum(n))
        sm.setMenu(m)
        t2.addWidget(sm)
        t2.addAction(self.a_fx)
        t2.addSeparator()
        t2.addAction(icon("sortaz"), "Sort A to Z", lambda: self.quick_sort(False))
        t2.addAction(icon("sortza"), "Sort Z to A", lambda: self.quick_sort(True))
        t2.addAction(self.a_filter)
        fz = QToolButton()
        fz.setIcon(icon("freeze"))
        fz.setToolTip("Freeze rows / columns")
        fz.setPopupMode(QToolButton.InstantPopup)
        m = QMenu(fz)
        m.addAction("Top row", lambda: self.freeze(1, 0))
        m.addAction("First column", lambda: self.freeze(0, 1))
        m.addAction("Up to the selected cell", lambda: self.freeze(self.grid.cur[0], self.grid.cur[1]))
        m.addAction("Unfreeze", lambda: self.freeze(0, 0))
        fz.setMenu(m)
        t2.addWidget(fz)
        t2.addSeparator()
        cf = QToolButton()
        cf.setIcon(icon("cf"))
        cf.setToolTip("Conditional formatting")
        cf.setPopupMode(QToolButton.InstantPopup)
        m = QMenu(cf)
        for label, kind in (("Greater than ...", "greater than"), ("Less than ...", "less than"), ("Equal to ...", "equal to"), ("Text contains ...", "text contains")):
            m.addAction(label, lambda kind=kind: self.cf_quick(kind))
        m.addSeparator()
        for label, kind in (("Duplicate values", "duplicate values"), ("Top 10", "top 10"), ("Bottom 10", "bottom 10"), ("Above average", "above average"), ("Colour scale", "color scale"), ("Data bars", "data bars")):
            m.addAction(label, lambda kind=kind: self.cf_quick(kind, ask=False))
        m.addSeparator()
        m.addAction("Manage rules ...", self.cf_dialog)
        cf.setMenu(m)
        t2.addWidget(cf)
        t2.addAction(self.a_chart)
        t2.addAction(self.a_find)
        cl = QToolButton()
        cl.setIcon(icon("clear"))
        cl.setToolTip("Clear")
        cl.setPopupMode(QToolButton.InstantPopup)
        m = QMenu(cl)
        m.addAction("Clear contents", self.grid.clear_contents)
        m.addAction("Clear formats", self.grid.clear_formats)
        m.addAction("Clear all", self.grid.clear_all)
        cl.setMenu(m)
        t2.addWidget(cl)

    def build_formula_bar(self):
        w = QWidget()
        w.setStyleSheet("background:#FFFFFF;")
        lay = QHBoxLayout(w)
        lay.setContentsMargins(6, 4, 6, 4)
        lay.setSpacing(6)
        self.name_box = QLineEdit()
        self.name_box.setFixedWidth(120)
        self.name_box.setToolTip("Name box: type a cell (B7), a range (A1:C9) or a name and press Enter")
        self.name_box.setStyleSheet("QLineEdit{padding:3px 6px; font-weight:600;}")
        self.fx_btn = QToolButton()
        self.fx_btn.setDefaultAction(self.a_fx)
        self.fx_btn.setAutoRaise(True)
        self.formula = QLineEdit()
        self.formula.setPlaceholderText("Type a value or a formula, like =SUM(A1:A9)")
        self.formula.setStyleSheet("QLineEdit{padding:3px 6px;}")
        self.hint = QLabel("")
        self.hint.setStyleSheet("color:#8A2150; background:#EFE9F4; border-radius:6px; padding:2px 8px;")
        self.hint.hide()
        lay.addWidget(self.name_box)
        lay.addWidget(self.fx_btn)
        lay.addWidget(self.formula, 1)
        lay.addWidget(self.hint)
        return w

    def build_tabs(self):
        w = QWidget()
        w.setStyleSheet("background:#F2EDF7;")
        lay = QHBoxLayout(w)
        lay.setContentsMargins(6, 2, 6, 0)
        lay.setSpacing(2)
        add = QToolButton()
        add.setIcon(icon("plus"))
        add.setToolTip("Add a sheet")
        add.setAutoRaise(True)
        add.clicked.connect(self.add_sheet)
        lay.addWidget(add)
        self.tabs = QTabBar()
        self.tabs.setMovable(True)
        self.tabs.setExpanding(False)
        self.tabs.setDrawBase(False)
        self.tabs.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tabs.setStyleSheet("QTabBar::tab{background:#EBE3F3; color:#3B2A63; padding:5px 16px; margin-right:1px; border:1px solid #D7CFDD; border-top:0; border-bottom-left-radius:6px; border-bottom-right-radius:6px;}"
                                "QTabBar::tab:selected{background:white; color:#8A2150; font-weight:600; border-bottom:3px solid #C2395C;}"
                                "QTabBar::tab:hover:!selected{background:#E3D7EF;}")
        lay.addWidget(self.tabs, 1)
        return w

    def build_status(self):
        sb = self.statusBar()
        self.stat_msg = QLabel("Ready")
        self.stat_sel = QLabel("")
        self.stat_sel.setStyleSheet("padding:0 12px;")
        self.zoom_lbl = QLabel("100%")
        self.zoom_sl = QSlider(Qt.Horizontal)
        self.zoom_sl.setRange(40, 300)
        self.zoom_sl.setValue(100)
        self.zoom_sl.setFixedWidth(130)
        self.zoom_sl.valueChanged.connect(lambda v: self.grid.set_zoom(v / 100))
        sb.addWidget(self.stat_msg, 1)
        sb.addPermanentWidget(self.stat_sel)
        sb.addPermanentWidget(self.zoom_sl)
        sb.addPermanentWidget(self.zoom_lbl)

    def connect_signals(self):
        g = self.grid
        g.selection_changed.connect(self.sel_changed)
        g.edit_text_changed.connect(self.editor_text_changed)
        g.edit_state.connect(self.edit_state_changed)
        g.zoom_changed.connect(self.zoom_changed)
        g.filter_clicked.connect(self.filter_menu)
        g.context_menu.connect(self.show_context)
        g.message.connect(self.stat_msg.setText)
        g.undo_requested.connect(self.undo)
        g.redo_requested.connect(self.redo)
        g.key_filter = self.editor_key_filter
        self.name_box.returnPressed.connect(self.name_entered)
        self.formula.returnPressed.connect(self.formula_entered)
        self.formula.textEdited.connect(self.formula_edited)
        self.tabs.currentChanged.connect(self.tab_changed)
        self.tabs.tabMoved.connect(self.tab_moved)
        self.tabs.customContextMenuRequested.connect(self.tab_menu)
        self.tabs.tabBarDoubleClicked.connect(lambda i: self.rename_sheet(i))
        self.formula.installEventFilter(self)

    def eventFilter(self, obj, ev):
        if obj is self.formula and ev.type() == QEvent.KeyPress and ev.key() == Qt.Key_Escape:
            self.grid.cancel_edit()
            self.sel_changed()
            self.grid.setFocus()
            return True
        return super().eventFilter(obj, ev)

    # ------------------------------------------------------------------ workbook plumbing
    def attach(self, wb):
        if self.wb.listeners and self._on_change in self.wb.listeners:
            self.wb.listeners.remove(self._on_change)
        self.wb = wb
        wb.listeners.append(self._on_change)
        self.grid.wb = wb
        self.grid.cf.ver = -1
        self.rebuild_tabs()
        self.tabs.setCurrentIndex(0)
        self.grid.set_sheet(wb.sheets[0].name)
        self.load_charts()
        self.refresh_all()
        self.update_title()

    def _on_change(self, info):
        kind = info[0] if info else "all"
        if kind in ("all", "props"):
            names = [s.name for s in self.wb.sheets]
            if names != [self.tabs.tabText(i) for i in range(self.tabs.count())]:
                cur = self.grid.sheet.name
                self.rebuild_tabs()
                self.tabs.blockSignals(True)
                idx = names.index(cur) if cur in names else 0
                self.tabs.setCurrentIndex(idx)
                self.tabs.blockSignals(False)
                self.grid.sheet = self.wb.sheet(names[idx])
            else:
                self.grid.sheet = self.wb.sheet(self.grid.sheet.name) or self.wb.sheets[0]
            self.grid.reload_axes()
        elif kind == "cells":
            self.grid.sheet = self.wb.sheet(self.grid.sheet.name) or self.grid.sheet
        self.grid.viewport().update()
        for c in self._charts:
            c.update()
        self.update_title()
        self.update_state()
        if kind == "all":
            self.load_charts()

    def update_title(self):
        if self.on_start():
            self.setWindowTitle(APP_NAME)
            return
        name = os.path.basename(self.wb.path) if self.wb.path else "Untitled"
        self.setWindowTitle(f"{'* ' if self.wb.dirty else ''}{name} - {APP_NAME}")

    def rebuild_tabs(self):
        self.tabs.blockSignals(True)
        while self.tabs.count():
            self.tabs.removeTab(0)
        for s in self.wb.sheets:
            i = self.tabs.addTab(s.name)
            if s.tab_color:
                self.tabs.setTabTextColor(i, QColor(s.tab_color))
        self.tabs.blockSignals(False)

    def refresh_all(self):
        self.sel_changed()
        self.update_state()
        self.grid.viewport().update()

    def load_charts(self):
        for c in self._charts:
            self.grid.remove_floating(c)
        self._charts = []
        for ch in self.grid.sheet.charts:
            self.add_chart_widget(ch)

    def add_chart_widget(self, ch):
        w = ChartWidget(self.wb, ch)
        w.changed.connect(lambda: self._chart_changed())
        w.edit_requested.connect(self.edit_chart)
        w.delete_requested.connect(self.delete_chart)
        self.grid.add_floating(w)
        self._charts.append(w)
        return w

    def _chart_changed(self):
        self.wb.dirty = True
        self.update_title()

    # ------------------------------------------------------------------ selection / formula bar
    def sel_changed(self):
        g = self.grid
        r1, c1, r2, c2 = g.sel[-1]
        if (r1, c1) == (r2, c2) or g.merge_at(r1, c1) == (r1, c1, r2, c2):
            nm = F.cell_name(r1, c1)
        elif r1 == 0 and r2 >= ROWS - 1:
            nm = f"{F.col_name(c1)}:{F.col_name(c2)}"
        elif c1 == 0 and c2 >= COLS - 1:
            nm = f"{r1 + 1}:{r2 + 1}"
        else:
            nm = f"{F.cell_name(r1, c1)}:{F.cell_name(r2, c2)}"
        if not self.name_box.hasFocus():
            self.name_box.setText(nm)
        if g.editor is None and not self.formula.hasFocus():
            r, c = g.cur
            self.formula.setText(self.wb.edit_text(g.sheet.name, r, c))
        self.update_state()
        self.update_stats()

    def update_state(self):
        g = self.grid
        r, c = g.cur
        st = self.wb.style_of(g.sheet, r, c)
        for a, v in ((self.a_bold, st.bold), (self.a_italic, st.italic), (self.a_under, st.underline), (self.a_strike, st.strike), (self.a_wrap, st.wrap)):
            a.blockSignals(True)
            a.setChecked(v)
            a.blockSignals(False)
        for key in ("left", "center", "right"):
            a = getattr(self, "a_h_" + key)
            a.blockSignals(True)
            a.setChecked(st.h == key)
            a.blockSignals(False)
        self.size_box.blockSignals(True)
        self.size_box.setCurrentText(f"{st.size:g}" if st.size else "11")
        self.size_box.blockSignals(False)
        self.font_box.blockSignals(True)
        self.font_box.setCurrentText(st.font or "Default")
        self.font_box.blockSignals(False)
        self.fmt_box.blockSignals(True)
        idx = next((i for i in range(self.fmt_box.count()) if self.fmt_box.itemData(i) == st.fmt), -1)
        if idx < 0:
            self.fmt_box.setCurrentIndex(self.fmt_box.count() - 1)
            self.fmt_box.setEditable(False)
        else:
            self.fmt_box.setCurrentIndex(idx)
        self.fmt_box.blockSignals(False)
        self.cb_text.color, self.cb_fill.color = st.color, st.bg
        self.cb_text.refresh()
        self.cb_fill.refresh()
        self.a_merge.setText("Unmerge cells" if g.merge_at(r, c) else "Merge cells")
        self.a_filter.setCheckable(True)
        self.a_filter.blockSignals(True)
        self.a_filter.setChecked(bool(g.sheet.filter))
        self.a_filter.blockSignals(False)
        self.a_undo.setEnabled(bool(self.wb.undo_stack))
        self.a_redo.setEnabled(bool(self.wb.redo_stack))
        self.a_grid.setChecked(g.sheet.grid)

    def update_stats(self):
        g = self.grid
        r1, c1, r2, c2 = g.sel[-1]
        if (r1, c1) == (r2, c2):
            self.stat_sel.setText("")
            return
        vals = []
        cnt = 0
        sheet = g.sheet
        area = (r2 - r1 + 1) * (c2 - c1 + 1)
        if area > len(sheet.cells):
            it = ((k) for k in sheet.cells if r1 <= k[0] <= r2 and c1 <= k[1] <= c2)
        else:
            it = ((r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1) if (r, c) in sheet.cells)
        for (r, c) in it:
            v = self.wb.cell_value(sheet.name, r, c)
            if v is None or v == "":
                continue
            cnt += 1
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                vals.append(v)
        if vals:
            s = sum(vals)
            self.stat_sel.setText(f"Sum {format_value(s, 'General')}    Average {format_value(s / len(vals), 'General')}    Count {cnt}    Min {format_value(min(vals), 'General')}    Max {format_value(max(vals), 'General')}")
        else:
            self.stat_sel.setText(f"Count {cnt}")

    def name_entered(self):
        t = self.name_box.text().strip()
        g = self.grid
        up = t.upper().replace("$", "")
        sheet = None
        if "!" in t:
            sheet, up = t.rsplit("!", 1)
            sheet = sheet.strip("'")
            up = up.upper().replace("$", "")
            if self.wb.has_sheet(sheet):
                self.tabs.setCurrentIndex(self.wb.sheet_names().index(self.wb.sheet(sheet).name))
        target = None
        if up in self.wb.names:
            ref = self.wb.names[up]
            sh = ref.split("!")[0].strip("'") if "!" in ref else None
            if sh and self.wb.has_sheet(sh):
                self.tabs.setCurrentIndex(self.wb.sheet_names().index(self.wb.sheet(sh).name))
            target = parse_range(ref.split("!")[-1])
        elif re.match(r"^[A-Z]{1,3}:[A-Z]{1,3}$", up):
            a, b = up.split(":")
            target = (0, F.col_index(a), ROWS - 1, F.col_index(b))
        elif re.match(r"^\d+:\d+$", up):
            a, b = up.split(":")
            target = (int(a) - 1, 0, int(b) - 1, COLS - 1)
        else:
            target = parse_range(up)
        if target:
            r1, c1, r2, c2 = target
            g.anchor = (r1, c1)
            g.set_selection(r1, c1, r2, c2, cur=(r1, c1))
            g.ensure_visible(r1, c1)
        elif re.match(r"^[A-Za-z_][\w\.]*$", t) and not re.match(r"^[A-Za-z]{1,3}\d+$", t):
            r1, c1, r2, c2 = g.sel[-1]
            self.wb.names[t.upper()] = f"{F.quote_sheet(g.sheet.name)}!{range_text(r1, c1, r2, c2)}".replace("!", "!")
            self.wb.touch(("props",))
            self.stat_msg.setText(f"Named {t.upper()}")
        else:
            QMessageBox.information(self, APP_NAME, "That is not a cell, range or name.")
        g.setFocus()

    def formula_edited(self, text):
        g = self.grid
        if g.editor is None:
            g.begin_edit(initial=text)
            g.editor.mode = "edit"
            self.formula.setFocus()
        else:
            g.set_editor_text(text)
        self.update_hint(text, self.formula.cursorPosition())

    def formula_entered(self):
        g = self.grid
        if g.editor is None:
            g.begin_edit(initial=self.formula.text())
            g.editor.mode = "edit"
        else:
            g.set_editor_text(self.formula.text())
        g.commit_edit(move=(1, 0))
        g.setFocus()

    def editor_text_changed(self, t):
        if self.formula.text() != t:
            self.formula.blockSignals(True)
            self.formula.setText(t)
            self.formula.blockSignals(False)
        ed = self.grid.editor
        self.update_hint(t, ed.cursorPosition() if ed else len(t))
        self.update_suggest(t, ed)

    def edit_state_changed(self, on):
        if not on:
            self.suggest.hide()
            self.hint.hide()
            self.sel_changed()
        else:
            ed = self.grid.editor
            if ed is not None:
                ed.cursorPositionChanged.connect(lambda a, b: self.update_hint(ed.text(), b))

    def update_hint(self, text, pos):
        if not text.startswith("="):
            self.hint.hide()
            return
        depth, i, name = 0, min(pos, len(text)) - 1, None
        instr = False
        while i >= 0:
            ch = text[i]
            if ch == '"':
                instr = not instr
            elif not instr:
                if ch == ")":
                    depth += 1
                elif ch == "(":
                    if depth == 0:
                        m = re.search(r"([A-Za-z_][\w\.]*)$", text[:i])
                        name = m.group(1).upper() if m else None
                        break
                    depth -= 1
            i -= 1
        h = F.HELP.get(name or "")
        if h:
            self.hint.setText(h[0])
            self.hint.show()
        else:
            self.hint.hide()

    def update_suggest(self, t, ed):
        if ed is None or not t.startswith("="):
            self.suggest.hide()
            return
        pos = ed.cursorPosition()
        m = re.search(r"([A-Za-z_][\w\.]*)$", t[:pos])
        if not m or (pos < len(t) and re.match(r"[\w\(]", t[pos])):
            self.suggest.hide()
            return
        pre = m.group(1).upper()
        before = t[:m.start()].rstrip()
        if before and before[-1] not in "=+-*/^&,;(<>:%":
            self.suggest.hide()
            return
        names = [n for n in sorted(F.FUNCS) if n.startswith(pre)][:8]
        if not names or (len(names) == 1 and names[0] == pre):
            self.suggest.hide()
            return
        self.suggest.clear()
        self.suggest.addItems(names)
        self.suggest.setCurrentRow(0)
        self.suggest.setFixedSize(220, min(8, len(names)) * 24 + 4)
        self.suggest.move(ed.mapToGlobal(QPoint(0, ed.height() + 2)))
        self.suggest.show()
        ed.setFocus()

    def accept_suggestion(self, name):
        ed = self.grid.editor
        if ed is None:
            return
        t, pos = ed.text(), ed.cursorPosition()
        m = re.search(r"([A-Za-z_][\w\.]*)$", t[:pos])
        if m:
            ed.setText(t[:m.start()] + name + "(" + t[pos:])
            ed.setCursorPosition(m.start() + len(name) + 1)
        self.suggest.hide()
        ed.setFocus()

    def editor_key_filter(self, e):
        if not self.suggest.isVisible():
            return False
        k = e.key()
        if k in (Qt.Key_Down, Qt.Key_Up):
            r = self.suggest.currentRow() + (1 if k == Qt.Key_Down else -1)
            self.suggest.setCurrentRow(max(0, min(self.suggest.count() - 1, r)))
            return True
        if k == Qt.Key_Tab:
            it = self.suggest.currentItem()
            if it:
                self.accept_suggestion(it.text())
            return True
        if k == Qt.Key_Escape:
            self.suggest.hide()
            return True
        return False

    def zoom_changed(self, z):
        self.zoom_lbl.setText(f"{int(round(z * 100))}%")
        self.zoom_sl.blockSignals(True)
        self.zoom_sl.setValue(int(round(z * 100)))
        self.zoom_sl.blockSignals(False)

    # ------------------------------------------------------------------ editing helpers
    def undo(self):
        self.grid.cancel_edit()
        if self.wb.undo():
            self.after_history()

    def redo(self):
        self.grid.cancel_edit()
        if self.wb.redo():
            self.after_history()

    def after_history(self):
        self.grid.sheet = self.wb.sheet(self.grid.sheet.name) or self.wb.sheets[0]
        self.grid.reload_axes()
        self.sel_changed()
        self.grid.viewport().update()

    def sel_cells(self):
        return self.grid.cells_in_selection()

    def apply(self, fn, label="Format"):
        if self.grid.editor is not None:
            self.grid.commit_edit()
        self.wb.set_style(self.grid.sheet.name, self.sel_cells(), fn, label)
        self.update_state()

    def toggle(self, attr):
        r, c = self.grid.cur
        cur = getattr(self.wb.style_of(self.grid.sheet, r, c), attr)
        self.apply(lambda s: s._replace(**{attr: not cur}), attr.capitalize())
        if attr == "wrap":
            self.grid.autofit_rows(self.sel_rows()[:2000])

    def set_h(self, key):
        r, c = self.grid.cur
        cur = self.wb.style_of(self.grid.sheet, r, c).h
        self.apply(lambda s: s._replace(h=None if cur == key else key), "Align")

    def set_size(self, text):
        try:
            v = float(text)
        except ValueError:
            return
        if not 2 <= v <= 400:
            return
        self.apply(lambda s: s._replace(size=None if abs(v - 11) < 0.01 else v), "Font size")
        self.grid.autofit_rows([r for r in self.sel_rows() if r not in self.grid.sheet.rowh][:500])
        self.grid.setFocus()

    def decimals(self, d):
        r, c = self.grid.cur
        fmt = self.wb.style_of(self.grid.sheet, r, c).fmt

        def f(s):
            code = s.fmt
            if code == "General":
                code = "0.00" if d > 0 else "0"
                return s._replace(fmt=code)
            m = re.search(r"0(\.(0+))?", code)
            if not m:
                return s
            n = len(m.group(2) or "")
            n = max(0, min(15, n + d))
            new = "0" + ("." + "0" * n if n else "")
            return s._replace(fmt=code[:m.start()] + new + code[m.end():])
        self.apply(f, "Decimals")

    def fmt_chosen(self, i):
        code = self.fmt_box.itemData(i)
        if code == "__more__":
            self.format_dialog()
        else:
            self.apply(lambda s: s._replace(fmt=code), "Number format")

    def format_dialog(self):
        r, c = self.grid.cur
        st = self.wb.style_of(self.grid.sheet, r, c)
        d = FormatCells(self, st, self.wb.cell_value(self.grid.sheet.name, r, c))
        if d.exec():
            fn = d.result()
            self.apply(fn, "Format cells")
            self.grid.autofit_rows([r_ for r_ in self.sel_rows() if r_ not in self.grid.sheet.rowh][:500])

    def borders(self, mode):
        col = "#000000"
        cells = set(self.sel_cells())
        r1, c1, r2, c2 = self.grid.sel[-1]
        sh = self.grid.sheet.name

        def mk(r, c):
            def f(s):
                if mode == "none":
                    return s._replace(bl=None, bt=None, br=None, bb=None)
                if mode == "all":
                    return s._replace(bl=col, bt=col, br=col, bb=col)
                kw = {}
                if mode in ("outer",):
                    if c == c1:
                        kw["bl"] = col
                    if c == c2:
                        kw["br"] = col
                    if r == r1:
                        kw["bt"] = col
                    if r == r2:
                        kw["bb"] = col
                elif mode == "bottom" and r == r2:
                    kw["bb"] = col
                elif mode == "top" and r == r1:
                    kw["bt"] = col
                elif mode == "left" and c == c1:
                    kw["bl"] = col
                elif mode == "right" and c == c2:
                    kw["br"] = col
                return s._replace(**kw)
            return f
        for (r, c) in cells:
            self.wb.set_style(sh, [(r, c)], mk(r, c), "Borders")

    def merge_toggle(self):
        g = self.grid
        sh = g.sheet
        r1, c1, r2, c2 = g.sel[-1]
        existing = [m for m in sh.merges if m[0] <= r2 and m[2] >= r1 and m[1] <= c2 and m[3] >= c1]
        if existing:
            self.wb.set_prop(sh.name, "merges", [m for m in sh.merges if m not in existing], "Unmerge")
        elif (r1, c1) != (r2, c2):
            if r2 >= ROWS - 1 or c2 >= COLS - 1:
                return
            self.wb.set_prop(sh.name, "merges", sh.merges + [(r1, c1, r2, c2)], "Merge")
            self.apply(lambda s: s._replace(h=s.h or "center", v=s.v or "middle"), "Merge")
        g.reload_axes()
        g.set_selection(r1, c1, r2, c2, cur=(r1, c1))
        self.update_state()

    def paint_format(self):
        g = self.grid
        if self.a_paintfmt.isChecked():
            r, c = g.cur
            self.format_painter = self.wb.style_of(g.sheet, r, c)
            self.stat_msg.setText("Select the cells to paint")
            g.selection_changed.connect(self._paint_apply)
        else:
            self._paint_done()

    def _paint_apply(self):
        if self.format_painter is not None and self.grid.editor is None:
            st = self.format_painter
            self.wb.set_style(self.grid.sheet.name, self.sel_cells(), lambda s: st, "Paint format")
            self._paint_done()

    def _paint_done(self):
        try:
            self.grid.selection_changed.disconnect(self._paint_apply)
        except (RuntimeError, TypeError):
            pass
        self.format_painter = None
        self.a_paintfmt.setChecked(False)
        self.stat_msg.setText("Ready")

    def sel_rows(self):
        out = []
        for (r1, c1, r2, c2) in self.grid.sel:
            out.extend(range(r1, min(r2, r1 + 5000) + 1))
        return sorted(set(out))

    def sel_columns(self):
        out = []
        for (r1, c1, r2, c2) in self.grid.sel:
            out.extend(range(c1, min(c2, c1 + 2000) + 1))
        return sorted(set(out))

    def insert_rows(self, below):
        r1, c1, r2, c2 = self.grid.sel[-1]
        n = r2 - r1 + 1
        self.wb.insert(self.grid.sheet.name, "row", (r2 + 1) if below else r1, n)
        self.after_history()

    def insert_cols(self, right):
        r1, c1, r2, c2 = self.grid.sel[-1]
        n = c2 - c1 + 1
        self.wb.insert(self.grid.sheet.name, "col", (c2 + 1) if right else c1, n)
        self.after_history()

    def delete_rows(self):
        r1, c1, r2, c2 = self.grid.sel[-1]
        self.wb.delete(self.grid.sheet.name, "row", r1, min(r2, ROWS - 1) - r1 + 1)
        self.after_history()

    def delete_cols(self):
        r1, c1, r2, c2 = self.grid.sel[-1]
        self.wb.delete(self.grid.sheet.name, "col", c1, min(c2, COLS - 1) - c1 + 1)
        self.after_history()

    def hide(self, axis, on):
        sh = self.grid.sheet
        if axis == "row":
            s = set(sh.hidden_rows) | set(self.sel_rows())
            self.wb.set_prop(sh.name, "hidden_rows", s, "Hide rows")
        else:
            s = set(sh.hidden_cols) | set(self.sel_columns())
            self.wb.set_prop(sh.name, "hidden_cols", s, "Hide columns")
        self.grid.reload_axes()
        self.grid.viewport().update()

    def unhide_all(self):
        sh = self.grid.sheet
        self.wb.set_prop(sh.name, "hidden_rows", set(), "Unhide")
        self.wb.set_prop(sh.name, "hidden_cols", set(), "Unhide")
        self.grid.reload_axes()

    def col_width_dialog(self):
        c = self.grid.cur[1]
        v, ok = QInputDialog.getInt(self, "Column width", "Width in pixels:", int(self.grid.ax_c.size(c)), 6, 1000)
        if ok:
            new = dict(self.grid.sheet.colw)
            for cc in self.sel_columns():
                new[cc] = v
            self.wb.set_prop(self.grid.sheet.name, "colw", new, "Column width")
            self.grid.reload_axes()

    def row_height_dialog(self):
        r = self.grid.cur[0]
        v, ok = QInputDialog.getInt(self, "Row height", "Height in pixels:", int(self.grid.ax_r.size(r)), 6, 1000)
        if ok:
            new = dict(self.grid.sheet.rowh)
            for rr in self.sel_rows():
                new[rr] = v
            self.wb.set_prop(self.grid.sheet.name, "rowh", new, "Row height")
            self.grid.reload_axes()

    def freeze(self, rows, cols):
        self.wb.set_prop(self.grid.sheet.name, "freeze", (rows, cols), "Freeze panes")
        self.grid.sx = self.grid.sy = 0
        self.grid._update_scrollbars()
        self.grid.viewport().update()

    def toggle_gridlines(self):
        self.wb.set_prop(self.grid.sheet.name, "grid", not self.grid.sheet.grid, "Gridlines")
        self.grid.viewport().update()

    def toggle_show_formulas(self):
        self.grid.show_formulas = self.a_formulas.isChecked()
        self.grid.viewport().update()

    def edit_note(self):
        r, c = self.grid.cur
        sh = self.grid.sheet
        text, ok = QInputDialog.getMultiLineText(self, "Note", f"Note for {F.cell_name(r, c)} (empty to remove):", sh.comments.get((r, c), ""))
        if ok:
            d = dict(sh.comments)
            if text.strip():
                d[(r, c)] = text
            else:
                d.pop((r, c), None)
            self.wb.set_prop(sh.name, "comments", d, "Note")
            self.grid.viewport().update()

    # ------------------------------------------------------------------ functions / sums
    def insert_function(self):
        d = InsertFunction(self)
        if d.exec() and d.chosen():
            g = self.grid
            name = d.chosen()
            if g.editor is None:
                g.begin_edit(initial="=" + name + "(")
            else:
                g.editor.insert(name + "(")
            self.formula.setText(g.editor.text())

    def autosum(self, fn):
        g = self.grid
        r1, c1, r2, c2 = g.sel[-1]
        sh = g.sheet
        wb = self.wb

        def numeric(r, c):
            cell = sh.cells.get((r, c))
            if not cell or cell[0] is None:
                return False
            v = wb.cell_value(sh.name, r, c)
            return isinstance(v, (int, float)) and not isinstance(v, bool)
        items = []
        if (r1, c1) != (r2, c2):
            # selection: put the result under each column
            below = r2 + 1
            for c in range(c1, c2 + 1):
                items.append((below, c, f"={fn}({F.cell_name(r1, c)}:{F.cell_name(r2, c)})"))
            if all(not numeric(r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)):
                items = []
            else:
                wb.set_values(sh.name, items, "AutoSum")
                g.go_to(below, c1)
                return
        r, c = g.cur
        top = r - 1
        while top >= 0 and numeric(top, c):
            top -= 1
        if top < r - 1:
            f = f"={fn}({F.cell_name(top + 1, c)}:{F.cell_name(r - 1, c)})"
        else:
            left = c - 1
            while left >= 0 and numeric(r, left):
                left -= 1
            if left < c - 1:
                f = f"={fn}({F.cell_name(r, left + 1)}:{F.cell_name(r, c - 1)})"
            else:
                f = f"={fn}()"
        g.begin_edit(initial=f)
        if f.endswith("()"):
            g.editor.setCursorPosition(len(f) - 1)

    # ------------------------------------------------------------------ sort / filter / data
    def current_region(self, r, c):
        cells = self.grid.sheet.cells

        def used(rr, cc):
            cell = cells.get((rr, cc))
            return cell is not None and cell[0] not in (None, "")
        r1 = r2 = r
        c1 = c2 = c
        changed = True
        while changed:
            changed = False
            if r1 > 0 and any(used(r1 - 1, cc) for cc in range(max(0, c1 - 1), c2 + 2)):
                r1 -= 1
                changed = True
            if any(used(r2 + 1, cc) for cc in range(max(0, c1 - 1), c2 + 2)):
                r2 += 1
                changed = True
            if c1 > 0 and any(used(rr, c1 - 1) for rr in range(max(0, r1 - 1), r2 + 2)):
                c1 -= 1
                changed = True
            if any(used(rr, c2 + 1) for rr in range(max(0, r1 - 1), r2 + 2)):
                c2 += 1
                changed = True
        return r1, c1, r2, c2

    def sort_region(self):
        g = self.grid
        r1, c1, r2, c2 = g.sel[-1]
        if (r1, c1) == (r2, c2):
            r1, c1, r2, c2 = self.current_region(*g.cur)
        ur, uc = self.wb.used_extent(g.sheet.name)
        return r1, c1, min(r2, ur), min(c2, uc)

    def guess_header(self, r1, c1, r2, c2):
        if r2 <= r1:
            return False
        vals1 = [self.wb.cell_value(self.grid.sheet.name, r1, c) for c in range(c1, c2 + 1)]
        vals2 = [self.wb.cell_value(self.grid.sheet.name, r1 + 1, c) for c in range(c1, c2 + 1)]
        return all(isinstance(v, str) for v in vals1 if v is not None) and any(isinstance(v, (int, float)) for v in vals2)

    def do_sort(self, keys, header):
        g = self.grid
        sh = g.sheet
        r1, c1, r2, c2 = self.sort_region()
        start = r1 + 1 if header else r1
        if start > r2:
            return
        rows = list(range(start, r2 + 1))

        def keyf(r, col):
            v = self.wb.cell_value(sh.name, r, c1 + col)
            if v is None or v == "":
                return None
            if isinstance(v, bool):
                return (2, int(v))
            if isinstance(v, (int, float)):
                return (0, v)
            if isinstance(v, F.XErr):
                return (3, v.code)
            return (1, str(v).lower())
        order = rows[:]
        for col, desc in reversed(keys):
            blanks = [r for r in order if keyf(r, col) is None]
            full = [r for r in order if keyf(r, col) is not None]
            full.sort(key=lambda r: keyf(r, col), reverse=desc)
            order = full + blanks
        ch = {}
        for new_r, old_r in zip(rows, order):
            for c in range(c1, c2 + 1):
                cell = sh.cells.get((old_r, c))
                if cell and isinstance(cell[0], str) and cell[0].startswith("="):
                    cell = (F.shift_formula(cell[0], new_r - old_r, 0), cell[1])
                ch[(sh.name, new_r, c)] = cell
        self.wb.apply(ch, "Sort")

    def quick_sort(self, desc):
        g = self.grid
        r1, c1, r2, c2 = self.sort_region()
        header = self.guess_header(r1, c1, r2, c2)
        col = min(max(g.cur[1], c1), c2) - c1
        self.do_sort([(col, desc)], header)

    def sort_dialog(self):
        r1, c1, r2, c2 = self.sort_region()
        header = self.guess_header(r1, c1, r2, c2)
        heads = [F.text(self.wb.cell_value(self.grid.sheet.name, r1, c)) for c in range(c1, c2 + 1)]
        d = SortDialog(self, heads, header)
        if d.exec() and d.keys():
            self.do_sort(d.keys(), d.hdr.isChecked())

    def toggle_filter(self):
        g = self.grid
        sh = g.sheet
        if sh.filter:
            def op():
                s = self.wb.sheet(sh.name)
                rows = s.filter.get("rows", set())
                s.hidden_rows = set(s.hidden_rows) - set(rows)
                s.filter = None
            self.wb.structural("Filter off", op)
        else:
            r1, c1, r2, c2 = self.sort_region()
            if not self.wb.sheet(sh.name).cells:
                return

            def op():
                self.wb.sheet(sh.name).filter = {"range": (r1, c1, r2, c2), "hidden": {}, "rows": set()}
            self.wb.structural("Filter on", op)
        self.after_history()

    def apply_filter(self, hidden_map):
        sh = self.grid.sheet
        f = sh.filter
        r1, c1, r2, c2 = f["range"]

        def op():
            s = self.wb.sheet(sh.name)
            fl = s.filter
            s.hidden_rows = set(s.hidden_rows) - set(fl.get("rows", set()))
            fl["hidden"] = {c: v for c, v in hidden_map.items() if v}
            rows = set()
            for r in range(r1 + 1, r2 + 1):
                for c, vals in fl["hidden"].items():
                    text = self.wb.display(s, r, c)[0]
                    if text in vals:
                        rows.add(r)
                        break
            fl["rows"] = rows
            s.hidden_rows = set(s.hidden_rows) | rows
        self.wb.structural("Filter", op)
        self.after_history()

    def filter_menu(self, col, pos):
        sh = self.grid.sheet
        f = sh.filter
        r1, c1, r2, c2 = f["range"]
        vals = {}
        for r in range(r1 + 1, r2 + 1):
            t = self.wb.display(sh, r, col)[0]
            vals[t] = vals.get(t, 0) + 1
        hidden = set(f["hidden"].get(col, set()))
        m = QMenu(self)
        m.addAction("Sort A to Z", lambda: self.sort_by_col(col, False))
        m.addAction("Sort Z to A", lambda: self.sort_by_col(col, True))
        m.addSeparator()
        lw = QListWidget()
        lw.setMinimumWidth(220)
        lw.setMinimumHeight(180)
        for t in sorted(vals, key=lambda x: (x == "", x.lower())):
            it = QListWidgetItem(f"{t if t else '(blank)'}  ({vals[t]})")
            it.setData(Qt.UserRole, t)
            it.setFlags(it.flags() | Qt.ItemIsUserCheckable)
            it.setCheckState(Qt.Unchecked if t in hidden else Qt.Checked)
            lw.addItem(it)
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(6, 6, 6, 6)
        row = QHBoxLayout()
        b_all, b_none = QPushButton("All"), QPushButton("None")
        row.addWidget(b_all)
        row.addWidget(b_none)
        lay.addLayout(row)
        lay.addWidget(lw)
        ok = ui.primary(QPushButton("OK"))
        lay.addWidget(ok)
        b_all.clicked.connect(lambda: [lw.item(i).setCheckState(Qt.Checked) for i in range(lw.count())])
        b_none.clicked.connect(lambda: [lw.item(i).setCheckState(Qt.Unchecked) for i in range(lw.count())])

        def done():
            new = {lw.item(i).data(Qt.UserRole) for i in range(lw.count()) if lw.item(i).checkState() == Qt.Unchecked}
            m.close()
            hm = dict(f["hidden"])
            hm[col] = new
            self.apply_filter(hm)
        ok.clicked.connect(done)
        wa = QWidgetAction(m)
        wa.setDefaultWidget(w)
        m.addAction(wa)
        m.exec(pos)

    def sort_by_col(self, col, desc):
        sh = self.grid.sheet
        r1, c1, r2, c2 = sh.filter["range"]
        self.grid.set_selection(r1, c1, r2, c2, cur=(r1, c1))
        self.do_sort([(col - c1, desc)], True)

    def remove_duplicates(self):
        g = self.grid
        r1, c1, r2, c2 = self.sort_region()
        header = self.guess_header(r1, c1, r2, c2)
        start = r1 + 1 if header else r1
        seen, keep = set(), []
        for r in range(start, r2 + 1):
            key = tuple((lambda v: v.lower() if isinstance(v, str) else v)(self.wb.cell_value(g.sheet.name, r, c)) for c in range(c1, c2 + 1))
            if key in seen:
                continue
            seen.add(key)
            keep.append(r)
        removed = (r2 - start + 1) - len(keep)
        if not removed:
            self.stat_msg.setText("No duplicate rows found")
            return
        sh = g.sheet
        ch = {}
        for i, old in enumerate(keep):
            for c in range(c1, c2 + 1):
                ch[(sh.name, start + i, c)] = sh.cells.get((old, c))
        for r in range(start + len(keep), r2 + 1):
            for c in range(c1, c2 + 1):
                ch[(sh.name, r, c)] = None
        self.wb.apply(ch, "Remove duplicates")
        self.stat_msg.setText(f"Removed {removed} duplicate row(s)")

    def text_to_columns(self):
        g = self.grid
        r1, c1, r2, c2 = g.sel[-1]
        ur, uc = self.wb.used_extent(g.sheet.name)
        r2 = min(r2, ur)
        d = TextToColumns(self)
        if not d.exec():
            return
        sep = d.sep()
        items = []
        for r in range(r1, r2 + 1):
            t = self.wb.cell_value(g.sheet.name, r, c1)
            if t is None:
                continue
            for j, part in enumerate(F.text(t).split(sep)):
                items.append((r, c1 + j, part.strip()))
        self.wb.set_values(g.sheet.name, items, "Split text")

    def names_dialog(self):
        g = self.grid
        r1, c1, r2, c2 = g.sel[-1]
        ref = f"{F.quote_sheet(g.sheet.name)}!{F.cell_name(r1, c1, True, True)}" + ("" if (r1, c1) == (r2, c2) else ":" + F.cell_name(r2, c2, True, True))
        d = NamesDialog(self, self.wb.names, ref)
        if d.exec():
            self.wb.names = d.names
            self.wb.touch(("props",))

    # ------------------------------------------------------------------ conditional formats / charts
    def cf_dialog(self):
        g = self.grid
        r1, c1, r2, c2 = g.sel[-1]
        d = CFDialog(self, range_text(r1, c1, min(r2, 500000), min(c2, 16383)), g.sheet.cf)
        if d.exec():
            self.wb.set_prop(g.sheet.name, "cf", d.rules, "Conditional format")
            g.cf.ver = -1
            g.viewport().update()

    def validation_dialog(self):
        from sheetkit import validation as V
        g = self.grid
        r1, c1, r2, c2 = g.sel[-1]
        rng = range_text(r1, c1, min(r2, 1048575), min(c2, 16383))
        existing = V.rule_at(g.sheet, r1, c1)
        d = ValidationDialog(self, rng, existing)
        if not d.exec():
            return
        rules = [dict(x) for x in g.sheet.dv]
        if existing is not None:
            rules = [x for x in rules if x.get("range") != existing["range"] or x.get("kind") != existing["kind"]]
        if not d.cleared:
            new = d.result_rule()
            if new["kind"] != "any":
                rules.append(new)
        self.wb.set_prop(g.sheet.name, "dv", rules, "Data validation")
        g.viewport().update()

    def cf_quick(self, kind, ask=True):
        g = self.grid
        r1, c1, r2, c2 = g.sel[-1]
        ur, uc = self.wb.used_extent(g.sheet.name)
        rng = range_text(r1, c1, min(r2, max(ur, r1)), min(c2, max(uc, c1)))
        a = ""
        if ask:
            a, ok = QInputDialog.getText(self, "Conditional formatting", f"Format cells in {rng} when {kind}:")
            if not ok:
                return
        rule = {"kind": kind, "range": rng, "a": a, "b": "", "fg": "#9C0006", "bg": "#FFC7CE"}
        if kind == "color scale":
            rule.update(lo="#F8696B", mid="#FFEB84", hi="#63BE7B")
        if kind == "data bars":
            rule.update(hi="#5B9BD5")
        if kind in ("top 10", "bottom 10"):
            rule["a"] = "10"
        self.wb.set_prop(g.sheet.name, "cf", g.sheet.cf + [rule], "Conditional format")
        g.cf.ver = -1
        g.viewport().update()

    def insert_chart(self):
        g = self.grid
        r1, c1, r2, c2 = g.sel[-1]
        if (r1, c1) == (r2, c2):
            r1, c1, r2, c2 = self.current_region(*g.cur)
        ur, uc = self.wb.used_extent(g.sheet.name)
        r2, c2 = min(r2, ur), min(c2, uc)
        if r2 < r1 or c2 < c1 or not g.sheet.cells:
            QMessageBox.information(self, APP_NAME, "Select the cells with your data first (with headers if you have them).")
            return
        x = g.ax_c.pos(c2 + 1) + 20
        y = g.ax_r.pos(r1)
        ch = {"kind": "column", "range": f"{F.quote_sheet(g.sheet.name)}!{range_text(r1, c1, r2, c2)}", "title": "", "x": x, "y": y, "w": 440, "h": 280, "legend": True}
        if c2 - c1 == 0 or (r2 - r1 + 1 <= 8 and c2 - c1 == 1 and False):
            ch["kind"] = "column"
        self.wb.set_prop(g.sheet.name, "charts", g.sheet.charts + [ch], "Insert chart")
        self.load_charts()
        w = self._charts[-1]
        for cw in self._charts:
            cw.selected = cw is w
        w.update()
        self.stat_msg.setText("Chart added - drag to move, drag the edges to resize, double-click to edit")

    def edit_chart(self, w):
        d = ChartDialog(self, w.ch, self.grid.sheet.name)
        if d.exec():
            d.apply()
            self.wb.dirty = True
            w.update()
            self.update_title()

    def delete_chart(self, w):
        sh = self.grid.sheet
        if w.ch in sh.charts:
            self.wb.set_prop(sh.name, "charts", [c for c in sh.charts if c is not w.ch and c != w.ch], "Delete chart")
        self.load_charts()

    # ------------------------------------------------------------------ find / goto
    def show_find(self, replace):
        if self.find_dlg is None:
            d = self.find_dlg = FindReplace(self)
            d.b_next.clicked.connect(lambda: self.find_next(1))
            d.b_prev.clicked.connect(lambda: self.find_next(-1))
            d.b_repl.clicked.connect(self.replace_one)
            d.b_all.clicked.connect(self.replace_all)
            d.find.returnPressed.connect(lambda: self.find_next(1))
        self.find_dlg.show()
        self.find_dlg.raise_()
        self.find_dlg.find.setFocus()
        self.find_dlg.find.selectAll()

    def _matcher(self):
        d = self.find_dlg
        q = d.find.text()
        if not q:
            return None
        case, whole = d.case.isChecked(), d.whole.isChecked()

        def match(s):
            a, b = (s, q) if case else (s.lower(), q.lower())
            return a == b if whole else b in a
        return match

    def _cell_text(self, sh, r, c):
        d = self.find_dlg
        cell = sh.cells.get((r, c))
        if not cell or cell[0] is None:
            return None
        if d.formulas.isChecked() and isinstance(cell[0], str):
            return cell[0]
        return self.wb.display(sh, r, c)[0]

    def find_next(self, direction):
        match = self._matcher()
        if match is None:
            return
        g = self.grid
        d = self.find_dlg
        sheets = self.wb.sheets if d.allsheets.isChecked() else [g.sheet]
        order = []
        for sh in sheets:
            keys = sorted(sh.cells.keys())
            order.extend((sh, k) for k in keys)
        if direction < 0:
            order.reverse()
        if not order:
            d.info.setText("Nothing to search")
            return
        cur = (g.sheet, g.cur)
        start = 0
        for i, (sh, k) in enumerate(order):
            if sh is cur[0] and ((k > cur[1]) if direction > 0 else (k < cur[1])):
                start = i
                break
        else:
            start = 0 if not any(sh is cur[0] for sh, _ in order) else 0
        n = len(order)
        for j in range(n):
            sh, k = order[(start + j) % n]
            t = self._cell_text(sh, *k)
            if t is not None and match(t):
                if sh is not g.sheet:
                    self.tabs.setCurrentIndex(self.wb.sheets.index(sh))
                g.go_to(*k)
                d.info.setText("")
                return
        d.info.setText("Not found")

    def replace_one(self):
        match = self._matcher()
        if match is None:
            return
        g = self.grid
        t = self._cell_text(g.sheet, *g.cur)
        if t is not None and match(t):
            self._replace_cell(g.sheet, g.cur)
        self.find_next(1)

    def _replace_cell(self, sh, k):
        d = self.find_dlg
        cell = sh.cells.get(k)
        raw = cell[0]
        txt = raw if (d.formulas.isChecked() or not isinstance(raw, (int, float))) and isinstance(raw, str) else self.wb.edit_text(sh.name, *k)
        pat = re.compile(re.escape(d.find.text()), 0 if d.case.isChecked() else re.I)
        new = d.repl.text() if d.whole.isChecked() else pat.sub(lambda m: d.repl.text(), txt)
        if d.whole.isChecked():
            new = d.repl.text()
        self.wb.set_values(sh.name, [(k[0], k[1], new)], "Replace")

    def replace_all(self):
        match = self._matcher()
        if match is None:
            return
        d = self.find_dlg
        sheets = self.wb.sheets if d.allsheets.isChecked() else [self.grid.sheet]
        n = 0
        for sh in sheets:
            for k in sorted(sh.cells):
                t = self._cell_text(sh, *k)
                if t is not None and match(t):
                    self._replace_cell(sh, k)
                    n += 1
        d.info.setText(f"Replaced {n}")

    def goto_dialog(self):
        self.name_box.setFocus()
        self.name_box.selectAll()

    # ------------------------------------------------------------------ sheets
    def add_sheet(self):
        name = self.wb.add_sheet()
        self.tabs.setCurrentIndex(self.wb.sheet_names().index(name))

    def tab_changed(self, i):
        if 0 <= i < len(self.wb.sheets):
            if self.grid.editor is not None:
                self.grid.commit_edit()
            self.grid.set_sheet(self.wb.sheets[i].name)
            self.load_charts()
            self.refresh_all()

    def tab_moved(self, frm, to):
        name = self.wb.sheets[frm].name
        self.wb.listeners.remove(self._on_change)
        self.wb.move_sheet(name, to)
        self.wb.listeners.append(self._on_change)

    def tab_menu(self, pos):
        i = self.tabs.tabAt(pos)
        if i < 0:
            return
        m = QMenu(self)
        m.addAction("Rename", lambda: self.rename_sheet(i))
        m.addAction("Duplicate", lambda: self.wb.duplicate_sheet(self.wb.sheets[i].name))
        m.addAction("Delete", lambda: self.delete_sheet(i))
        m.addAction("Tab colour ...", lambda: self.tab_color(i))
        m.addSeparator()
        m.addAction("Move left", lambda: self.move_sheet(i, -1))
        m.addAction("Move right", lambda: self.move_sheet(i, 1))
        m.exec(self.tabs.mapToGlobal(pos))

    def rename_sheet(self, i):
        old = self.wb.sheets[i].name
        new, ok = QInputDialog.getText(self, "Rename sheet", "Sheet name:", text=old)
        new = new.strip()[:31]
        if ok and new and new != old:
            if self.wb.sheet(new) and new.lower() != old.lower():
                QMessageBox.warning(self, APP_NAME, "A sheet with that name exists.")
                return
            if re.search(r"[\[\]\\/*?:]", new):
                QMessageBox.warning(self, APP_NAME, "A sheet name cannot contain  [ ] \\ / * ? :")
                return
            self.wb.rename_sheet(old, new)
            self.grid.sheet = self.wb.sheet(new) if self.grid.sheet.name == old else self.grid.sheet

    def delete_sheet(self, i):
        if len(self.wb.sheets) < 2:
            return
        if QMessageBox.question(self, APP_NAME, f"Delete sheet '{self.wb.sheets[i].name}'?") == QMessageBox.Yes:
            self.wb.delete_sheet(self.wb.sheets[i].name)
            self.tabs.setCurrentIndex(max(0, min(i, len(self.wb.sheets) - 1)))
            self.tab_changed(self.tabs.currentIndex())

    def move_sheet(self, i, d):
        j = i + d
        if 0 <= j < len(self.wb.sheets):
            self.wb.move_sheet(self.wb.sheets[i].name, j)
            self.tabs.setCurrentIndex(j)

    def tab_color(self, i):
        from PySide6.QtWidgets import QColorDialog
        c = QColorDialog.getColor(QColor(self.wb.sheets[i].tab_color or "#C2395C"), self, "Tab colour")
        if c.isValid():
            self.wb.set_prop(self.wb.sheets[i].name, "tab_color", c.name().upper(), "Tab colour")
            self.tabs.setTabTextColor(i, c)

    def show_context(self, pos, area):
        m = QMenu(self)
        g = self.grid
        if area == "cell":
            m.addAction("Cut", lambda: g.copy(cut=True))
            m.addAction("Copy", g.copy)
            m.addAction("Paste", lambda: g.paste("all"))
            ps = m.addMenu("Paste special")
            for t, mode in (("Values only", "values"), ("Formats only", "formats"), ("Transposed", "transpose")):
                ps.addAction(t, lambda mode=mode: g.paste(mode))
            m.addSeparator()
            m.addAction("Insert row above", lambda: self.insert_rows(False))
            m.addAction("Insert column left", lambda: self.insert_cols(False))
            m.addAction("Delete row", self.delete_rows)
            m.addAction("Delete column", self.delete_cols)
            m.addSeparator()
            m.addAction("Clear contents", g.clear_contents)
            m.addAction(self.a_fmt)
            m.addAction("Sort A to Z", lambda: self.quick_sort(False))
            m.addAction("Note ...", self.edit_note)
        elif area == "col":
            m.addAction("Cut", lambda: g.copy(cut=True))
            m.addAction("Copy", g.copy)
            m.addAction("Paste", lambda: g.paste("all"))
            m.addSeparator()
            m.addAction("Insert column left", lambda: self.insert_cols(False))
            m.addAction("Insert column right", lambda: self.insert_cols(True))
            m.addAction("Delete column", self.delete_cols)
            m.addAction("Clear contents", g.clear_contents)
            m.addSeparator()
            m.addAction("Hide column", lambda: self.hide("col", True))
            m.addAction("Unhide all", self.unhide_all)
            m.addAction("Column width ...", self.col_width_dialog)
            m.addAction("Auto-fit width", lambda: g.autofit_columns(self.sel_columns()))
        else:
            m.addAction("Cut", lambda: g.copy(cut=True))
            m.addAction("Copy", g.copy)
            m.addAction("Paste", lambda: g.paste("all"))
            m.addSeparator()
            m.addAction("Insert row above", lambda: self.insert_rows(False))
            m.addAction("Insert row below", lambda: self.insert_rows(True))
            m.addAction("Delete row", self.delete_rows)
            m.addAction("Clear contents", g.clear_contents)
            m.addSeparator()
            m.addAction("Hide row", lambda: self.hide("row", True))
            m.addAction("Unhide all", self.unhide_all)
            m.addAction("Row height ...", self.row_height_dialog)
            m.addAction("Auto-fit height", lambda: g.autofit_rows(self.sel_rows()))
        m.exec(pos)

    # ------------------------------------------------------------------ files
    def maybe_save(self):
        if not self.wb.dirty:
            return True
        r = QMessageBox.question(self, APP_NAME, "Save the changes first?", QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel)
        if r == QMessageBox.Save:
            return self.save()
        return r == QMessageBox.Discard

    def on_start(self):
        return hasattr(self, "stack") and self.stack.currentWidget() is self.start

    def show_start(self):
        self.start.refresh()
        self.stack.setCurrentWidget(self.start)
        self.statusBar().setVisible(False)
        self.setWindowTitle(APP_NAME)

    def show_editor(self):
        self.stack.setCurrentWidget(self.work)
        self.statusBar().setVisible(True)
        self.update_title()
        self.grid.setFocus()

    def start_new(self):
        self.grid.cancel_edit()
        self.attach(Workbook())
        self.show_editor()

    def close_document(self):
        if not self.maybe_save():
            return
        self.grid.cancel_edit()
        self.attach(Workbook())
        self.wb.dirty = False
        self.show_start()

    def new_file(self):
        if not self.on_start() and not self.maybe_save():
            return
        self.start_new()

    def open_dialog(self):
        if not self.on_start() and not self.maybe_save():
            return
        d = office.OpenDialog(self, SLUG, FILTERS)
        if d.exec() == QDialog.Accepted and d.path:
            self.open_path(d.path)

    def open_path(self, path):
        try:
            ext = os.path.splitext(path)[1].lower()
            if ext in (".xlsx", ".xlsm"):
                wb = xlsxio.read_xlsx(path)
            elif ext == ".dsht":
                wb = dsht.read_dsht(path)
            else:
                wb = xlsxio.read_csv(path)
        except Exception as e:  # noqa
            QMessageBox.critical(self, APP_NAME, f"Could not open this file.\n\n{e}")
            return
        wb.path = path if ext in (".xlsx", ".xlsm", ".dsht") else None
        self.grid.cancel_edit()
        self.attach(wb)
        wb.dirty = False
        self.show_editor()
        self.update_title()
        ui.add_recent(SLUG, path)
        self.stat_msg.setText(f"Opened {os.path.basename(path)}")

    def save(self, as_new=False):
        self.grid.commit_edit() if self.grid.editor else None
        path = self.wb.path
        if as_new or not path:
            path, flt = QFileDialog.getSaveFileName(self, "Save", path or "Untitled.xlsx", "Excel workbook (*.xlsx);;Digonto spreadsheet (*.dsht)")
            if not path:
                return False
            if not path.lower().endswith((".xlsx", ".dsht")):
                path += ".dsht" if "dsht" in flt else ".xlsx"
        try:
            if path.lower().endswith(".dsht"):
                dsht.write_dsht(self.wb, path)
            else:
                xlsxio.write_xlsx(self.wb, path)
        except Exception as e:  # noqa
            QMessageBox.critical(self, APP_NAME, f"Could not save.\n\n{e}")
            return False
        self.wb.path = path
        self.wb.dirty = False
        self.update_title()
        ui.add_recent(SLUG, path)
        extra = "  (charts are saved as Excel charts; Excel may place them a little differently)" if any(s.charts for s in self.wb.sheets) and path.lower().endswith(".xlsx") else ""
        self.stat_msg.setText(f"Saved {os.path.basename(path)}{extra}")
        return True

    def export_csv(self):
        p, _ = QFileDialog.getSaveFileName(self, "Export CSV", self.grid.sheet.name + ".csv", "CSV (*.csv);;Tab separated (*.tsv)")
        if p:
            xlsxio.write_csv(self.wb, p, self.grid.sheet.name, "\t" if p.lower().endswith(".tsv") else ",")
            self.stat_msg.setText(f"Exported {os.path.basename(p)}")

    def html_table(self):
        sh = self.grid.sheet
        ur, uc = self.wb.used_extent(sh.name)
        out = ["<table border='1' cellspacing='0' cellpadding='4' style='border-collapse:collapse; border-color:#BBB;'>"]
        for r in range(ur + 1):
            if r in sh.hidden_rows:
                continue
            out.append("<tr>")
            for c in range(uc + 1):
                if c in sh.hidden_cols:
                    continue
                text, v, st = self.wb.display(sh, r, c)
                al = st.h or ("right" if isinstance(v, (int, float)) and not isinstance(v, bool) else "left")
                css = f"text-align:{al};" + ("font-weight:bold;" if st.bold else "") + ("font-style:italic;" if st.italic else "") + (f"background-color:{st.bg};" if st.bg else "") + (f"color:{st.color};" if st.color else "") \
                      + (f"font-size:{st.size}pt;" if st.size else "") + f"width:{int(sh.colw.get(c, DEFAULT_W) * 0.75)}pt;"
                t = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br>") or "&nbsp;"
                out.append(f"<td style='{css}'>{t}</td>")
            out.append("</tr>")
        out.append("</table>")
        return "".join(out)

    def print_sheet(self):
        from PySide6.QtPrintSupport import QPrintDialog, QPrinter
        pr = QPrinter(QPrinter.HighResolution)
        pr.setPageOrientation(QPageLayout.Landscape)
        pr.setDocName(self.grid.sheet.name)
        if QPrintDialog(pr, self).exec() != QDialog.Accepted:
            return
        doc = QTextDocument()
        doc.setHtml(f"<h3>{self.grid.sheet.name}</h3>" + self.html_table())
        doc.print_(pr)

    def export_pdf(self):
        p, _ = QFileDialog.getSaveFileName(self, "Export PDF", (os.path.splitext(self.wb.path)[0] if self.wb.path else self.grid.sheet.name) + ".pdf", "PDF (*.pdf)")
        if not p:
            return
        w = QPdfWriter(p)
        w.setPageSize(QPageSize(QPageSize.A4))
        w.setPageOrientation(QPageLayout.Landscape)
        doc = QTextDocument()
        doc.setHtml(f"<h3>{self.grid.sheet.name}</h3>" + self.html_table())
        doc.print_(w)
        self.stat_msg.setText(f"Saved {os.path.basename(p)}")

    def export_html(self):
        p, _ = QFileDialog.getSaveFileName(self, "Export HTML", self.grid.sheet.name + ".html", "HTML (*.html)")
        if p:
            with open(p, "w", encoding="utf-8") as f:
                f.write("<!doctype html><meta charset='utf-8'><title>" + self.grid.sheet.name + "</title>" + self.html_table())

    # ------------------------------------------------------------------ misc
    def shortcuts_help(self):
        QMessageBox.information(self, "Keyboard shortcuts", (
            "<b>Move</b>: arrows, Ctrl+arrows (to the edge of data), Home, Ctrl+Home, Ctrl+End, PgUp / PgDn<br>"
            "<b>Select</b>: Shift+arrows, Ctrl+A, Ctrl+Space (column), Shift+Space (row), Ctrl+click<br>"
            "<b>Edit</b>: type to replace, F2 to edit, Enter / Tab to confirm and move, Esc to cancel, Alt+Enter new line in cell<br>"
            "<b>Formulas</b>: start with =, click cells to add references, F4 toggles $, Tab accepts a function suggestion, Alt+= AutoSum<br>"
            "<b>Copy</b>: Ctrl+C / X / V, Ctrl+Shift+V values only, Ctrl+D fill down, Ctrl+R fill right, drag the small square to fill<br>"
            "<b>Format</b>: Ctrl+B / I / U, Ctrl+1 format cells<br>"
            "<b>Other</b>: Ctrl+Z / Y undo redo, Ctrl+F find, Ctrl+H replace, Ctrl+G go to, Ctrl+; today's date, Ctrl+wheel zoom"))

    def about(self):
        QMessageBox.about(self, f"About {APP_NAME}", f"<h3>{APP_NAME} {VERSION}</h3><p>Spreadsheets with 170+ formulas, charts, conditional formats, filters, "
                          "XLSX and CSV. Built with PySide6 only.<p>Charts are kept while the file is open; the XLSX file stores cells, styles, "
                          "widths, merges, frozen panes, names and basic conditional formats.")

    def closeEvent(self, e):
        if not self.maybe_save():
            e.ignore()
            return
        e.accept()
        ui.force_exit_soon()


def main():
    if "--version" in sys.argv:
        print(VERSION)
        return
    ui.mac_app_name(APP_NAME)
    app = QApplication(sys.argv)
    if ui.forward_to_running(SLUG, sys.argv[1:]):
        os._exit(0)
    ui.setup_app(app, APP_NAME, SLUG, ui.app_icon(HERE, SLUG), theme=True)
    path = next((a for a in sys.argv[1:] if os.path.isfile(a)), None)
    w = SheetsWindow(path)
    w.show()
    _extra = []

    def _new_window(p):
        x = SheetsWindow(p)
        x.show()
        _extra.append(x)
        return x
    ui.serve_single(SLUG, w, _new_window)
    os._exit(app.exec())


if __name__ == "__main__":
    main()
