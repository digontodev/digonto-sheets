"""Dialogs for Digonto Sheets."""
from PySide6.QtCore import QPoint, Qt, Signal
from PySide6.QtGui import QColor, QFont, QPixmap, QPainter
from PySide6.QtWidgets import (QCheckBox, QColorDialog, QComboBox, QDialog, QDialogButtonBox, QDoubleSpinBox, QFormLayout, QFrame, QGridLayout,
                               QHBoxLayout, QLabel, QLineEdit, QListWidget, QListWidgetItem, QMenu, QPlainTextEdit, QPushButton, QRadioButton,
                               QSpinBox, QTabWidget, QToolButton, QVBoxLayout, QWidget, QWidgetAction)

from . import formula as F
from .cfmt import CF, parse_range, range_text
from .model import CF_KINDS
from .numfmt import FORMATS, format_value

SWATCHES = ["#000000", "#434343", "#666666", "#999999", "#B7B7B7", "#CCCCCC", "#D9D9D9", "#EFEFEF", "#F3F3F3", "#FFFFFF",
            "#980000", "#FF0000", "#FF9900", "#FFFF00", "#00FF00", "#00FFFF", "#4A86E8", "#0000FF", "#9900FF", "#FF00FF",
            "#E6B8AF", "#F4CCCC", "#FCE5CD", "#FFF2CC", "#D9EAD3", "#D0E0E3", "#C9DAF8", "#CFE2F3", "#D9D2E9", "#EAD1DC",
            "#DD7E6B", "#EA9999", "#F9CB9C", "#FFE599", "#B6D7A8", "#A2C4C9", "#A4C2F4", "#9FC5E8", "#B4A7D6", "#D5A6BD",
            "#CC4125", "#E06666", "#F6B26B", "#FFD966", "#93C47D", "#76A5AF", "#6D9EEB", "#6FA8DC", "#8E7CC3", "#C27BA0",
            "#C2395C", "#1E7B45", "#2B5FB3", "#E0662A", "#7A1F1F", "#274E13", "#0C343D", "#1C4587", "#20124D", "#4C1130"]


class ColorButton(QToolButton):
    """toolbar button with a colour bar underneath and a palette drop down"""
    picked = Signal(object)     # "#RRGGBB" or None (automatic / none)

    def __init__(self, icon_fn, tip, default=None, none_label="None", parent=None):
        super().__init__(parent)
        self.icon_fn, self.color, self.none_label = icon_fn, default, none_label
        self.setToolTip(tip)
        self.setPopupMode(QToolButton.MenuButtonPopup)
        self.menu_ = QMenu(self)
        self.setMenu(self.menu_)
        w = QWidget()
        g = QGridLayout(w)
        g.setSpacing(2)
        g.setContentsMargins(6, 6, 6, 6)
        for i, col in enumerate(SWATCHES):
            b = QToolButton()
            b.setFixedSize(18, 18)
            b.setStyleSheet(f"QToolButton{{background:{col}; border:1px solid #B9C4BE;}} QToolButton:hover{{border:2px solid #C2395C;}}")
            b.setToolTip(col)
            b.clicked.connect(lambda _=False, c=col: self._pick(c))
            g.addWidget(b, i // 10, i % 10)
        act = QWidgetAction(self.menu_)
        act.setDefaultWidget(w)
        self.menu_.addAction(none_label, lambda: self._pick(None))
        self.menu_.addAction(act)
        self.menu_.addAction("More colours ...", self._more)
        self.clicked.connect(lambda: self.picked.emit(self.color))
        self.refresh()

    def _pick(self, c):
        self.color = c
        self.menu_.close()
        self.refresh()
        self.picked.emit(c)

    def _more(self):
        c = QColorDialog.getColor(QColor(self.color or "#000000"), self, "Colour")
        if c.isValid():
            self._pick(c.name().upper())

    def refresh(self):
        from .icons import icon
        pm = icon(self.icon_fn).pixmap(22, 22)
        out = QPixmap(26, 26)
        out.fill(Qt.transparent)
        p = QPainter(out)
        p.drawPixmap(2, 0, pm)
        p.fillRect(3, 21, 20, 4, QColor(self.color) if self.color else QColor("#FFFFFF"))
        if not self.color:
            p.setPen(QColor("#B9C4BE"))
            p.drawRect(3, 21, 19, 3)
        p.end()
        from PySide6.QtGui import QIcon
        self.setIcon(QIcon(out))


class InsertFunction(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Insert function")
        self.resize(560, 440)
        lay = QVBoxLayout(self)
        top = QHBoxLayout()
        self.search = QLineEdit()
        self.search.setPlaceholderText("Search a function ... (sum, lookup, date, text)")
        self.cat = QComboBox()
        self.cat.addItem("All")
        self.cat.addItems(F.CATEGORIES)
        top.addWidget(self.search, 1)
        top.addWidget(self.cat)
        lay.addLayout(top)
        self.lst = QListWidget()
        lay.addWidget(self.lst, 1)
        self.sig = QLabel()
        self.sig.setWordWrap(True)
        self.sig.setStyleSheet("background:#EAF4EF; border-radius:6px; padding:8px;")
        lay.addWidget(self.sig)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)
        self.search.textChanged.connect(self.fill)
        self.cat.currentIndexChanged.connect(self.fill)
        self.lst.currentItemChanged.connect(self.show_sig)
        self.lst.itemDoubleClicked.connect(lambda _: self.accept())
        self.fill()

    def fill(self):
        q = self.search.text().strip().lower()
        cat = self.cat.currentText()
        self.lst.clear()
        for name in sorted(F.HELP):
            sig, doc, c = F.HELP[name]
            if cat != "All" and c != cat:
                continue
            if q and q not in name.lower() and q not in doc.lower():
                continue
            self.lst.addItem(name)
        if self.lst.count():
            self.lst.setCurrentRow(0)

    def show_sig(self, item):
        if item:
            sig, doc, c = F.HELP[item.text()]
            self.sig.setText(f"<b>{sig}</b><br>{doc}")

    def chosen(self):
        it = self.lst.currentItem()
        return it.text() if it else None


class FindReplace(QDialog):
    def __init__(self, parent, replace=False):
        super().__init__(parent)
        self.setWindowTitle("Find and replace")
        self.setWindowFlag(Qt.Tool, True)
        lay = QFormLayout(self)
        self.find = QLineEdit()
        self.repl = QLineEdit()
        lay.addRow("Find", self.find)
        lay.addRow("Replace with", self.repl)
        self.case = QCheckBox("Match case")
        self.whole = QCheckBox("Match entire cell")
        self.formulas = QCheckBox("Look in formulas")
        self.allsheets = QCheckBox("All sheets")
        row = QHBoxLayout()
        for w in (self.case, self.whole, self.formulas, self.allsheets):
            row.addWidget(w)
        lay.addRow(row)
        self.info = QLabel("")
        lay.addRow(self.info)
        btns = QHBoxLayout()
        self.b_next = QPushButton("Find next")
        self.b_prev = QPushButton("Previous")
        self.b_repl = QPushButton("Replace")
        self.b_all = QPushButton("Replace all")
        for b in (self.b_prev, self.b_next, self.b_repl, self.b_all):
            btns.addWidget(b)
        lay.addRow(btns)
        self.b_next.setDefault(True)


class SortDialog(QDialog):
    def __init__(self, parent, headers, has_header):
        super().__init__(parent)
        self.setWindowTitle("Sort range")
        lay = QVBoxLayout(self)
        self.hdr = QCheckBox("Data has a header row")
        self.hdr.setChecked(has_header)
        lay.addWidget(self.hdr)
        self.rows = []
        grid = QGridLayout()
        self.headers = headers
        for i in range(3):
            lbl = QLabel("Sort by" if i == 0 else "Then by")
            cb = QComboBox()
            order = QComboBox()
            order.addItems(["A to Z (smallest first)", "Z to A (largest first)"])
            grid.addWidget(lbl, i, 0)
            grid.addWidget(cb, i, 1)
            grid.addWidget(order, i, 2)
            self.rows.append((cb, order))
        lay.addLayout(grid)
        self.hdr.toggled.connect(self.fill)
        self.fill()
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)

    def fill(self):
        for i, (cb, _) in enumerate(self.rows):
            cb.clear()
            if i:
                cb.addItem("(none)", -1)
            for j, h in enumerate(self.headers):
                cb.addItem(h if self.hdr.isChecked() and h else f"Column {F.col_name(j)}", j)

    def keys(self):
        out = []
        for cb, order in self.rows:
            idx = cb.currentData()
            if idx is not None and idx >= 0:
                out.append((idx, order.currentIndex() == 1))
        return out


class CFDialog(QDialog):
    """add a conditional format rule / manage rules"""

    def __init__(self, parent, rng_text, rules):
        super().__init__(parent)
        self.setWindowTitle("Conditional formatting")
        self.resize(520, 420)
        self.rules = [dict(r) for r in rules]
        lay = QVBoxLayout(self)
        form = QFormLayout()
        self.rng = QLineEdit(rng_text)
        self.kind = QComboBox()
        self.kind.addItems(CF_KINDS + ["color scale", "data bars"])
        self.a = QLineEdit()
        self.b = QLineEdit()
        form.addRow("Apply to range", self.rng)
        form.addRow("Format cells if", self.kind)
        form.addRow("Value / text", self.a)
        form.addRow("and (between)", self.b)
        lay.addLayout(form)
        row = QHBoxLayout()
        self.preset = QComboBox()
        self.presets = [("Light red fill, dark red text", "#9C0006", "#FFC7CE"), ("Yellow fill, dark yellow text", "#9C5700", "#FFEB9C"),
                        ("Green fill, dark green text", "#006100", "#C6EFCE"), ("Blue fill, dark blue text", "#1F4E79", "#DDEBF7"), ("Red text", "#C00000", ""),
                        ("Bold white on green", "#FFFFFF", "#C2395C")]
        for t, _, _ in self.presets:
            self.preset.addItem(t)
        row.addWidget(QLabel("Format"))
        row.addWidget(self.preset, 1)
        lay.addLayout(row)
        add = QPushButton("Add rule")
        add.clicked.connect(self.add_rule)
        lay.addWidget(add)
        lay.addWidget(QLabel("Rules on this sheet:"))
        self.lst = QListWidget()
        lay.addWidget(self.lst, 1)
        dl = QPushButton("Delete selected rule")
        dl.clicked.connect(self.del_rule)
        lay.addWidget(dl)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)
        self.kind.currentTextChanged.connect(self.kind_changed)
        self.kind_changed(self.kind.currentText())
        self.refresh()

    def kind_changed(self, k):
        self.a.setEnabled(k in ("greater than", "less than", "equal to", "not equal to", "between", "text contains", "top 10", "bottom 10"))
        self.b.setEnabled(k == "between")
        self.preset.setEnabled(k not in ("color scale", "data bars"))
        self.a.setPlaceholderText("how many (default 10)" if "10" in k else "")

    def add_rule(self):
        k = self.kind.currentText()
        _, fg, bg = self.presets[self.preset.currentIndex()]
        rule = {"kind": k, "range": self.rng.text().strip(), "a": self.a.text().strip(), "b": self.b.text().strip(), "fg": fg, "bg": bg}
        if k == "color scale":
            rule.update(lo="#F8696B", mid="#FFEB84", hi="#63BE7B")
        if k == "data bars":
            rule.update(hi="#5B9BD5")
        if k == "Bold white on green" or (fg == "#FFFFFF"):
            rule["bold"] = True
        self.rules.append(rule)
        self.refresh()

    def del_rule(self):
        i = self.lst.currentRow()
        if i >= 0:
            self.rules.pop(i)
            self.refresh()

    def refresh(self):
        self.lst.clear()
        for r in self.rules:
            self.lst.addItem(f"{r['range']}:  {r['kind']} {r.get('a', '')} {r.get('b', '')}".strip())


class ChartDialog(QDialog):
    def __init__(self, parent, ch, sheet_name):
        super().__init__(parent)
        self.setWindowTitle("Chart")
        self.ch = ch
        lay = QFormLayout(self)
        from .charts import KINDS
        self.kind = QComboBox()
        for k, label in KINDS:
            self.kind.addItem(label, k)
        self.kind.setCurrentIndex([k for k, _ in KINDS].index(ch["kind"]))
        self.rng = QLineEdit(ch["range"])
        self.title = QLineEdit(ch.get("title", ""))
        self.by_rows = QCheckBox("Series are in rows")
        self.by_rows.setChecked(bool(ch.get("by_rows")))
        self.stacked = QCheckBox("Stacked")
        self.stacked.setChecked(bool(ch.get("stacked")))
        self.legend = QCheckBox("Show legend")
        self.legend.setChecked(ch.get("legend", True))
        lay.addRow("Type", self.kind)
        lay.addRow("Data range", self.rng)
        lay.addRow("Title", self.title)
        lay.addRow(self.by_rows)
        lay.addRow(self.stacked)
        lay.addRow(self.legend)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        lay.addRow(bb)

    def apply(self):
        self.ch.update(kind=self.kind.currentData(), range=self.rng.text().strip(), title=self.title.text(), by_rows=self.by_rows.isChecked(),
                       stacked=self.stacked.isChecked(), legend=self.legend.isChecked())


class FormatCells(QDialog):
    """Number / Alignment / Font / Border / Fill in one place; returns a function Style -> Style"""

    def __init__(self, parent, style, sample):
        super().__init__(parent)
        self.setWindowTitle("Format cells")
        self.resize(500, 460)
        self.st = style
        self.sample = sample
        lay = QVBoxLayout(self)
        tabs = QTabWidget()
        lay.addWidget(tabs)
        # number
        w = QWidget()
        nl = QVBoxLayout(w)
        self.fmt_list = QListWidget()
        for label, code in FORMATS:
            it = QListWidgetItem(label)
            it.setData(Qt.UserRole, code)
            self.fmt_list.addItem(it)
        nl.addWidget(self.fmt_list, 1)
        self.code = QLineEdit(style.fmt)
        nl.addWidget(QLabel("Format code (like Excel):"))
        nl.addWidget(self.code)
        self.preview = QLabel("")
        self.preview.setStyleSheet("background:#EAF4EF; padding:6px; border-radius:6px;")
        nl.addWidget(self.preview)
        tabs.addTab(w, "Number")
        self.fmt_list.currentItemChanged.connect(lambda it: it and self.code.setText(it.data(Qt.UserRole)))
        self.code.textChanged.connect(self.upd)
        for i in range(self.fmt_list.count()):
            if self.fmt_list.item(i).data(Qt.UserRole) == style.fmt:
                self.fmt_list.setCurrentRow(i)
        self.upd()
        # alignment
        w = QWidget()
        al = QFormLayout(w)
        self.h = QComboBox()
        self.h.addItems(["General", "Left", "Center", "Right"])
        self.h.setCurrentIndex({None: 0, "left": 1, "center": 2, "right": 3}.get(style.h, 0))
        self.v = QComboBox()
        self.v.addItems(["Bottom", "Middle", "Top"])
        self.v.setCurrentIndex({None: 0, "bottom": 0, "middle": 1, "top": 2}.get(style.v, 0))
        self.wrap = QCheckBox("Wrap text")
        self.wrap.setChecked(style.wrap)
        self.indent = QSpinBox()
        self.indent.setRange(0, 15)
        self.indent.setValue(style.indent)
        al.addRow("Horizontal", self.h)
        al.addRow("Vertical", self.v)
        al.addRow(self.wrap)
        al.addRow("Indent", self.indent)
        tabs.addTab(w, "Alignment")
        # font
        w = QWidget()
        fl = QFormLayout(w)
        self.font = QComboBox()
        self.font.setEditable(True)
        from PySide6.QtGui import QFontDatabase
        self.font.addItem("(default)")
        self.font.addItems(sorted(QFontDatabase.families()))
        self.font.setCurrentText(style.font or "(default)")
        self.size = QDoubleSpinBox()
        self.size.setRange(4, 200)
        self.size.setValue(style.size or 11)
        self.bold, self.italic, self.under, self.strike = (QCheckBox(t) for t in ("Bold", "Italic", "Underline", "Strikethrough"))
        for cb, val in ((self.bold, style.bold), (self.italic, style.italic), (self.under, style.underline), (self.strike, style.strike)):
            cb.setChecked(val)
        self.fg = ColorButton("textcolor", "Text colour", style.color, "Automatic")
        fl.addRow("Font", self.font)
        fl.addRow("Size", self.size)
        for cb in (self.bold, self.italic, self.under, self.strike):
            fl.addRow(cb)
        fl.addRow("Colour", self.fg)
        tabs.addTab(w, "Font")
        # border + fill
        w = QWidget()
        bl = QFormLayout(w)
        self.bcol = ColorButton("borders", "Border colour", style.bl or style.bt or style.br or style.bb or "#000000", "None")
        self.sides = {}
        for key, label, cur in (("bl", "Left", style.bl), ("bt", "Top", style.bt), ("br", "Right", style.br), ("bb", "Bottom", style.bb)):
            cb = QCheckBox(label)
            cb.setChecked(bool(cur))
            self.sides[key] = cb
            bl.addRow(cb)
        bl.addRow("Border colour", self.bcol)
        self.bg = ColorButton("fill", "Fill colour", style.bg, "No fill")
        bl.addRow("Fill colour", self.bg)
        tabs.addTab(w, "Border and fill")
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)

    def upd(self):
        try:
            self.preview.setText("Preview:  " + format_value(self.sample if self.sample is not None else 1234.5, self.code.text() or "General"))
        except Exception:
            self.preview.setText("Preview:  ?")

    def result(self):
        col = self.bcol.color or "#000000"
        return lambda s: s._replace(
            fmt=self.code.text() or "General", h={0: None, 1: "left", 2: "center", 3: "right"}[self.h.currentIndex()],
            v={0: None, 1: "middle", 2: "top"}[self.v.currentIndex()], wrap=self.wrap.isChecked(), indent=self.indent.value(),
            font=None if self.font.currentText() in ("(default)", "") else self.font.currentText(), size=None if abs(self.size.value() - 11) < 0.01 else self.size.value(),
            bold=self.bold.isChecked(), italic=self.italic.isChecked(), underline=self.under.isChecked(), strike=self.strike.isChecked(),
            color=self.fg.color, bg=self.bg.color,
            **{k: (col if cb.isChecked() else None) for k, cb in self.sides.items()})


class TextToColumns(QDialog):
    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle("Split text into columns")
        lay = QFormLayout(self)
        self.delim = QComboBox()
        self.delim.addItems(["Comma", "Semicolon", "Tab", "Space", "Custom"])
        self.custom = QLineEdit()
        self.custom.setMaxLength(5)
        lay.addRow("Separator", self.delim)
        lay.addRow("Custom", self.custom)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        lay.addRow(bb)

    def sep(self):
        return {0: ",", 1: ";", 2: "\t", 3: " "}.get(self.delim.currentIndex()) or self.custom.text() or ","


class NamesDialog(QDialog):
    def __init__(self, parent, names, current_ref):
        super().__init__(parent)
        self.setWindowTitle("Named ranges")
        self.names = dict(names)
        lay = QVBoxLayout(self)
        self.lst = QListWidget()
        lay.addWidget(self.lst, 1)
        form = QFormLayout()
        self.name = QLineEdit()
        self.ref = QLineEdit(current_ref)
        form.addRow("Name", self.name)
        form.addRow("Refers to", self.ref)
        lay.addLayout(form)
        row = QHBoxLayout()
        add = QPushButton("Add / update")
        rem = QPushButton("Delete")
        row.addWidget(add)
        row.addWidget(rem)
        lay.addLayout(row)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)
        add.clicked.connect(self.add)
        rem.clicked.connect(self.rem)
        self.lst.currentItemChanged.connect(lambda it: it and (self.name.setText(it.text().split("  =  ")[0]), self.ref.setText(self.names.get(it.text().split("  =  ")[0], ""))))
        self.fill()

    def fill(self):
        self.lst.clear()
        for n, t in sorted(self.names.items()):
            self.lst.addItem(f"{n}  =  {t}")

    def add(self):
        n = self.name.text().strip().upper().replace(" ", "_")
        if n and self.ref.text().strip():
            self.names[n] = self.ref.text().strip().lstrip("=")
            self.fill()

    def rem(self):
        it = self.lst.currentItem()
        if it:
            self.names.pop(it.text().split("  =  ")[0], None)
            self.fill()


class ValidationDialog(QDialog):
    """Data > Data validation: what may be typed into the selected cells"""

    def __init__(self, parent, rng_text, rule):
        super().__init__(parent)
        from . import validation as V
        self.V = V
        self.setWindowTitle("Data validation")
        self.resize(520, 460)
        self.cleared = False
        self.rule = dict(rule) if rule else V.new_rule(rng_text)
        self.rule["range"] = rng_text if not rule else rule["range"]
        lay = QVBoxLayout(self)
        tabs = QTabWidget()
        lay.addWidget(tabs, 1)
        # --- criteria
        t1 = QWidget()
        f = QFormLayout(t1)
        self.rng = QLineEdit(self.rule["range"])
        f.addRow("Apply to range", self.rng)
        self.kind = QComboBox()
        for k in V.KINDS:
            self.kind.addItem(V.KIND_LABEL[k], k)
        self.kind.setCurrentIndex(V.KINDS.index(self.rule["kind"]) if self.rule["kind"] in V.KINDS else 1)
        f.addRow("Allow", self.kind)
        self.op = QComboBox()
        for o in V.OPS:
            self.op.addItem(V.OP_LABEL[o], o)
        self.op.setCurrentIndex(V.OPS.index(self.rule.get("op", "between")))
        self.lbl_op = QLabel("Data")
        f.addRow(self.lbl_op, self.op)
        self.a = QLineEdit(str(self.rule.get("a", "")))
        self.b = QLineEdit(str(self.rule.get("b", "")))
        self.lbl_a, self.lbl_b = QLabel("Minimum / value"), QLabel("Maximum")
        f.addRow(self.lbl_a, self.a)
        f.addRow(self.lbl_b, self.b)
        self.src = QLineEdit(str(self.rule.get("list", "")))
        self.src.setPlaceholderText("Yes,No,Maybe   or   =A1:A10   or   =Lists!$A$1:$A$5")
        self.lbl_src = QLabel("Source")
        f.addRow(self.lbl_src, self.src)
        self.blank = QCheckBox("Ignore blank")
        self.blank.setChecked(self.rule.get("blank", True))
        self.drop = QCheckBox("In-cell dropdown (list)")
        self.drop.setChecked(self.rule.get("dropdown", True))
        f.addRow(self.blank)
        f.addRow(self.drop)
        tabs.addTab(t1, "Criteria")
        # --- input message
        t2 = QWidget()
        f2 = QFormLayout(t2)
        self.in_title = QLineEdit(self.rule.get("in_title", ""))
        self.in_msg = QPlainTextEdit(self.rule.get("in_msg", ""))
        f2.addRow("Title", self.in_title)
        f2.addRow("Message shown when the cell is selected", self.in_msg)
        tabs.addTab(t2, "Input message")
        # --- error alert
        t3 = QWidget()
        f3 = QFormLayout(t3)
        self.style = QComboBox()
        for k, t in (("stop", "Stop (value is refused)"), ("warning", "Warning (you may continue)"), ("information", "Information (you may continue)")):
            self.style.addItem(t, k)
        self.style.setCurrentIndex(["stop", "warning", "information"].index(self.rule.get("style", "stop")))
        self.title = QLineEdit(self.rule.get("title", ""))
        self.msg = QPlainTextEdit(self.rule.get("msg", ""))
        f3.addRow("Style", self.style)
        f3.addRow("Title", self.title)
        f3.addRow("Message", self.msg)
        tabs.addTab(t3, "Error alert")
        row = QHBoxLayout()
        clr = QPushButton("Remove validation")
        clr.clicked.connect(self.clear_all)
        row.addWidget(clr)
        row.addStretch()
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        row.addWidget(bb)
        lay.addLayout(row)
        self.kind.currentIndexChanged.connect(self.update_fields)
        self.op.currentIndexChanged.connect(self.update_fields)
        self.update_fields()

    def update_fields(self):
        k = self.kind.currentData()
        op = self.op.currentData()
        num = k in ("whole", "decimal", "date", "length")
        self.op.setVisible(num)
        self.lbl_op.setVisible(num)
        between = num and op in ("between", "notBetween")
        self.a.setVisible(num or k == "custom")
        self.lbl_a.setVisible(num or k == "custom")
        self.lbl_a.setText("Formula (true = allowed)" if k == "custom" else ("Minimum" if between else "Value"))
        self.b.setVisible(between)
        self.lbl_b.setVisible(between)
        self.src.setVisible(k == "list")
        self.lbl_src.setVisible(k == "list")
        self.drop.setVisible(k == "list")
        if k == "date":
            self.a.setPlaceholderText("2026-01-31")
            self.b.setPlaceholderText("2026-12-31")
        else:
            self.a.setPlaceholderText("")
            self.b.setPlaceholderText("")

    def clear_all(self):
        self.cleared = True
        self.accept()

    def result_rule(self):
        r = dict(self.rule)
        r.update(range=self.rng.text().strip(), kind=self.kind.currentData(), op=self.op.currentData(), a=self.a.text().strip(), b=self.b.text().strip(),
                 list=self.src.text().strip(), blank=self.blank.isChecked(), dropdown=self.drop.isChecked(), style=self.style.currentData(), title=self.title.text(),
                 msg=self.msg.toPlainText(), in_title=self.in_title.text(), in_msg=self.in_msg.toPlainText())
        return r
