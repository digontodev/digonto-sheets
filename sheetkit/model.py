"""Workbook / Sheet data, styles, typed input, recalculation cache, undo, row & column editing."""
import re
from collections import namedtuple

from . import formula as F
from .numfmt import format_value, general, is_date_code

Style = namedtuple("Style", "bold italic underline strike color bg font size h v wrap fmt bl bt br bb indent", defaults=(
    False, False, False, False, None, None, None, None, None, None, False, "General", None, None, None, None, 0))
DEFAULT = Style()
DEFAULT_W, DEFAULT_H = 100, 24
MAX_SHEET_NAME = 31

CF_KINDS = ["greater than", "less than", "equal to", "not equal to", "between", "text contains", "is empty", "is not empty", "duplicate values", "top 10", "bottom 10", "above average", "below average"]


class Sheet:
    def __init__(self, name):
        self.name = name
        self.cells = {}          # (r, c) -> (raw, style_id)
        self.colw = {}           # c -> px
        self.rowh = {}           # r -> px
        self.hidden_cols = set()
        self.hidden_rows = set()
        self.merges = []         # (r1, c1, r2, c2)
        self.freeze = (0, 0)     # rows, cols
        self.cf = []             # conditional formats: dict(kind, ranges, a, b, fg, bg)
        self.tab_color = None
        self.grid = True
        self.filter = None       # dict(range=(r1,c1,r2,c2), hidden={col: set(values)})
        self.dv = []             # data validation: dict(range, kind, items/min/max)
        self.charts = []         # dict(kind, range, title, row, col)
        self.comments = {}       # (r,c) -> text

    def copy_props(self):
        s = Sheet(self.name)
        s.cells = dict(self.cells)
        s.colw, s.rowh = dict(self.colw), dict(self.rowh)
        s.hidden_cols, s.hidden_rows = set(self.hidden_cols), set(self.hidden_rows)
        s.merges = list(self.merges)
        s.freeze = self.freeze
        s.cf = [dict(x) for x in self.cf]
        s.tab_color, s.grid = self.tab_color, self.grid
        s.filter = dict(self.filter) if self.filter else None
        s.dv = [dict(x) for x in self.dv]
        s.charts = [dict(x) for x in self.charts]
        s.comments = dict(self.comments)
        return s


NUM_RE = re.compile(r"^\s*([-+]?)\s*(\$|€|£|¥)?\s*([-+]?(?:\d{1,3}(?:,\d{3})+|\d+)?(?:\.\d+)?(?:[eE][-+]?\d+)?)\s*(%)?\s*$")


def parse_input(text):
    """what the user typed -> (raw value, number format or None)"""
    if text is None or text == "":
        return None, None
    if text.startswith("="):
        return (text, None) if len(text) > 1 else (text, None)
    if text.startswith("'"):
        return text, None
    s = text.strip()
    u = s.upper()
    if u == "TRUE":
        return True, None
    if u == "FALSE":
        return False, None
    m = NUM_RE.match(s)
    if m and m.group(3) and re.search(r"\d", m.group(3)):
        try:
            v = float(m.group(3).replace(",", ""))
            if m.group(1) == "-":
                v = -v
            fmt = None
            if m.group(2):
                fmt = m.group(2) + "#,##0.00"
            if m.group(4):
                v /= 100
                dec = len(m.group(3).partition(".")[2]) if "." in m.group(3) else 0
                fmt = "0" + ("." + "0" * dec if dec else "") + "%"
            if abs(v) < 1e15 and v == int(v):
                v = int(v) if float(int(v)) == v and "." not in m.group(3) and "e" not in m.group(3).lower() else v
            return v, fmt
        except ValueError:
            pass
    if re.match(r"^\d{1,4}[-/\.]\d{1,2}[-/\.]\d{1,4}", s) or re.match(r"^\d{1,2}:\d{2}", s) or re.match(r"^\d{1,2}[- ][A-Za-z]{3}[- ]\d{4}$", s):
        d = F.parse_date(s.replace(".", "/"))
        if d is not None:
            if ":" in s and "-" not in s and "/" not in s:
                return d, "hh:mm:ss" if s.count(":") == 2 else "hh:mm"
            return (d, "yyyy-mm-dd hh:mm") if (d != int(d)) else (d, "yyyy-mm-dd" if re.match(r"^\d{4}", s) else "m/d/yyyy")
    return text, None


class Workbook:
    def __init__(self):
        self.sheets = [Sheet("Sheet1")]
        self.styles = [DEFAULT]
        self._sid = {DEFAULT: 0}
        self.names = {}          # defined names: NAME -> "Sheet1!$A$1:$B$2"
        self.version = 0
        self._cache = {}
        self._stack = set()
        self._extent = {}
        self.undo_stack, self.redo_stack = [], []
        self.listeners = []
        self.path = None
        self.dirty = False
        self.date1904 = False
        self.calc_manual = False

    # ---------------------------------------------------------------- sheets
    def sheet(self, name):
        for s in self.sheets:
            if s.name.lower() == name.lower():
                return s
        return None

    def has_sheet(self, name):
        return self.sheet(name) is not None

    def sheet_names(self):
        return [s.name for s in self.sheets]

    def unique_name(self, base="Sheet"):
        n = 1
        while self.sheet(f"{base}{n}"):
            n += 1
        return f"{base}{n}"

    # ---------------------------------------------------------------- styles
    def style_id(self, st):
        sid = self._sid.get(st)
        if sid is None:
            sid = len(self.styles)
            self.styles.append(st)
            self._sid[st] = sid
        return sid

    def style_of(self, sheet, r, c):
        cell = self.sheet(sheet).cells.get((r, c)) if isinstance(sheet, str) else sheet.cells.get((r, c))
        return self.styles[cell[1]] if cell else DEFAULT

    # ---------------------------------------------------------------- values
    def touch(self, info=None):
        self.version += 1
        self._cache.clear()
        self._extent.clear()
        self.dirty = True
        for fn in self.listeners:
            fn(info)

    def used_extent(self, sheet):
        e = self._extent.get(sheet)
        if e is None:
            s = self.sheet(sheet)
            if s is None or not s.cells:
                e = (-1, -1)
            else:
                e = (max(r for r, _ in s.cells), max(c for _, c in s.cells))
            self._extent[sheet] = e
        return e

    def raw(self, sheet, r, c):
        s = self.sheet(sheet) if isinstance(sheet, str) else sheet
        cell = s.cells.get((r, c)) if s else None
        return cell[0] if cell else None

    def cell_value(self, sheet, r, c):
        key = (sheet, r, c)
        v = self._cache.get(key, self)
        if v is not self:
            return v
        s = self.sheet(sheet)
        if s is None:
            return F.REF
        cell = s.cells.get((r, c))
        if cell is None:
            return None
        raw = cell[0]
        if isinstance(raw, str):
            if raw.startswith("=") and len(raw) > 1:
                if key in self._stack:
                    return F.CIRC
                self._stack.add(key)
                try:
                    try:
                        v = F.evaluate(raw, F.Ctx(self, s.name, r, c))
                    except RecursionError:
                        if len(self._stack) > 1:
                            raise
                        v = self._deep_eval(key)
                finally:
                    self._stack.discard(key)
            elif raw.startswith("'"):
                v = raw[1:]
            else:
                v = raw if raw != "" else None
        else:
            v = raw
        self._cache[key] = v
        return v

    def _deep_eval(self, key):
        """very long chains of formulas: evaluate precedents first, from the bottom of the chain up"""
        order, seen, stack = [], set(), [key]
        while stack:
            k = stack.pop()
            if k in seen:
                continue
            seen.add(k)
            order.append(k)
            sh, r, c = k
            s = self.sheet(sh)
            cell = s.cells.get((r, c)) if s else None
            if cell and isinstance(cell[0], str) and cell[0].startswith("="):
                for (psh, r1, c1, r2, c2) in F.precedents(cell[0], sh):
                    ps = self.sheet(psh)
                    if ps is None or (r2 - r1 + 1) * (c2 - c1 + 1) > 5000:
                        continue
                    for rr in range(r1, r2 + 1):
                        for cc in range(c1, c2 + 1):
                            if (rr, cc) in ps.cells:
                                stack.append((psh, rr, cc))
        for k in reversed(order[1:]):
            if k not in self._cache and k not in self._stack:
                try:
                    self.cell_value(*k)
                except RecursionError:
                    pass
        return F.evaluate(self.sheet(key[0]).cells[(key[1], key[2])][0], F.Ctx(self, key[0], key[1], key[2]))

    def named_value(self, name, ctx):
        t = self.names.get(name.upper())
        if t is None:
            raise F.NAME
        m = F.TOKEN_RE.fullmatch(t)
        if not m or m.group("ref") is None:
            raise F.NAME
        r = F.Parser("", ctx.sheet).ref_node(m)(ctx)
        if isinstance(r, F.XErr):
            raise r
        return r

    def display(self, sheet, r, c):
        """(text, value, style) for a cell, as shown on screen"""
        s = self.sheet(sheet) if isinstance(sheet, str) else sheet
        cell = s.cells.get((r, c))
        if cell is None:
            return "", None, DEFAULT
        st = self.styles[cell[1]]
        v = self.cell_value(s.name, r, c)
        if isinstance(v, F.XErr):
            return v.code, v, st
        if v is None:
            return "", v, st
        if isinstance(v, bool):
            return ("TRUE" if v else "FALSE"), v, st
        if isinstance(v, (int, float)):
            return format_value(v, st.fmt), v, st
        return (format_value(v, st.fmt) if st.fmt not in ("General", "@") and "@" in st.fmt else str(v)), v, st

    def edit_text(self, sheet, r, c):
        """text shown in the formula bar / editor"""
        s = self.sheet(sheet)
        cell = s.cells.get((r, c))
        if cell is None:
            return ""
        raw = cell[0]
        if raw is None:
            return ""
        if isinstance(raw, bool):
            return "TRUE" if raw else "FALSE"
        if isinstance(raw, (int, float)):
            fmt = self.styles[cell[1]].fmt
            if fmt != "General" and is_date_code(fmt):
                return format_value(raw, "yyyy-mm-dd hh:mm:ss" if raw != int(raw) else "yyyy-mm-dd")
            if "%" in fmt and "E" not in fmt:
                return ("%.12g" % (raw * 100)) + "%"
            return ("%.15g" % raw) if not (isinstance(raw, int)) else str(raw)
        return raw

    # ---------------------------------------------------------------- changes and undo
    def _push(self, label, undo, redo):
        self.undo_stack.append((label, undo, redo))
        if len(self.undo_stack) > 200:
            self.undo_stack.pop(0)
        self.redo_stack.clear()

    def apply(self, changes, label="Edit"):
        """changes: {(sheetname, r, c): (raw, sid) or None}"""
        old = {}
        for (sh, r, c), new in changes.items():
            s = self.sheet(sh)
            old[(sh, r, c)] = s.cells.get((r, c))
        new_all = dict(changes)

        def run(d):
            for (sh, r, c), cell in d.items():
                s = self.sheet(sh)
                if s is None:
                    continue
                if cell is None or (cell[0] is None and cell[1] == 0):
                    s.cells.pop((r, c), None)
                else:
                    s.cells[(r, c)] = cell
            self.touch(("cells", list(d)))
        run(new_all)
        self._push(label, lambda: run(old), lambda: run(new_all))

    def set_values(self, sheet, items, label="Edit"):
        """items: [(r, c, text)] typed input"""
        ch = {}
        s = self.sheet(sheet)
        for r, c, text in items:
            raw, fmt = parse_input(text)
            cur = s.cells.get((r, c))
            st = self.styles[cur[1]] if cur else DEFAULT
            if fmt and st.fmt == "General":
                st = st._replace(fmt=fmt)
            elif raw is not None and isinstance(raw, (int, float)) and not isinstance(raw, bool) and st.fmt == "@":
                raw = text
            ch[(sheet, r, c)] = (raw, self.style_id(st)) if (raw is not None or st != DEFAULT) else None
        self.apply(ch, label)

    def set_style(self, sheet, cells, fn, label="Format"):
        """apply fn(Style) -> Style to every cell in cells [(r,c)]"""
        ch = {}
        s = self.sheet(sheet)
        for (r, c) in cells:
            cur = s.cells.get((r, c))
            raw = cur[0] if cur else None
            st = self.styles[cur[1]] if cur else DEFAULT
            ns = fn(st)
            if ns != st:
                ch[(sheet, r, c)] = (raw, self.style_id(ns))
        if ch:
            self.apply(ch, label)

    def undo(self):
        if not self.undo_stack:
            return False
        label, undo, redo = self.undo_stack.pop()
        undo()
        self.redo_stack.append((label, undo, redo))
        return True

    def redo(self):
        if not self.redo_stack:
            return False
        label, undo, redo = self.redo_stack.pop()
        redo()
        self.undo_stack.append((label, undo, redo))
        return True

    def snapshot(self):
        return ([s.copy_props() for s in self.sheets], dict(self.names))

    def restore(self, snap):
        sheets, names = snap
        self.sheets = [s.copy_props() for s in sheets]
        self.names = dict(names)
        self.touch(("all",))

    def structural(self, label, op):
        before = self.snapshot()
        op()
        after = self.snapshot()
        self.touch(("all",))
        self._push(label, lambda: self.restore(before), lambda: self.restore(after))

    # ---------------------------------------------------------------- rows / columns
    def insert(self, sheet, axis, index, count=1):
        self.structural("Insert " + axis + "s", lambda: self._edit_axis(sheet, axis, index, count))

    def delete(self, sheet, axis, index, count=1):
        self.structural("Delete " + axis + "s", lambda: self._edit_axis(sheet, axis, index, -count))

    def _edit_axis(self, sheet, axis, index, count):
        s = self.sheet(sheet)
        ax = 0 if axis == "row" else 1
        # formulas everywhere
        for sh in self.sheets:
            for k, (raw, sid) in list(sh.cells.items()):
                if isinstance(raw, str) and raw.startswith("=") and sheet.lower() in raw.lower() or (sh is s and isinstance(raw, str) and raw.startswith("=")):
                    nr = F.adjust_for_edit(raw, sh.name, s.name, axis, index, count)
                    if nr != raw:
                        sh.cells[k] = (nr, sid)
        for nm, t in list(self.names.items()):
            self.names[nm] = F.adjust_for_edit("=" + t, s.name, s.name, axis, index, count)[1:]

        def mv(i):
            if count > 0:
                return i + count if i >= index else i
            n = -count
            if i < index:
                return i
            if i < index + n:
                return None
            return i - n
        new = {}
        for (r, c), v in s.cells.items():
            k = (r, c)
            ni = mv(k[ax])
            if ni is None:
                continue
            nk = (ni, c) if ax == 0 else (r, ni)
            new[nk] = v
        s.cells = new
        for attr in (("rowh" if ax == 0 else "colw"),):
            d = getattr(s, attr)
            nd = {}
            for i, v in d.items():
                ni = mv(i)
                if ni is not None:
                    nd[ni] = v
            setattr(s, attr, nd)
        hid = "hidden_rows" if ax == 0 else "hidden_cols"
        setattr(s, hid, {mv(i) for i in getattr(s, hid) if mv(i) is not None})
        merges = []
        for (r1, c1, r2, c2) in s.merges:
            a1, a2 = mv((r1, c1)[ax]), mv((r2, c2)[ax])
            if count < 0 and a1 is None and a2 is None:
                continue
            lo, hi = (r1, r2) if ax == 0 else (c1, c2)
            if count > 0:
                nlo, nhi = mv(lo), mv(hi)
                if lo < index <= hi:
                    nhi = hi + count
            else:
                n = -count
                nlo = lo if lo < index else max(index, lo - n)
                nhi = hi if hi < index else hi - min(n, hi - index + 1)
                if nhi < nlo:
                    continue
            m = (nlo, c1, nhi, c2) if ax == 0 else (r1, nlo, r2, nhi)
            if m[0] != m[2] or m[1] != m[3]:
                merges.append(m)
        s.merges = merges
        fr, fc = s.freeze
        s.freeze = (fr, fc)

    # ---------------------------------------------------------------- sheets
    def add_sheet(self, name=None, index=None):
        name = name or self.unique_name()

        def op():
            sh = Sheet(name)
            self.sheets.insert(len(self.sheets) if index is None else index, sh)
        self.structural("Add sheet", op)
        return name

    def delete_sheet(self, name):
        if len(self.sheets) < 2:
            return

        def op():
            self.sheets = [s for s in self.sheets if s.name != name]
        self.structural("Delete sheet", op)

    def rename_sheet(self, old, new):
        def op():
            for sh in self.sheets:
                for k, (raw, sid) in list(sh.cells.items()):
                    if isinstance(raw, str) and raw.startswith("=") and "!" in raw:
                        nr = F.rename_sheet_refs(raw, old, new)
                        if nr != raw:
                            sh.cells[k] = (nr, sid)
            self.sheet(old).name = new
            for nm, t in list(self.names.items()):
                self.names[nm] = F.rename_sheet_refs("=" + t, old, new)[1:]
        self.structural("Rename sheet", op)

    def duplicate_sheet(self, name):
        def op():
            src = self.sheet(name)
            cp = src.copy_props()
            n = 2
            while self.sheet(f"{name} ({n})"):
                n += 1
            cp.name = f"{name} ({n})"
            self.sheets.insert(self.sheets.index(src) + 1, cp)
        self.structural("Duplicate sheet", op)

    def move_sheet(self, name, new_index):
        def op():
            s = self.sheet(name)
            self.sheets.remove(s)
            self.sheets.insert(new_index, s)
        self.structural("Move sheet", op)

    def set_prop(self, sheet, attr, value, label="Change"):
        """simple sheet property with undo (column widths, merges, freeze ...)"""
        s = self.sheet(sheet)
        old = getattr(s, attr)
        import copy
        newv = copy.deepcopy(value)
        oldv = copy.deepcopy(old)

        def run(v):
            setattr(self.sheet(sheet), attr, copy.deepcopy(v))
            self.touch(("props", sheet))
        run(newv)
        self._push(label, lambda: run(oldv), lambda: run(newv))
