"""Formula engine: tokenizer, parser (compiled to closures) and the function library."""
import datetime
import math
import random
import re
import statistics

from .numfmt import EPOCH, dt_to_serial, format_value, round_half_up, serial_to_dt


class XErr(Exception):
    """an Excel error value (#DIV/0! ...). Used as a value and as an exception"""

    def __init__(self, code="#VALUE!"):
        super().__init__(code)
        self.code = code

    def __repr__(self):
        return self.code

    __str__ = __repr__

    def __eq__(self, o):
        return isinstance(o, XErr) and o.code == self.code

    def __hash__(self):
        return hash(self.code)


DIV0, VALUE, REF, NAME, NA, NUM, NULL, CIRC = (XErr(c) for c in ("#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#N/A", "#NUM!", "#NULL!", "#CIRC!"))
ERRORS = {e.code: e for e in (DIV0, VALUE, REF, NAME, NA, NUM, NULL, CIRC)}

MAXROW, MAXCOL = 1048576, 16384


# ---------------------------------------------------------------- cell names
def col_name(c):  # 0 based -> A
    s = ""
    c += 1
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def col_index(s):
    n = 0
    for ch in s.upper():
        n = n * 26 + ord(ch) - 64
    return n - 1


def cell_name(r, c, absr=False, absc=False):
    return ("$" if absc else "") + col_name(c) + ("$" if absr else "") + str(r + 1)


CELL = r"\$?[A-Za-z]{1,3}\$?\d+"
SHEET = r"(?:'(?:[^']|'')+'|[A-Za-z_À-￿][\w\.À-￿]*)"
TOKEN_RE = re.compile(
    r"""(?P<ws>\s+)
    |(?P<str>"(?:[^"]|"")*")
    |(?P<ref>(?:(?P<sh>""" + SHEET + r""")!)?(?:(?P<a>""" + CELL + r""")(?::(?P<b>""" + CELL + r"""))?|(?P<c1>\$?[A-Za-z]{1,3}):(?P<c2>\$?[A-Za-z]{1,3})|(?P<r1>\$?\d+):(?P<r2>\$?\d+)))(?![\w\(\.!])
    |(?P<num>(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)
    |(?P<err>\#(?:DIV/0!|VALUE!|REF!|NAME\?|N/A|NUM!|NULL!|CIRC!))
    |(?P<func>[A-Za-z_][\w\.]*)(?=\()
    |(?P<name>[A-Za-z_][\w\.]*)
    |(?P<op><>|<=|>=|[-+*/^&=<>%])
    |(?P<lp>\()|(?P<rp>\))|(?P<sep>[,;])|(?P<colon>:)|(?P<brace>[{}])
    """, re.X)


def tokenize(text):
    pos, out = 0, []
    while pos < len(text):
        m = TOKEN_RE.match(text, pos)
        if not m:
            raise XErr("#NAME?")
        kind = m.lastgroup
        for k in ("ref", "str", "num", "err", "func", "name", "op", "lp", "rp", "sep", "colon", "brace", "ws"):
            if m.group(k) is not None:
                kind = k
                break
        if kind != "ws":
            out.append((kind, m.group(0), m.start(), m.end(), m))
        pos = m.end()
    return out


def unquote_sheet(s):
    if s and s.startswith("'"):
        return s[1:-1].replace("''", "'")
    return s


def quote_sheet(s):
    if re.match(r"^[A-Za-z_][\w\.]*$", s) and not re.match(r"^" + CELL + "$", s):
        return s
    return "'" + s.replace("'", "''") + "'"


def parse_cell(s):
    m = re.match(r"^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$", s)
    return int(m.group(4)) - 1, col_index(m.group(2)), bool(m.group(3)), bool(m.group(1))


class Ref:
    """a reference to one cell or a rectangle on a sheet"""
    __slots__ = ("sheet", "r1", "c1", "r2", "c2")

    def __init__(self, sheet, r1, c1, r2=None, c2=None):
        self.sheet = sheet
        self.r1, self.c1 = min(r1, r2 if r2 is not None else r1), min(c1, c2 if c2 is not None else c1)
        self.r2, self.c2 = max(r1, r2 if r2 is not None else r1), max(c1, c2 if c2 is not None else c1)

    @property
    def single(self):
        return self.r1 == self.r2 and self.c1 == self.c2


class Ctx:
    """what a formula can see: the workbook and the cell it lives in"""

    def __init__(self, wb, sheet, row=0, col=0):
        self.wb, self.sheet, self.row, self.col = wb, sheet, row, col


# ---------------------------------------------------------------- parser -> closures
class Parser:
    def __init__(self, text, sheet):
        self.toks = tokenize(text)
        self.i = 0
        self.sheet = sheet

    def peek(self):
        return self.toks[self.i] if self.i < len(self.toks) else (None, None, 0, 0, None)

    def take(self):
        t = self.peek()
        self.i += 1
        return t

    def parse(self):
        node = self.expr(0)
        if self.i < len(self.toks):
            raise XErr("#NAME?")
        return node

    LEVELS = [("=", "<>", "<", ">", "<=", ">="), ("&",), ("+", "-"), ("*", "/"), ("^",)]

    def expr(self, lvl):
        if lvl >= len(self.LEVELS):
            return self.unary()
        left = self.expr(lvl + 1)
        while True:
            k, t, *_ = self.peek()
            if k == "op" and t in self.LEVELS[lvl]:
                self.take()
                right = self.expr(lvl + 1)
                left = bin_node(t, left, right)
            else:
                return left

    def unary(self):
        k, t, *_ = self.peek()
        if k == "op" and t in "+-":
            self.take()
            x = self.unary()
            if t == "-":
                return lambda ctx, x=x: neg(x(ctx))
            return x
        return self.postfix()

    def postfix(self):
        node = self.primary()
        while self.peek()[0] == "op" and self.peek()[1] == "%":
            self.take()
            node = (lambda ctx, n=node: elementwise(lambda a: num(a) / 100, n(ctx)))
        return node

    def primary(self):
        k, t, a, b, m = self.take()
        if k == "num":
            v = float(t)
            return lambda ctx, v=v: v
        if k == "str":
            s = t[1:-1].replace('""', '"')
            return lambda ctx, s=s: s
        if k == "err":
            e = ERRORS[t]
            return lambda ctx, e=e: e
        if k == "lp":
            n = self.expr(0)
            if self.take()[0] != "rp":
                raise XErr("#NAME?")
            return n
        if k == "ref":
            return self.ref_node(m)
        if k == "func":
            return self.call(t)
        if k == "name":
            u = t.upper()
            if u == "TRUE":
                return lambda ctx: True
            if u == "FALSE":
                return lambda ctx: False
            return lambda ctx, t=t: ctx.wb.named_value(t, ctx)
        if k == "brace":
            return self.array()
        raise XErr("#NAME?")

    def array(self):
        rows, row = [], []
        while True:
            k, t, *_ = self.take()
            if k is None:
                raise XErr("#NAME?")
            if k == "brace":
                break
            if k == "sep":
                if t == ";":
                    rows.append(row)
                    row = []
                continue
            neg_ = False
            if k == "op" and t == "-":
                neg_ = True
                k, t, *_ = self.take()
            if k == "num":
                row.append(-float(t) if neg_ else float(t))
            elif k == "str":
                row.append(t[1:-1].replace('""', '"'))
            elif k == "name":
                row.append(t.upper() == "TRUE")
            else:
                raise XErr("#NAME?")
        rows.append(row)
        return lambda ctx, rows=rows: [r[:] for r in rows]

    def ref_node(self, m):
        sh = unquote_sheet(m.group("sh")) if m.group("sh") else None
        sheet = sh or self.sheet
        if m.group("a"):
            r1, c1, *_ = parse_cell(m.group("a"))
            if m.group("b"):
                r2, c2, *_ = parse_cell(m.group("b"))
            else:
                r2, c2 = r1, c1
        elif m.group("c1"):
            c1, c2 = col_index(m.group("c1").lstrip("$")), col_index(m.group("c2").lstrip("$"))
            r1, r2 = 0, MAXROW - 1
        else:
            r1, r2 = int(m.group("r1").lstrip("$")) - 1, int(m.group("r2").lstrip("$")) - 1
            c1, c2 = 0, MAXCOL - 1
        if sh is not None:
            return lambda ctx, sh=sh, a=(r1, c1, r2, c2): Ref(sh if ctx.wb.has_sheet(sh) else None, *a) if ctx.wb.has_sheet(sh) else REF
        ref = (r1, c1, r2, c2)
        return lambda ctx, sheet=sheet, ref=ref: Ref(sheet, *ref)

    def call(self, name):
        self.take()  # (
        args = []
        if self.peek()[0] == "rp":
            self.take()
        else:
            while True:
                if self.peek()[0] == "sep":
                    args.append(lambda ctx: None)  # empty argument
                else:
                    args.append(self.expr(0))
                k, *_ = self.take()
                if k == "rp":
                    break
                if k != "sep":
                    raise XErr("#NAME?")
                if self.peek()[0] == "rp":
                    args.append(lambda ctx: None)
                    self.take()
                    break
        up = name.upper()
        if up.startswith("_XLFN."):
            up = up[6:]
        f = FUNCS.get(up)
        if f is None:
            return lambda ctx, n=name: NAME
        fn, lazy = f
        if lazy:
            return lambda ctx, fn=fn, args=args: fn(ctx, *args)
        return lambda ctx, fn=fn, args=args: fn(ctx, *[a(ctx) for a in args])


def bin_node(op, a, b):
    f = BINOPS[op]
    return lambda ctx: binary(f, a(ctx), b(ctx))


def binary(f, x, y):
    x, y = deref(x), deref(y)
    if isinstance(x, XErr):
        raise x
    if isinstance(y, XErr):
        raise y
    if isinstance(x, list) or isinstance(y, list):
        xm, ym = as_matrix(x), as_matrix(y)
        R, C = max(len(xm), len(ym)), max(len(xm[0]), len(ym[0]))

        def at(m, i, j):
            if len(m) == 1:
                i = 0
            if len(m[0]) == 1:
                j = 0
            if i >= len(m) or j >= len(m[0]):
                return NA
            return m[i][j]

        out = []
        for i in range(R):
            row = []
            for j in range(C):
                p, q = at(xm, i, j), at(ym, i, j)
                try:
                    if isinstance(p, XErr):
                        row.append(p)
                    elif isinstance(q, XErr):
                        row.append(q)
                    else:
                        row.append(f(p, q))
                except XErr as e:
                    row.append(e)
            out.append(row)
        return out
    return f(x, y)


def elementwise(f, x):
    x = deref(x)
    if isinstance(x, XErr):
        raise x
    if isinstance(x, list):
        out = []
        for row in x:
            r = []
            for v in row:
                try:
                    r.append(v if isinstance(v, XErr) else f(v))
                except XErr as e:
                    r.append(e)
            out.append(r)
        return out
    return f(x)


def neg(x):
    return elementwise(lambda v: -num(v), x)


def num(v):
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, (int, float)):
        return v
    if v is None:
        return 0.0
    if isinstance(v, XErr):
        raise v
    if isinstance(v, str):
        s = v.strip().replace(",", "") if re.match(r"^\s*[-+]?[\d,]*\.?\d+([eE][-+]?\d+)?%?\s*$", v) else v
        try:
            if s.endswith("%"):
                return float(s[:-1]) / 100
            return float(s)
        except ValueError:
            pass
        d = parse_date(v)
        if d is not None:
            return d
        raise VALUE
    raise VALUE


def text(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (int, float)):
        return format_value(v, "General")
    if isinstance(v, XErr):
        raise v
    return str(v)


def boolean(v):
    if isinstance(v, bool):
        return v
    if v is None:
        return False
    if isinstance(v, (int, float)):
        return v != 0
    if isinstance(v, str):
        u = v.strip().upper()
        if u == "TRUE":
            return True
        if u == "FALSE":
            return False
        raise VALUE
    if isinstance(v, XErr):
        raise v
    raise VALUE


def parse_date(s):
    s = s.strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%d-%b-%Y", "%d %b %Y", "%b %d, %Y", "%B %d, %Y", "%Y-%m-%d %H:%M", "%Y-%m-%d %H:%M:%S", "%H:%M", "%H:%M:%S"):
        try:
            d = datetime.datetime.strptime(s, fmt)
            if fmt.startswith("%H"):
                return (d.hour * 3600 + d.minute * 60 + d.second) / 86400
            return dt_to_serial(d)
        except ValueError:
            continue
    return None


def deref(x):
    """Ref -> value (single cell) or matrix; callers that need the Ref itself don't call this"""
    if isinstance(x, Ref):
        if x.single:
            return CTX[-1].wb.cell_value(x.sheet, x.r1, x.c1)
        return matrix(x)
    return x


CTX = []  # stack of active contexts (so deref knows the workbook)


def matrix(ref):
    wb = CTX[-1].wb
    sh = ref.sheet
    if sh is None:
        raise REF
    r2, c2 = ref.r2, ref.c2
    if r2 >= MAXROW - 1 or c2 >= MAXCOL - 1:
        ur, uc = wb.used_extent(sh)
        r2, c2 = min(r2, max(ur, ref.r1)), min(c2, max(uc, ref.c1))
    if (r2 - ref.r1 + 1) * (c2 - ref.c1 + 1) > 4_000_000:
        raise NUM
    return [[wb.cell_value(sh, r, c) for c in range(ref.c1, c2 + 1)] for r in range(ref.r1, r2 + 1)]


def as_matrix(x):
    x = deref(x) if isinstance(x, Ref) else x
    if isinstance(x, list):
        return x
    return [[x]]


def scalar(x):
    x = deref(x)
    if isinstance(x, list):
        x = x[0][0]
    if isinstance(x, XErr):
        raise x
    return x


def compare(a, b):
    """-1, 0, 1 with Excel ordering: numbers < text < logical; empty acts as 0 or ''"""
    if a is None:
        a = "" if isinstance(b, str) else (False if isinstance(b, bool) else 0.0)
    if b is None:
        b = "" if isinstance(a, str) else (False if isinstance(a, bool) else 0.0)

    def rank(v):
        return 2 if isinstance(v, bool) else 1 if isinstance(v, str) else 0

    ra, rb = rank(a), rank(b)
    if ra != rb:
        return -1 if ra < rb else 1
    if ra == 1:
        a, b = a.lower(), b.lower()
    return (a > b) - (a < b)


def _div(a, b):
    b = num(b)
    if b == 0:
        raise DIV0
    return num(a) / b


def _pow(a, b):
    a, b = num(a), num(b)
    try:
        r = a ** b
    except ZeroDivisionError:
        raise DIV0
    if isinstance(r, complex):
        raise NUM
    return r


BINOPS = {
    "+": lambda a, b: num(a) + num(b), "-": lambda a, b: num(a) - num(b), "*": lambda a, b: num(a) * num(b),
    "/": _div, "^": _pow, "&": lambda a, b: text(a) + text(b),
    "=": lambda a, b: compare(a, b) == 0, "<>": lambda a, b: compare(a, b) != 0, "<": lambda a, b: compare(a, b) < 0,
    ">": lambda a, b: compare(a, b) > 0, "<=": lambda a, b: compare(a, b) <= 0, ">=": lambda a, b: compare(a, b) >= 0,
}

_COMPILED = {}


def compile_formula(src, sheet):
    """'=A1+1' -> closure(ctx)"""
    key = (src, sheet)
    f = _COMPILED.get(key)
    if f is None:
        try:
            f = Parser(src[1:] if src.startswith("=") else src, sheet).parse()
        except XErr as e:
            f = lambda ctx, e=e: e
        except (IndexError, ValueError):
            f = lambda ctx: NAME
        if len(_COMPILED) > 50000:
            _COMPILED.clear()
        _COMPILED[key] = f
    return f


def evaluate(src, ctx):
    """evaluate a formula string in a context, return a plain value (or XErr)"""
    CTX.append(ctx)
    try:
        v = compile_formula(src, ctx.sheet)(ctx)
        if isinstance(v, Ref):
            v = deref(v)
        if isinstance(v, list):
            v = v[0][0] if v and v[0] else None
        if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
            return NUM
        return v
    except XErr as e:
        return e
    except ZeroDivisionError:
        return DIV0
    except (OverflowError, ValueError):
        return NUM
    except RecursionError:
        raise
    finally:
        CTX.pop()


def precedents(src, sheet):
    """list of (sheet, r1, c1, r2, c2) used by a formula"""
    out = []
    try:
        for k, t, a, b, m in tokenize(src):
            if k == "ref":
                p = Parser("", sheet)
                # compute bounds without creating closure
                sh = unquote_sheet(m.group("sh")) if m.group("sh") else sheet
                if m.group("a"):
                    r1, c1, *_ = parse_cell(m.group("a"))
                    r2, c2 = (parse_cell(m.group("b"))[:2] if m.group("b") else (r1, c1))
                elif m.group("c1"):
                    c1, c2, r1, r2 = col_index(m.group("c1").lstrip("$")), col_index(m.group("c2").lstrip("$")), 0, MAXROW - 1
                else:
                    r1, r2, c1, c2 = int(m.group("r1").lstrip("$")) - 1, int(m.group("r2").lstrip("$")) - 1, 0, MAXCOL - 1
                out.append((sh, r1, c1, r2, c2))
    except XErr:
        pass
    return out


# ---------------------------------------------------------------- reference rewriting
def rewrite_refs(src, fn):
    """rebuild a formula, calling fn(sheet_or_None, r1, c1, r2, c2, flags, kind) -> new ref text for every reference"""
    try:
        toks = tokenize(src)
    except XErr:
        return src
    out, last = [], 0
    for k, t, a, b, m in toks:
        if k != "ref":
            continue
        out.append(src[last:a])
        last = b
        sh_raw = m.group("sh")
        pre = (sh_raw + "!") if sh_raw else ""
        sh = unquote_sheet(sh_raw) if sh_raw else None
        if m.group("a"):
            r1, c1, ar1, ac1 = parse_cell(m.group("a"))
            if m.group("b"):
                r2, c2, ar2, ac2 = parse_cell(m.group("b"))
            else:
                r2, c2, ar2, ac2 = r1, c1, ar1, ac1
            res = fn(sh, (r1, c1, ar1, ac1), (r2, c2, ar2, ac2), "cell", bool(m.group("b")))
            out.append(pre + res if res != "#REF!" else "#REF!")
        elif m.group("c1"):
            c1, c2 = m.group("c1"), m.group("c2")
            res = fn(sh, (0, col_index(c1.lstrip("$")), False, c1.startswith("$")), (MAXROW - 1, col_index(c2.lstrip("$")), False, c2.startswith("$")), "col", True)
            out.append(pre + res if res != "#REF!" else "#REF!")
        else:
            r1, r2 = m.group("r1"), m.group("r2")
            res = fn(sh, (int(r1.lstrip("$")) - 1, 0, r1.startswith("$"), False), (int(r2.lstrip("$")) - 1, MAXCOL - 1, r2.startswith("$"), False), "row", True)
            out.append(pre + res if res != "#REF!" else "#REF!")
    out.append(src[last:])
    return "".join(out)


def _fmt_ref(p1, p2, kind, is_range):
    (r1, c1, ar1, ac1), (r2, c2, ar2, ac2) = p1, p2
    if kind == "col":
        return ("$" if ac1 else "") + col_name(c1) + ":" + ("$" if ac2 else "") + col_name(c2)
    if kind == "row":
        return ("$" if ar1 else "") + str(r1 + 1) + ":" + ("$" if ar2 else "") + str(r2 + 1)
    a = cell_name(r1, c1, ar1, ac1)
    return a + ":" + cell_name(r2, c2, ar2, ac2) if is_range else a


def shift_formula(src, dr, dc):
    """copy/paste/fill: move relative references"""
    def fn(sh, p1, p2, kind, is_range):
        def mv(p):
            r, c, ar, ac = p
            if kind == "col":
                return (r, c if ac else c + dc, ar, ac)
            if kind == "row":
                return (r if ar else r + dr, c, ar, ac)
            return (r if ar else r + dr, c if ac else c + dc, ar, ac)
        q1, q2 = mv(p1), mv(p2)
        if min(q1[0], q2[0]) < 0 or min(q1[1], q2[1]) < 0:
            return "#REF!"
        return _fmt_ref(q1, q2, kind, is_range)
    return rewrite_refs(src, fn) if src.startswith("=") else src


def adjust_for_edit(src, home_sheet, target_sheet, axis, index, count):
    """rows/columns inserted (count>0) or deleted (count<0) in target_sheet at index: fix references in a formula that lives on home_sheet"""
    def fn(sh, p1, p2, kind, is_range):
        if (sh or home_sheet) != target_sheet:
            return _fmt_ref(p1, p2, kind, is_range)
        i = 0 if axis == "row" else 1
        if (kind == "col" and axis == "row") or (kind == "row" and axis == "col"):
            return _fmt_ref(p1, p2, kind, is_range)
        a, b = [p1[i], p2[i]], None
        lo, hi = min(a), max(a)
        if count > 0:
            lo2 = lo + count if lo >= index else lo
            hi2 = hi + count if hi >= index else hi
        else:
            n = -count
            end = index + n - 1
            if lo >= index and hi <= end:
                return "#REF!"
            lo2 = lo if lo < index else (index if lo <= end else lo - n)
            hi2 = hi if hi < index else (index - 1 if hi <= end else hi - n)
            if hi2 < lo2:
                return "#REF!"
        q1, q2 = list(p1), list(p2)
        if p1[i] <= p2[i]:
            q1[i], q2[i] = lo2, hi2
        else:
            q1[i], q2[i] = hi2, lo2
        return _fmt_ref(tuple(q1), tuple(q2), kind, is_range)
    return rewrite_refs(src, fn)


def rename_sheet_refs(src, old, new):
    try:
        toks = tokenize(src)
    except XErr:
        return src
    out, last = [], 0
    for k, t, a, b, m in toks:
        if k == "ref" and m.group("sh") and unquote_sheet(m.group("sh")) == old:
            sa, sb = m.start("sh"), m.end("sh")
            out.append(src[last:sa] + quote_sheet(new))
            last = sb
    out.append(src[last:])
    return "".join(out)


# ---------------------------------------------------------------- helpers for functions
def flat_numbers(args, strict=False):
    """numbers from arguments: cells in ranges only count if numeric; direct arguments are converted"""
    out = []
    for a in args:
        if isinstance(a, Ref):
            a = deref(a)
        if isinstance(a, list):
            for row in a:
                for v in row:
                    if isinstance(v, XErr):
                        raise v
                    if isinstance(v, bool) or v is None or isinstance(v, str):
                        continue
                    out.append(v)
        else:
            if isinstance(a, XErr):
                raise a
            if a is None:
                continue
            out.append(num(a))
    return out


def flat_all(args):
    out = []
    for a in args:
        if isinstance(a, Ref):
            a = deref(a)
        if isinstance(a, list):
            for row in a:
                out.extend(row)
        else:
            out.append(a)
    return out


def crit_pred(c):
    c = deref(c) if isinstance(c, Ref) else c
    if isinstance(c, list):
        c = c[0][0]
    if isinstance(c, XErr):
        raise c
    op = "="
    val = c
    if isinstance(c, str):
        m = re.match(r"^(<=|>=|<>|<|>|=)(.*)$", c, re.S)
        if m:
            op, val = m.group(1), m.group(2)
        else:
            val = c
        try:
            if val.strip() != "":
                val = float(val.replace(",", "")) if not val.upper() in ("TRUE", "FALSE") else val.upper() == "TRUE"
        except ValueError:
            pass
    if isinstance(val, str):
        if val == "" and op in ("=", "<>"):
            return (lambda v: v is None or v == "") if op == "=" else (lambda v: not (v is None or v == ""))
        pat = re.compile("^" + "".join(".*" if ch == "*" else "." if ch == "?" else re.escape(ch) for ch in re.sub(r"~(.)", r"\1", val)) + "$", re.I | re.S) if ("*" in val or "?" in val) else None

        def test(v):
            if op in ("=", "<>"):
                ok = (pat.match(v) is not None) if (pat and isinstance(v, str)) else (isinstance(v, str) and v.lower() == val.lower())
                return ok if op == "=" else not ok
            if not isinstance(v, str):
                return False
            cmpv = (v.lower() > val.lower()) - (v.lower() < val.lower())
            return {"<": cmpv < 0, ">": cmpv > 0, "<=": cmpv <= 0, ">=": cmpv >= 0}[op]
        return test

    def test(v):
        if isinstance(val, bool):
            if not isinstance(v, bool):
                return op == "<>"
            return (v == val) if op == "=" else (v != val) if op == "<>" else False
        if isinstance(v, bool) or not isinstance(v, (int, float)):
            return op == "<>"
        return {"=": v == val, "<>": v != val, "<": v < val, ">": v > val, "<=": v <= val, ">=": v >= val}[op]
    return test


def matching_indices(pairs):
    """pairs: [(range_arg, criteria)] -> list of (i, j) cells meeting all"""
    mats = [(as_matrix(r), crit_pred(c)) for r, c in pairs]
    R, C = len(mats[0][0]), len(mats[0][0][0])
    for m, _ in mats:
        if len(m) != R or len(m[0]) != C:
            raise VALUE
    return [(i, j) for i in range(R) for j in range(C) if all(p(m[i][j]) for m, p in mats)]


def vec(x):
    """argument -> flat list of values (row-major)"""
    return flat_all([x])


def split_pairs(args):
    if len(args) % 2:
        raise VALUE
    return [(args[i], args[i + 1]) for i in range(0, len(args), 2)]


def date_part(v):
    return serial_to_dt(num(v))


def to_serial(y, m, d):
    y, m = int(y), int(m)
    y += (m - 1) // 12
    m = (m - 1) % 12 + 1
    base = datetime.datetime(y, m, 1)
    return (base - EPOCH).days + int(d) - 1


def add_months(serial, n):
    d = serial_to_dt(serial)
    mm = d.month - 1 + int(n)
    y = d.year + mm // 12
    mm = mm % 12 + 1
    last = (datetime.datetime(y + (mm == 12), mm % 12 + 1, 1) - datetime.timedelta(days=1)).day
    return dt_to_serial(datetime.datetime(y, mm, min(d.day, last)))


# ---------------------------------------------------------------- the function library
FUNCS = {}
HELP = {}  # name -> (signature, description, category)


def fn(name, sig="", doc="", cat="Math", lazy=False):
    def deco(f):
        for n in name.split("|"):
            FUNCS[n] = (f, lazy)
            HELP[n] = (n + "(" + sig + ")", doc, cat)
        return f
    return deco


# --- math
@fn("SUM", "number1, [number2], ...", "Adds all the numbers.")
def f_sum(ctx, *a):
    return sum(flat_numbers(a))


@fn("PRODUCT", "number1, [number2], ...", "Multiplies the numbers.")
def f_product(ctx, *a):
    return math.prod(flat_numbers(a))


@fn("AVERAGE|AVG", "number1, [number2], ...", "Average of the numbers.", "Statistical")
def f_average(ctx, *a):
    n = flat_numbers(a)
    if not n:
        raise DIV0
    return sum(n) / len(n)


@fn("AVERAGEA", "value1, ...", "Average, text counts as 0 and TRUE as 1.", "Statistical")
def f_averagea(ctx, *a):
    vals = [0.0 if isinstance(v, str) else float(v) for v in flat_all(a) if v is not None]
    if not vals:
        raise DIV0
    return sum(vals) / len(vals)


@fn("MIN", "number1, [number2], ...", "Smallest number.", "Statistical")
def f_min(ctx, *a):
    n = flat_numbers(a)
    return min(n) if n else 0.0


@fn("MAX", "number1, [number2], ...", "Largest number.", "Statistical")
def f_max(ctx, *a):
    n = flat_numbers(a)
    return max(n) if n else 0.0


@fn("COUNT", "value1, ...", "How many cells hold numbers.", "Statistical")
def f_count(ctx, *a):
    n = 0
    for x in a:
        if isinstance(x, Ref):
            x = deref(x)
        if isinstance(x, list):
            n += sum(1 for v in flat_all([x]) if isinstance(v, (int, float)) and not isinstance(v, bool))
        else:
            try:
                num(x)
                n += 1
            except XErr:
                pass
    return float(n)


@fn("COUNTA", "value1, ...", "How many cells are not empty.", "Statistical")
def f_counta(ctx, *a):
    return float(sum(1 for v in flat_all(a) if v is not None and v != ""))


@fn("COUNTBLANK", "range", "How many cells are empty.", "Statistical")
def f_countblank(ctx, r):
    return float(sum(1 for v in vec(r) if v is None or v == ""))


@fn("ABS", "number", "Number without its sign.")
def f_abs(ctx, x):
    return elementwise(lambda v: abs(num(v)), x)


@fn("SQRT", "number", "Square root.")
def f_sqrt(ctx, x):
    def s(v):
        v = num(v)
        if v < 0:
            raise NUM
        return math.sqrt(v)
    return elementwise(s, x)


@fn("POWER", "number, power", "Number raised to a power.")
def f_power(ctx, a, b):
    return _pow(scalar(a), scalar(b))


@fn("MOD", "number, divisor", "Remainder after division.")
def f_mod(ctx, a, b):
    a, b = num(scalar(a)), num(scalar(b))
    if b == 0:
        raise DIV0
    return a - b * math.floor(a / b)


@fn("QUOTIENT", "numerator, denominator", "Integer part of a division.")
def f_quotient(ctx, a, b):
    b = num(scalar(b))
    if b == 0:
        raise DIV0
    return float(math.trunc(num(scalar(a)) / b))


@fn("INT", "number", "Rounds down to the nearest whole number.")
def f_int(ctx, x):
    return elementwise(lambda v: float(math.floor(num(v))), x)


@fn("TRUNC", "number, [digits]", "Cuts off the decimals.")
def f_trunc(ctx, x, d=0.0):
    d = int(num(scalar(d) if d is not None else 0))
    k = 10 ** d
    return elementwise(lambda v: math.trunc(num(v) * k) / k, x)


def _round_to(v, d, mode):
    v, d = num(v), int(num(d))
    k = 10.0 ** d
    if mode == "r":
        return round_half_up(v, d) if d >= 0 else math.copysign(math.floor(abs(v) * k + 0.5) / k, v)
    s = abs(v) * k
    s = math.ceil(s - 1e-9) if mode == "u" else math.floor(s + 1e-9)
    return math.copysign(s / k, v)


@fn("ROUND", "number, digits", "Rounds to a number of digits.")
def f_round(ctx, x, d=0.0):
    d = scalar(d)
    return elementwise(lambda v: _round_to(v, d, "r"), x)


@fn("ROUNDUP", "number, digits", "Rounds away from zero.")
def f_roundup(ctx, x, d=0.0):
    d = scalar(d)
    return elementwise(lambda v: _round_to(v, d, "u"), x)


@fn("ROUNDDOWN", "number, digits", "Rounds towards zero.")
def f_rounddown(ctx, x, d=0.0):
    d = scalar(d)
    return elementwise(lambda v: _round_to(v, d, "d"), x)


@fn("CEILING", "number, [significance]", "Rounds up to a multiple.")
def f_ceiling(ctx, x, s=1.0):
    s = num(scalar(s))
    if s == 0:
        return 0.0
    return math.ceil(num(scalar(x)) / s - 1e-12) * s


@fn("FLOOR", "number, [significance]", "Rounds down to a multiple.")
def f_floor(ctx, x, s=1.0):
    s = num(scalar(s))
    if s == 0:
        raise DIV0
    return math.floor(num(scalar(x)) / s + 1e-12) * s


@fn("MROUND", "number, multiple", "Rounds to the nearest multiple.")
def f_mround(ctx, x, m):
    m = num(scalar(m))
    return 0.0 if m == 0 else math.floor(num(scalar(x)) / m + 0.5) * m


@fn("EVEN", "number", "Rounds up to the nearest even integer.")
def f_even(ctx, x):
    v = num(scalar(x))
    r = math.ceil(abs(v) / 2) * 2
    return math.copysign(r, v)


@fn("ODD", "number", "Rounds up to the nearest odd integer.")
def f_odd(ctx, x):
    v = num(scalar(x))
    r = math.ceil(abs(v))
    r = r + 1 if r % 2 == 0 else r
    return math.copysign(r, v)


@fn("SIGN", "number", "1, 0 or -1.")
def f_sign(ctx, x):
    v = num(scalar(x))
    return float((v > 0) - (v < 0))


@fn("EXP", "number", "e to a power.")
def f_exp(ctx, x):
    return math.exp(num(scalar(x)))


@fn("LN", "number", "Natural logarithm.")
def f_ln(ctx, x):
    v = num(scalar(x))
    if v <= 0:
        raise NUM
    return math.log(v)


@fn("LOG", "number, [base]", "Logarithm to a base (default 10).")
def f_log(ctx, x, b=10.0):
    v, b = num(scalar(x)), num(scalar(b))
    if v <= 0 or b <= 0 or b == 1:
        raise NUM
    return math.log(v, b)


@fn("LOG10", "number", "Base 10 logarithm.")
def f_log10(ctx, x):
    v = num(scalar(x))
    if v <= 0:
        raise NUM
    return math.log10(v)


@fn("PI", "", "3.14159...")
def f_pi(ctx):
    return math.pi


for _n, _f in (("SIN", math.sin), ("COS", math.cos), ("TAN", math.tan), ("ASIN", math.asin), ("ACOS", math.acos), ("ATAN", math.atan),
               ("SINH", math.sinh), ("COSH", math.cosh), ("TANH", math.tanh), ("DEGREES", math.degrees), ("RADIANS", math.radians)):
    FUNCS[_n] = ((lambda g: (lambda ctx, x: g(num(scalar(x)))))(_f), False)
    HELP[_n] = (_n + "(number)", "Trigonometry (angles in radians)." if _n not in ("DEGREES", "RADIANS") else "Converts angles.", "Math")


@fn("ATAN2", "x, y", "Angle of a point.")
def f_atan2(ctx, x, y):
    return math.atan2(num(scalar(y)), num(scalar(x)))


@fn("RAND", "", "Random number from 0 up to 1.")
def f_rand(ctx):
    return random.random()


@fn("RANDBETWEEN", "bottom, top", "Random whole number between two numbers.")
def f_randbetween(ctx, a, b):
    return float(random.randint(int(math.ceil(num(scalar(a)))), int(math.floor(num(scalar(b))))))


@fn("FACT", "number", "Factorial.")
def f_fact(ctx, x):
    v = num(scalar(x))
    if v < 0:
        raise NUM
    return float(math.factorial(int(v)))


@fn("GCD", "number1, number2, ...", "Greatest common divisor.")
def f_gcd(ctx, *a):
    return float(math.gcd(*[int(v) for v in flat_numbers(a)]))


@fn("LCM", "number1, number2, ...", "Least common multiple.")
def f_lcm(ctx, *a):
    r = 1
    for v in flat_numbers(a):
        r = abs(r * int(v)) // math.gcd(r, int(v)) if int(v) else 0
    return float(r)


@fn("COMBIN", "n, k", "Number of combinations.")
def f_combin(ctx, n, k):
    return float(math.comb(int(num(scalar(n))), int(num(scalar(k)))))


@fn("PERMUT", "n, k", "Number of permutations.")
def f_permut(ctx, n, k):
    return float(math.perm(int(num(scalar(n))), int(num(scalar(k)))))


@fn("SUMPRODUCT", "array1, [array2], ...", "Multiplies arrays and adds the results.")
def f_sumproduct(ctx, *a):
    mats = [as_matrix(x) for x in a]
    R, C = len(mats[0]), len(mats[0][0])
    tot = 0.0
    for m in mats:
        if len(m) != R or len(m[0]) != C:
            raise VALUE
    for i in range(R):
        for j in range(C):
            p = 1.0
            for m in mats:
                v = m[i][j]
                if isinstance(v, XErr):
                    raise v
                p *= 0.0 if (v is None or isinstance(v, str)) else float(v)
            tot += p
    return tot


@fn("SUMSQ", "number1, ...", "Sum of squares.")
def f_sumsq(ctx, *a):
    return sum(v * v for v in flat_numbers(a))


@fn("SUMIF", "range, criteria, [sum_range]", "Adds cells that meet a condition.")
def f_sumif(ctx, r, c, s=None):
    rm = as_matrix(r)
    sm = as_matrix(s) if s is not None else rm
    pred = crit_pred(c)
    tot = 0.0
    for i, row in enumerate(rm):
        for j, v in enumerate(row):
            if pred(v) and i < len(sm) and j < len(sm[0]):
                x = sm[i][j]
                if isinstance(x, XErr):
                    raise x
                if isinstance(x, (int, float)) and not isinstance(x, bool):
                    tot += x
    return tot


@fn("SUMIFS", "sum_range, range1, criteria1, ...", "Adds cells that meet all conditions.")
def f_sumifs(ctx, s, *rest):
    sm = as_matrix(s)
    return float(sum(sm[i][j] for i, j in matching_indices(split_pairs(rest)) if isinstance(sm[i][j], (int, float)) and not isinstance(sm[i][j], bool)))


@fn("COUNTIF", "range, criteria", "Counts cells that meet a condition.", "Statistical")
def f_countif(ctx, r, c):
    pred = crit_pred(c)
    return float(sum(1 for v in vec(r) if pred(v)))


@fn("COUNTIFS", "range1, criteria1, ...", "Counts cells that meet all conditions.", "Statistical")
def f_countifs(ctx, *a):
    return float(len(matching_indices(split_pairs(a))))


@fn("AVERAGEIF", "range, criteria, [average_range]", "Average of cells that meet a condition.", "Statistical")
def f_averageif(ctx, r, c, s=None):
    rm = as_matrix(r)
    sm = as_matrix(s) if s is not None else rm
    pred = crit_pred(c)
    vals = [sm[i][j] for i, row in enumerate(rm) for j, v in enumerate(row) if pred(v) and isinstance(sm[i][j], (int, float)) and not isinstance(sm[i][j], bool)]
    if not vals:
        raise DIV0
    return sum(vals) / len(vals)


@fn("AVERAGEIFS", "average_range, range1, criteria1, ...", "Average of cells that meet all conditions.", "Statistical")
def f_averageifs(ctx, s, *rest):
    sm = as_matrix(s)
    vals = [sm[i][j] for i, j in matching_indices(split_pairs(rest)) if isinstance(sm[i][j], (int, float)) and not isinstance(sm[i][j], bool)]
    if not vals:
        raise DIV0
    return sum(vals) / len(vals)


@fn("MAXIFS", "max_range, range1, criteria1, ...", "Largest value that meets conditions.", "Statistical")
def f_maxifs(ctx, s, *rest):
    sm = as_matrix(s)
    vals = [sm[i][j] for i, j in matching_indices(split_pairs(rest)) if isinstance(sm[i][j], (int, float))]
    return max(vals) if vals else 0.0


@fn("MINIFS", "min_range, range1, criteria1, ...", "Smallest value that meets conditions.", "Statistical")
def f_minifs(ctx, s, *rest):
    sm = as_matrix(s)
    vals = [sm[i][j] for i, j in matching_indices(split_pairs(rest)) if isinstance(sm[i][j], (int, float))]
    return min(vals) if vals else 0.0


# --- statistics
@fn("MEDIAN", "number1, ...", "Middle value.", "Statistical")
def f_median(ctx, *a):
    n = flat_numbers(a)
    if not n:
        raise NUM
    return float(statistics.median(n))


@fn("MODE|MODE.SNGL", "number1, ...", "Most common value.", "Statistical")
def f_mode(ctx, *a):
    n = flat_numbers(a)
    best = None
    cnt = {}
    for v in n:
        cnt[v] = cnt.get(v, 0) + 1
    top = max(cnt.values()) if cnt else 0
    if top < 2:
        raise NA
    for v in n:
        if cnt[v] == top:
            return v


def _var(n, sample):
    if len(n) < (2 if sample else 1):
        raise DIV0
    m = sum(n) / len(n)
    return sum((x - m) ** 2 for x in n) / (len(n) - (1 if sample else 0))


for _names, _samp, _sd in (("STDEV|STDEV.S", True, True), ("STDEVP|STDEV.P", False, True), ("VAR|VAR.S", True, False), ("VARP|VAR.P", False, False)):
    def _mk(samp, sd):
        def f(ctx, *a):
            v = _var(flat_numbers(a), samp)
            return math.sqrt(v) if sd else v
        return f
    for _n in _names.split("|"):
        FUNCS[_n] = (_mk(_samp, _sd), False)
        HELP[_n] = (_n + "(number1, ...)", "Standard deviation." if _sd else "Variance.", "Statistical")


@fn("LARGE", "array, k", "k-th largest value.", "Statistical")
def f_large(ctx, r, k):
    n = sorted(flat_numbers([r]), reverse=True)
    k = int(num(scalar(k)))
    if not 1 <= k <= len(n):
        raise NUM
    return n[k - 1]


@fn("SMALL", "array, k", "k-th smallest value.", "Statistical")
def f_small(ctx, r, k):
    n = sorted(flat_numbers([r]))
    k = int(num(scalar(k)))
    if not 1 <= k <= len(n):
        raise NUM
    return n[k - 1]


@fn("RANK|RANK.EQ", "number, ref, [order]", "Position of a number in a list.", "Statistical")
def f_rank(ctx, x, r, order=0.0):
    x = num(scalar(x))
    n = flat_numbers([r])
    if x not in n:
        raise NA
    if num(scalar(order) if order is not None else 0):
        return float(sum(1 for v in n if v < x) + 1)
    return float(sum(1 for v in n if v > x) + 1)


def _percentile(n, k):
    n = sorted(n)
    if not n or not 0 <= k <= 1:
        raise NUM
    pos = k * (len(n) - 1)
    lo = int(math.floor(pos))
    hi = min(lo + 1, len(n) - 1)
    return n[lo] + (n[hi] - n[lo]) * (pos - lo)


@fn("PERCENTILE|PERCENTILE.INC", "array, k", "Value at a percentile (0 to 1).", "Statistical")
def f_percentile(ctx, r, k):
    return _percentile(flat_numbers([r]), num(scalar(k)))


@fn("QUARTILE|QUARTILE.INC", "array, quart", "Quartile of a list (0 to 4).", "Statistical")
def f_quartile(ctx, r, q):
    return _percentile(flat_numbers([r]), num(scalar(q)) / 4)


def _pairs(x, y):
    a, b = vec(x), vec(y)
    if len(a) != len(b):
        raise NA
    p = [(u, v) for u, v in zip(a, b) if isinstance(u, (int, float)) and isinstance(v, (int, float)) and not isinstance(u, bool) and not isinstance(v, bool)]
    if len(p) < 2:
        raise DIV0
    return p


@fn("CORREL|PEARSON", "array1, array2", "Correlation of two lists.", "Statistical")
def f_correl(ctx, x, y):
    p = _pairs(x, y)
    mx, my = sum(u for u, _ in p) / len(p), sum(v for _, v in p) / len(p)
    sxy = sum((u - mx) * (v - my) for u, v in p)
    sxx, syy = sum((u - mx) ** 2 for u, _ in p), sum((v - my) ** 2 for _, v in p)
    if sxx == 0 or syy == 0:
        raise DIV0
    return sxy / math.sqrt(sxx * syy)


def _line(ys, xs):
    p = _pairs(xs, ys)
    mx, my = sum(u for u, _ in p) / len(p), sum(v for _, v in p) / len(p)
    sxx = sum((u - mx) ** 2 for u, _ in p)
    if sxx == 0:
        raise DIV0
    slope = sum((u - mx) * (v - my) for u, v in p) / sxx
    return slope, my - slope * mx


@fn("SLOPE", "known_ys, known_xs", "Slope of the best line.", "Statistical")
def f_slope(ctx, ys, xs):
    return _line(ys, xs)[0]


@fn("INTERCEPT", "known_ys, known_xs", "Where the best line crosses the y axis.", "Statistical")
def f_intercept(ctx, ys, xs):
    return _line(ys, xs)[1]


@fn("FORECAST|FORECAST.LINEAR", "x, known_ys, known_xs", "Predicts a value on the best line.", "Statistical")
def f_forecast(ctx, x, ys, xs):
    s, i = _line(ys, xs)
    return i + s * num(scalar(x))


# --- logic
@fn("IF", "condition, if_true, [if_false]", "Chooses between two results.", "Logical", lazy=True)
def f_if(ctx, c, t=None, f=None):
    cv = scalar(c(ctx))
    if boolean(cv):
        return t(ctx) if t else True
    return f(ctx) if f else False


@fn("IFS", "condition1, value1, ...", "First value whose condition is true.", "Logical", lazy=True)
def f_ifs(ctx, *a):
    for i in range(0, len(a) - 1, 2):
        if boolean(scalar(a[i](ctx))):
            return a[i + 1](ctx)
    raise NA


@fn("IFERROR", "value, value_if_error", "Replaces an error with another result.", "Logical", lazy=True)
def f_iferror(ctx, v, alt):
    try:
        r = v(ctx)
        r = deref(r) if isinstance(r, Ref) else r
        if isinstance(r, XErr):
            return alt(ctx)
        if isinstance(r, list) and r and isinstance(r[0][0], XErr):
            return alt(ctx)
        return r
    except XErr:
        return alt(ctx)
    except (ZeroDivisionError, OverflowError, ValueError):
        return alt(ctx)


@fn("IFNA", "value, value_if_na", "Replaces #N/A with another result.", "Logical", lazy=True)
def f_ifna(ctx, v, alt):
    try:
        r = v(ctx)
        r = deref(r) if isinstance(r, Ref) else r
        if r == NA:
            return alt(ctx)
        return r
    except XErr as e:
        if e == NA:
            return alt(ctx)
        raise


@fn("SWITCH", "expression, value1, result1, ..., [default]", "Picks a result by matching a value.", "Logical", lazy=True)
def f_switch(ctx, e, *a):
    v = scalar(e(ctx))
    i = 0
    while i + 1 < len(a):
        if compare(v, scalar(a[i](ctx))) == 0:
            return a[i + 1](ctx)
        i += 2
    if i < len(a):
        return a[i](ctx)
    raise NA


@fn("AND", "logical1, ...", "TRUE if all are true.", "Logical")
def f_and(ctx, *a):
    vals = [boolean(v) for v in flat_all(a) if v is not None and v != ""]
    if not vals:
        raise VALUE
    return all(vals)


@fn("OR", "logical1, ...", "TRUE if any is true.", "Logical")
def f_or(ctx, *a):
    vals = [boolean(v) for v in flat_all(a) if v is not None and v != ""]
    if not vals:
        raise VALUE
    return any(vals)


@fn("XOR", "logical1, ...", "TRUE if an odd number are true.", "Logical")
def f_xor(ctx, *a):
    return sum(1 for v in flat_all(a) if boolean(v)) % 2 == 1


@fn("NOT", "logical", "Flips TRUE and FALSE.", "Logical")
def f_not(ctx, x):
    return not boolean(scalar(x))


@fn("TRUE", "", "The value TRUE.", "Logical")
def f_true(ctx):
    return True


@fn("FALSE", "", "The value FALSE.", "Logical")
def f_false(ctx):
    return False


# --- text
@fn("LEN", "text", "Number of characters.", "Text")
def f_len(ctx, x):
    return elementwise(lambda v: float(len(text(v))), x)


@fn("LEFT", "text, [count]", "First characters of a text.", "Text")
def f_left(ctx, s, n=1.0):
    n = int(num(scalar(n)))
    if n < 0:
        raise VALUE
    return text(scalar(s))[:n]


@fn("RIGHT", "text, [count]", "Last characters of a text.", "Text")
def f_right(ctx, s, n=1.0):
    n = int(num(scalar(n)))
    if n < 0:
        raise VALUE
    t = text(scalar(s))
    return t[len(t) - n:] if n else ""


@fn("MID", "text, start, count", "Characters from the middle of a text.", "Text")
def f_mid(ctx, s, st, n):
    st, n = int(num(scalar(st))), int(num(scalar(n)))
    if st < 1 or n < 0:
        raise VALUE
    return text(scalar(s))[st - 1:st - 1 + n]


@fn("UPPER", "text", "Capital letters.", "Text")
def f_upper(ctx, x):
    return elementwise(lambda v: text(v).upper(), x)


@fn("LOWER", "text", "Small letters.", "Text")
def f_lower(ctx, x):
    return elementwise(lambda v: text(v).lower(), x)


@fn("PROPER", "text", "First Letter Of Each Word Capital.", "Text")
def f_proper(ctx, x):
    return elementwise(lambda v: re.sub(r"[A-Za-z]+", lambda m: m.group(0).capitalize(), text(v)), x)


@fn("TRIM", "text", "Removes extra spaces.", "Text")
def f_trim(ctx, x):
    return elementwise(lambda v: re.sub(r" +", " ", text(v)).strip(" "), x)


@fn("CLEAN", "text", "Removes unprintable characters.", "Text")
def f_clean(ctx, x):
    return elementwise(lambda v: "".join(ch for ch in text(v) if ord(ch) >= 32), x)


@fn("CONCAT|CONCATENATE", "text1, [text2], ...", "Joins texts together.", "Text")
def f_concat(ctx, *a):
    return "".join(text(v) for v in flat_all(a))


@fn("TEXTJOIN", "delimiter, ignore_empty, text1, ...", "Joins texts with a separator.", "Text")
def f_textjoin(ctx, d, ign, *a):
    d, ign = text(scalar(d)), boolean(scalar(ign))
    parts = [text(v) for v in flat_all(a) if not (ign and (v is None or v == ""))]
    return d.join(parts)


@fn("FIND", "find_text, within_text, [start]", "Position of a text (matches capitals).", "Text")
def f_find(ctx, f, w, st=1.0):
    i = text(scalar(w)).find(text(scalar(f)), int(num(scalar(st))) - 1)
    if i < 0:
        raise VALUE
    return float(i + 1)


@fn("SEARCH", "find_text, within_text, [start]", "Position of a text (ignores capitals, allows * and ?).", "Text")
def f_search(ctx, f, w, st=1.0):
    pat = "".join(".*" if c == "*" else "." if c == "?" else re.escape(c) for c in text(scalar(f)))
    m = re.compile(pat, re.I | re.S).search(text(scalar(w)), int(num(scalar(st))) - 1)
    if not m:
        raise VALUE
    return float(m.start() + 1)


@fn("SUBSTITUTE", "text, old_text, new_text, [instance]", "Replaces text inside a text.", "Text")
def f_substitute(ctx, t, o, n, inst=None):
    t, o, n = text(scalar(t)), text(scalar(o)), text(scalar(n))
    if not o:
        return t
    if inst is None:
        return t.replace(o, n)
    k, idx, start = int(num(scalar(inst))), -1, 0
    for _ in range(k):
        idx = t.find(o, start)
        if idx < 0:
            return t
        start = idx + len(o)
    return t[:idx] + n + t[idx + len(o):]


@fn("REPLACE", "text, start, count, new_text", "Replaces characters by position.", "Text")
def f_replace(ctx, t, st, n, new):
    t, st, n = text(scalar(t)), int(num(scalar(st))), int(num(scalar(n)))
    return t[:st - 1] + text(scalar(new)) + t[st - 1 + n:]


@fn("REPT", "text, times", "Repeats a text.", "Text")
def f_rept(ctx, t, n):
    n = int(num(scalar(n)))
    if n < 0 or n > 32767:
        raise VALUE
    return text(scalar(t)) * n


@fn("TEXT", "value, format", "Shows a value with a format, like 0.00 or yyyy-mm-dd.", "Text")
def f_text(ctx, v, f):
    v = scalar(v)
    if isinstance(v, str):
        try:
            v = num(v)
        except XErr:
            pass
    return format_value(v, text(scalar(f)))


@fn("VALUE|NUMBERVALUE", "text", "Turns a text into a number.", "Text")
def f_value(ctx, x, *_):
    return num(scalar(x))


@fn("EXACT", "text1, text2", "TRUE if two texts are exactly the same.", "Text")
def f_exact(ctx, a, b):
    return text(scalar(a)) == text(scalar(b))


@fn("CHAR", "number", "Character for a code.", "Text")
def f_char(ctx, n):
    return chr(int(num(scalar(n))))


@fn("CODE|UNICODE", "text", "Code of the first character.", "Text")
def f_code(ctx, t):
    t = text(scalar(t))
    if not t:
        raise VALUE
    return float(ord(t[0]))


@fn("T", "value", "The value if it is text, else empty.", "Text")
def f_t(ctx, v):
    v = scalar(v)
    return v if isinstance(v, str) else ""


@fn("N", "value", "The value as a number.", "Text")
def f_n(ctx, v):
    v = scalar(v)
    return float(v) if isinstance(v, (int, float)) else 0.0


@fn("TEXTBEFORE", "text, delimiter", "Text before a delimiter.", "Text")
def f_textbefore(ctx, t, d):
    t, d = text(scalar(t)), text(scalar(d))
    i = t.find(d)
    if i < 0:
        raise NA
    return t[:i]


@fn("TEXTAFTER", "text, delimiter", "Text after a delimiter.", "Text")
def f_textafter(ctx, t, d):
    t, d = text(scalar(t)), text(scalar(d))
    i = t.find(d)
    if i < 0:
        raise NA
    return t[i + len(d):]


# --- lookup
def _approx_eq(a, b):
    return compare(a, b) == 0


@fn("VLOOKUP", "lookup_value, table, column, [exact_match]", "Looks down the first column and returns a value from another column.", "Lookup")
def f_vlookup(ctx, v, t, ci, approx=True):
    v = scalar(v)
    m = as_matrix(t)
    ci = int(num(scalar(ci)))
    if ci < 1 or ci > len(m[0]):
        raise REF
    approx = True if approx is None else boolean(scalar(approx))
    if not approx:
        for row in m:
            if row[0] is not None and compare(row[0], v) == 0 and type(row[0]) is not type(None):
                if isinstance(row[ci - 1], XErr):
                    raise row[ci - 1]
                return row[ci - 1]
        raise NA
    best = None
    for row in m:
        if row[0] is None:
            continue
        if compare(row[0], v) <= 0 and isinstance(row[0], type(v)) or (isinstance(row[0], (int, float)) and isinstance(v, (int, float)) and row[0] <= v):
            best = row
        elif compare(row[0], v) > 0:
            break
    if best is None:
        raise NA
    return best[ci - 1]


@fn("HLOOKUP", "lookup_value, table, row, [exact_match]", "Looks across the first row and returns a value from another row.", "Lookup")
def f_hlookup(ctx, v, t, ri, approx=True):
    m = as_matrix(t)
    tm = [list(c) for c in zip(*m)]
    return f_vlookup(ctx, v, tm, ri, approx)


@fn("LOOKUP", "lookup_value, lookup_vector, [result_vector]", "Finds a value in a list and returns the matching item.", "Lookup")
def f_lookup(ctx, v, lv, rv=None):
    v = scalar(v)
    keys = vec(lv)
    res = vec(rv) if rv is not None else keys
    best = None
    for i, k in enumerate(keys):
        if k is None:
            continue
        if compare(k, v) <= 0 and type(k) is type(v) or (isinstance(k, (int, float)) and isinstance(v, (int, float)) and k <= v):
            best = i
        elif compare(k, v) > 0:
            break
    if best is None:
        raise NA
    return res[best]


@fn("MATCH", "lookup_value, lookup_array, [match_type]", "Position of a value in a list.", "Lookup")
def f_match(ctx, v, arr, mt=1.0):
    v = scalar(v)
    a = vec(arr)
    mt = int(num(scalar(mt))) if mt is not None else 1
    if mt == 0:
        pred = crit_pred(v) if isinstance(v, str) and ("*" in v or "?" in v) else (lambda x: x is not None and compare(x, v) == 0)
        for i, x in enumerate(a):
            if pred(x) if isinstance(v, str) and ("*" in v or "?" in v) else (x is not None and compare(x, v) == 0 and type(x) is type(v) or (isinstance(x, (int, float)) and isinstance(v, (int, float)) and x == v)):
                return float(i + 1)
        raise NA
    best = None
    for i, x in enumerate(a):
        if x is None:
            continue
        c = compare(x, v)
        if mt == 1:
            if c <= 0 and (isinstance(x, str) == isinstance(v, str)):
                best = i
            elif c > 0:
                break
        else:
            if c >= 0 and (isinstance(x, str) == isinstance(v, str)):
                best = i
            elif c < 0:
                break
    if best is None:
        raise NA
    return float(best + 1)


@fn("XLOOKUP", "lookup_value, lookup_array, return_array, [if_not_found]", "Finds a value and returns the matching item from another list.", "Lookup", lazy=True)
def f_xlookup(ctx, v, la, ra, nf=None, *_):
    v = scalar(v(ctx))
    keys = vec(la(ctx))
    rr = ra(ctx)
    rm = as_matrix(rr)
    for i, k in enumerate(keys):
        if k is not None and compare(k, v) == 0 and isinstance(k, str) == isinstance(v, str):
            vecs = vec(rr)
            if len(rm) > 1 and len(rm[0]) > 1:
                return rm[i // len(rm[0])][i % len(rm[0])]
            return vecs[i]
    if nf:
        return nf(ctx)
    raise NA


@fn("INDEX", "array, row, [column]", "Value at a position inside a table.", "Lookup")
def f_index(ctx, arr, r=None, c=None):
    m = as_matrix(arr)
    r = int(num(scalar(r))) if r is not None else 0
    c = int(num(scalar(c))) if c is not None else 0
    if len(m) == 1 and c == 0 and r > 0:
        r, c = 1, r
    if r == 0 and c == 0:
        return m
    if r < 0 or c < 0 or r > len(m) or c > len(m[0]):
        raise REF
    if r == 0:
        return [[row[c - 1]] for row in m]
    if c == 0:
        return [m[r - 1][:]]
    return m[r - 1][c - 1]


@fn("CHOOSE", "index, value1, value2, ...", "Picks one value from a list.", "Lookup", lazy=True)
def f_choose(ctx, i, *a):
    k = int(num(scalar(i(ctx))))
    if not 1 <= k <= len(a):
        raise VALUE
    return a[k - 1](ctx)


@fn("ROW", "[reference]", "Row number.", "Lookup")
def f_row(ctx, r=None):
    if isinstance(r, Ref):
        return float(r.r1 + 1)
    return float(ctx.row + 1)


@fn("COLUMN", "[reference]", "Column number.", "Lookup")
def f_column(ctx, r=None):
    if isinstance(r, Ref):
        return float(r.c1 + 1)
    return float(ctx.col + 1)


@fn("ROWS", "array", "How many rows.", "Lookup")
def f_rows(ctx, r):
    return float(r.r2 - r.r1 + 1) if isinstance(r, Ref) else float(len(as_matrix(r)))


@fn("COLUMNS", "array", "How many columns.", "Lookup")
def f_columns(ctx, r):
    return float(r.c2 - r.c1 + 1) if isinstance(r, Ref) else float(len(as_matrix(r)[0]))


@fn("ADDRESS", "row, column, [abs_num]", "Cell name as text.", "Lookup")
def f_address(ctx, r, c, ab=1.0):
    ab = int(num(scalar(ab))) if ab is not None else 1
    return cell_name(int(num(scalar(r))) - 1, int(num(scalar(c))) - 1, ab in (1, 2), ab in (1, 3))


@fn("INDIRECT", "ref_text", "Reference given as text.", "Lookup")
def f_indirect(ctx, t):
    t = text(scalar(t)).strip()
    m = TOKEN_RE.fullmatch(t)
    if not m or m.group("ref") is None:
        raise REF
    node = Parser("", ctx.sheet).ref_node(m)
    r = node(ctx)
    if isinstance(r, XErr):
        raise r
    return r


@fn("OFFSET", "reference, rows, cols, [height], [width]", "Reference moved from another one.", "Lookup")
def f_offset(ctx, r, dr, dc, h=None, w=None):
    if not isinstance(r, Ref):
        raise VALUE
    h = int(num(scalar(h))) if h is not None else r.r2 - r.r1 + 1
    w = int(num(scalar(w))) if w is not None else r.c2 - r.c1 + 1
    r1, c1 = r.r1 + int(num(scalar(dr))), r.c1 + int(num(scalar(dc)))
    if r1 < 0 or c1 < 0 or h < 1 or w < 1:
        raise REF
    return Ref(r.sheet, r1, c1, r1 + h - 1, c1 + w - 1)


@fn("TRANSPOSE", "array", "Swaps rows and columns.", "Lookup")
def f_transpose(ctx, a):
    return [list(c) for c in zip(*as_matrix(a))]


# --- date and time
@fn("TODAY", "", "Today's date.", "Date")
def f_today(ctx):
    return float(dt_to_serial(datetime.date.today()))


@fn("NOW", "", "Date and time now.", "Date")
def f_now(ctx):
    return dt_to_serial(datetime.datetime.now())


@fn("DATE", "year, month, day", "Builds a date.", "Date")
def f_date(ctx, y, m, d):
    y = int(num(scalar(y)))
    if 0 <= y < 100:
        y += 1900
    return float(to_serial(y, num(scalar(m)), num(scalar(d))))


@fn("TIME", "hour, minute, second", "Builds a time.", "Date")
def f_time(ctx, h, m, s):
    return (num(scalar(h)) * 3600 + num(scalar(m)) * 60 + num(scalar(s))) / 86400 % 1


@fn("YEAR", "date", "Year of a date.", "Date")
def f_year(ctx, d):
    return float(date_part(scalar(d)).year)


@fn("MONTH", "date", "Month of a date (1 to 12).", "Date")
def f_month(ctx, d):
    return float(date_part(scalar(d)).month)


@fn("DAY", "date", "Day of the month.", "Date")
def f_day(ctx, d):
    return float(date_part(scalar(d)).day)


@fn("HOUR", "time", "Hour of a time.", "Date")
def f_hour(ctx, d):
    return float(date_part(scalar(d)).hour)


@fn("MINUTE", "time", "Minute of a time.", "Date")
def f_minute(ctx, d):
    return float(date_part(scalar(d)).minute)


@fn("SECOND", "time", "Second of a time.", "Date")
def f_second(ctx, d):
    return float(date_part(scalar(d)).second)


@fn("WEEKDAY", "date, [type]", "Day of the week (type 1: Sunday=1).", "Date")
def f_weekday(ctx, d, t=1.0):
    wd = date_part(scalar(d)).weekday()  # Monday=0
    t = int(num(scalar(t))) if t is not None else 1
    if t == 1:
        return float((wd + 1) % 7 + 1)
    if t == 2:
        return float(wd + 1)
    return float(wd)


@fn("WEEKNUM", "date", "Week of the year.", "Date")
def f_weeknum(ctx, d):
    dt = date_part(scalar(d))
    jan1 = datetime.datetime(dt.year, 1, 1)
    return float((dt.timetuple().tm_yday + (jan1.weekday() + 1) % 7 - 1) // 7 + 1)


@fn("DATEVALUE", "date_text", "Turns a date text into a number.", "Date")
def f_datevalue(ctx, t):
    d = parse_date(text(scalar(t)))
    if d is None:
        raise VALUE
    return float(int(d))


@fn("EDATE", "start_date, months", "Date some months later.", "Date")
def f_edate(ctx, d, n):
    return float(int(add_months(num(scalar(d)), num(scalar(n)))))


@fn("EOMONTH", "start_date, months", "Last day of a month.", "Date")
def f_eomonth(ctx, d, n):
    s = add_months(num(scalar(d)), num(scalar(n)))
    dt = serial_to_dt(s)
    nxt = datetime.datetime(dt.year + (dt.month == 12), dt.month % 12 + 1, 1)
    return float(dt_to_serial(nxt) - 1)


@fn("DAYS", "end_date, start_date", "Days between two dates.", "Date")
def f_days(ctx, e, s):
    return float(int(num(scalar(e))) - int(num(scalar(s))))


@fn("DATEDIF", "start_date, end_date, unit", "Difference between dates in Y, M or D.", "Date")
def f_datedif(ctx, s, e, u):
    a, b = date_part(scalar(s)), date_part(scalar(e))
    u = text(scalar(u)).upper()
    if b < a:
        raise NUM
    months = (b.year - a.year) * 12 + b.month - a.month - (1 if b.day < a.day else 0)
    if u == "D":
        return float((b - a).days)
    if u == "M":
        return float(months)
    if u == "Y":
        return float(months // 12)
    if u == "YM":
        return float(months % 12)
    if u == "MD":
        return float((b - datetime.datetime(b.year, b.month, 1)).days + 1 - a.day if b.day >= a.day else (b - a).days % 30)
    raise NUM


@fn("NETWORKDAYS", "start_date, end_date", "Working days between two dates.", "Date")
def f_networkdays(ctx, s, e):
    a, b = int(num(scalar(s))), int(num(scalar(e)))
    step = 1 if b >= a else -1
    n = 0
    for d in range(a, b + step, step):
        if serial_to_dt(d).weekday() < 5:
            n += 1
    return float(n * step)


# --- info
@fn("ISNUMBER", "value", "TRUE if a number.", "Information")
def f_isnumber(ctx, v):
    v = deref(v)
    return isinstance(v, (int, float)) and not isinstance(v, bool)


@fn("ISTEXT", "value", "TRUE if a text.", "Information")
def f_istext(ctx, v):
    return isinstance(deref(v), str)


@fn("ISLOGICAL", "value", "TRUE if TRUE or FALSE.", "Information")
def f_islogical(ctx, v):
    return isinstance(deref(v), bool)


@fn("ISBLANK", "value", "TRUE if the cell is empty.", "Information")
def f_isblank(ctx, v):
    return deref(v) is None


@fn("ISERROR", "value", "TRUE if any error.", "Information", lazy=True)
def f_iserror(ctx, v):
    try:
        r = deref(v(ctx))
        return isinstance(r, XErr)
    except XErr:
        return True


@fn("ISERR", "value", "TRUE if an error other than #N/A.", "Information", lazy=True)
def f_iserr(ctx, v):
    try:
        r = deref(v(ctx))
        return isinstance(r, XErr) and r != NA
    except XErr as e:
        return e != NA


@fn("ISNA", "value", "TRUE if #N/A.", "Information", lazy=True)
def f_isna(ctx, v):
    try:
        return deref(v(ctx)) == NA
    except XErr as e:
        return e == NA


@fn("ISEVEN", "number", "TRUE if even.", "Information")
def f_iseven(ctx, v):
    return int(math.floor(num(scalar(v)))) % 2 == 0


@fn("ISODD", "number", "TRUE if odd.", "Information")
def f_isodd(ctx, v):
    return int(math.floor(num(scalar(v)))) % 2 == 1


@fn("NA", "", "The #N/A error.", "Information")
def f_na(ctx):
    raise NA


# --- financial
@fn("PMT", "rate, nper, pv, [fv], [type]", "Loan payment per period.", "Financial")
def f_pmt(ctx, rate, n, pv, fv=0.0, typ=0.0):
    r, n, pv, fv, t = (num(scalar(x)) if x is not None else 0.0 for x in (rate, n, pv, fv, typ))
    if r == 0:
        return -(pv + fv) / n
    k = (1 + r) ** n
    return -(r * (pv * k + fv)) / ((1 + r * t) * (k - 1))


@fn("FV", "rate, nper, pmt, [pv], [type]", "Future value.", "Financial")
def f_fv(ctx, rate, n, pmt, pv=0.0, typ=0.0):
    r, n, p, pv, t = (num(scalar(x)) if x is not None else 0.0 for x in (rate, n, pmt, pv, typ))
    if r == 0:
        return -(pv + p * n)
    k = (1 + r) ** n
    return -(pv * k + p * (1 + r * t) * (k - 1) / r)


@fn("PV", "rate, nper, pmt, [fv], [type]", "Present value.", "Financial")
def f_pv(ctx, rate, n, pmt, fv=0.0, typ=0.0):
    r, n, p, fv, t = (num(scalar(x)) if x is not None else 0.0 for x in (rate, n, pmt, fv, typ))
    if r == 0:
        return -(fv + p * n)
    k = (1 + r) ** n
    return -(fv + p * (1 + r * t) * (k - 1) / r) / k


@fn("NPER", "rate, pmt, pv, [fv], [type]", "Number of payment periods.", "Financial")
def f_nper(ctx, rate, pmt, pv, fv=0.0, typ=0.0):
    r, p, pv, fv, t = (num(scalar(x)) if x is not None else 0.0 for x in (rate, pmt, pv, fv, typ))
    if r == 0:
        return -(pv + fv) / p
    a = p * (1 + r * t) / r
    return math.log((a - fv) / (pv + a)) / math.log(1 + r)


@fn("NPV", "rate, value1, ...", "Net present value.", "Financial")
def f_npv(ctx, rate, *a):
    r = num(scalar(rate))
    return sum(v / (1 + r) ** (i + 1) for i, v in enumerate(flat_numbers(a)))


@fn("IRR", "values, [guess]", "Internal rate of return.", "Financial")
def f_irr(ctx, vals, guess=0.1):
    cf = flat_numbers([vals])
    if not (any(v > 0 for v in cf) and any(v < 0 for v in cf)):
        raise NUM
    x = num(scalar(guess)) if guess is not None else 0.1
    for _ in range(100):
        f = sum(v / (1 + x) ** i for i, v in enumerate(cf))
        d = sum(-i * v / (1 + x) ** (i + 1) for i, v in enumerate(cf))
        if d == 0:
            break
        nx = x - f / d
        if abs(nx - x) < 1e-10:
            return nx
        x = nx
    raise NUM


@fn("RATE", "nper, pmt, pv, [fv], [type]", "Interest rate per period.", "Financial")
def f_rate(ctx, n, pmt, pv, fv=0.0, typ=0.0, guess=0.1):
    n, p, pv, fv, t = (num(scalar(x)) if x is not None else 0.0 for x in (n, pmt, pv, fv, typ))
    x = 0.1
    for _ in range(100):
        k = (1 + x) ** n
        f = pv * k + p * (1 + x * t) * (k - 1) / x + fv
        h = 1e-6
        k2 = (1 + x + h) ** n
        f2 = pv * k2 + p * (1 + (x + h) * t) * (k2 - 1) / (x + h) + fv
        d = (f2 - f) / h
        if d == 0:
            break
        nx = x - f / d
        if abs(nx - x) < 1e-12:
            return nx
        x = nx
    raise NUM


CATEGORIES = ["Math", "Statistical", "Logical", "Text", "Lookup", "Date", "Financial", "Information"]
