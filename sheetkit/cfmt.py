"""Conditional formatting rules evaluated for the cells on screen."""
from . import formula as F


def parse_range(text, default_sheet=None):
    """'A1:B5' / 'C3' -> (r1, c1, r2, c2) or None"""
    try:
        text = text.strip().replace("$", "")
        if ":" in text:
            a, b = text.split(":")
            r1, c1, *_ = F.parse_cell(a)
            r2, c2, *_ = F.parse_cell(b)
        else:
            r1, c1, *_ = F.parse_cell(text)
            r2, c2 = r1, c1
        return min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)
    except Exception:
        return None


def range_text(r1, c1, r2, c2):
    a = F.cell_name(r1, c1)
    return a if (r1, c1) == (r2, c2) else a + ":" + F.cell_name(r2, c2)


class CF:
    """per sheet: given a cell and its value tell what to change (fg, bg, bar)"""

    def __init__(self):
        self.ver = -1
        self.rules = []

    def prepare(self, wb, sheet):
        key = (wb.version, id(sheet), len(sheet.cf))
        if key == self.ver:
            return
        self.ver = key
        self.rules = []
        for rule in sheet.cf:
            rg = []
            for part in str(rule["range"]).split(","):
                p = parse_range(part)
                if p:
                    rg.append(p)
            if not rg:
                continue
            info = {"rule": rule, "ranges": rg, "stats": None}
            k = rule["kind"]
            if k in ("duplicate values", "top 10", "bottom 10", "above average", "below average", "color scale", "data bars"):
                vals = []
                for (r1, c1, r2, c2) in rg:
                    for (r, c), cell in sheet.cells.items():
                        if r1 <= r <= r2 and c1 <= c <= c2:
                            v = wb.cell_value(sheet.name, r, c)
                            if k == "duplicate values":
                                vals.append(v if not isinstance(v, str) else v.lower())
                            elif isinstance(v, (int, float)) and not isinstance(v, bool):
                                vals.append(v)
                if k == "duplicate values":
                    seen, dup = set(), set()
                    for v in vals:
                        if v is None or v == "":
                            continue
                        (dup if v in seen else seen).add(v)
                    info["stats"] = dup
                elif vals:
                    s = sorted(vals)
                    n = int(float(rule.get("a") or 10)) if k in ("top 10", "bottom 10") else 0
                    info["stats"] = {"min": s[0], "max": s[-1], "avg": sum(vals) / len(vals), "cut": (s[-n] if k == "top 10" and n else s[n - 1] if n else None)}
            self.rules.append(info)

    def has_rules(self):
        return bool(self.rules)

    def apply(self, r, c, value):
        """-> dict(fg=, bg=, bar=(fraction, color)) or None"""
        out = None
        for info in self.rules:
            if not any(r1 <= r <= r2 and c1 <= c <= c2 for (r1, c1, r2, c2) in info["ranges"]):
                continue
            rule, k = info["rule"], info["rule"]["kind"]
            hit = False
            try:
                if k == "is empty":
                    hit = value is None or value == ""
                elif k == "is not empty":
                    hit = not (value is None or value == "")
                elif k == "text contains":
                    hit = isinstance(value, str) and str(rule["a"]).lower() in value.lower()
                elif k in ("greater than", "less than", "equal to", "not equal to", "between"):
                    a = _num(rule["a"])
                    if isinstance(value, (int, float)) and not isinstance(value, bool) and a is not None:
                        if k == "greater than":
                            hit = value > a
                        elif k == "less than":
                            hit = value < a
                        elif k == "equal to":
                            hit = value == a
                        elif k == "not equal to":
                            hit = value != a
                        else:
                            b = _num(rule["b"])
                            hit = b is not None and min(a, b) <= value <= max(a, b)
                    elif k in ("equal to", "not equal to") and a is None and isinstance(value, str):
                        eq = value.lower() == str(rule["a"]).lower()
                        hit = eq if k == "equal to" else not eq
                elif k == "duplicate values":
                    v = value.lower() if isinstance(value, str) else value
                    hit = info["stats"] is not None and v in info["stats"]
                elif k in ("top 10", "bottom 10", "above average", "below average") and isinstance(value, (int, float)) and info["stats"]:
                    s = info["stats"]
                    hit = {"top 10": lambda: value >= s["cut"] if s["cut"] is not None else False, "bottom 10": lambda: value <= s["cut"] if s["cut"] is not None else False,
                           "above average": lambda: value > s["avg"], "below average": lambda: value < s["avg"]}[k]()
                elif k == "color scale" and isinstance(value, (int, float)) and not isinstance(value, bool) and info["stats"]:
                    s = info["stats"]
                    t = 0.5 if s["max"] == s["min"] else (value - s["min"]) / (s["max"] - s["min"])
                    out = dict(out or {}, bg=_mix(rule.get("lo", "#F8696B"), rule.get("mid"), rule.get("hi", "#63BE7B"), t))
                    continue
                elif k == "data bars" and isinstance(value, (int, float)) and not isinstance(value, bool) and info["stats"]:
                    s = info["stats"]
                    lo = min(0, s["min"])
                    span = (s["max"] - lo) or 1
                    out = dict(out or {}, bar=((value - lo) / span, rule.get("hi", "#63BE7B")))
                    continue
            except (TypeError, ValueError, KeyError):
                hit = False
            if hit:
                out = dict(out or {})
                if rule.get("fg"):
                    out["fg"] = rule["fg"]
                if rule.get("bg"):
                    out["bg"] = rule["bg"]
                if rule.get("bold"):
                    out["bold"] = True
        return out


def _num(x):
    try:
        return float(str(x).replace(",", ""))
    except ValueError:
        return None


def _mix(lo, mid, hi, t):
    def rgb(h):
        h = h.lstrip("#")
        return [int(h[i:i + 2], 16) for i in (0, 2, 4)]
    a, b = rgb(lo), rgb(hi)
    if mid:
        m = rgb(mid)
        if t < 0.5:
            a, b, t = a, m, t * 2
        else:
            a, b, t = m, b, (t - 0.5) * 2
    return "#%02X%02X%02X" % tuple(int(round(a[i] + (b[i] - a[i]) * t)) for i in range(3))
