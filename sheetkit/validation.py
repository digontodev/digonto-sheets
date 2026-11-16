"""Data validation: rules per sheet (list, whole number, decimal, date, text length, custom formula)."""
import datetime

from . import formula as F
from .cfmt import parse_range
from .model import parse_input

KINDS = ["any", "list", "whole", "decimal", "date", "length", "custom"]
KIND_LABEL = {"any": "Any value", "list": "List", "whole": "Whole number", "decimal": "Decimal", "date": "Date", "length": "Text length", "custom": "Custom formula"}
OPS = ["between", "notBetween", "equal", "notEqual", "greater", "less", "ge", "le"]
OP_LABEL = {"between": "between", "notBetween": "not between", "equal": "equal to", "notEqual": "not equal to", "greater": "greater than", "less": "less than",
            "ge": "greater than or equal to", "le": "less than or equal to"}


def new_rule(rng):
    return {"range": rng, "kind": "list", "op": "between", "a": "", "b": "", "list": "", "blank": True, "dropdown": True, "style": "stop",
            "title": "", "msg": "", "in_title": "", "in_msg": ""}


def rects(rule):
    out = []
    for part in str(rule.get("range", "")).replace(",", " ").split():
        r = parse_range(part)
        if r:
            out.append(r)
    return out


def rule_at(sheet, r, c):
    for rule in reversed(sheet.dv):
        for (r1, c1, r2, c2) in rects(rule):
            if r1 <= r <= r2 and c1 <= c <= c2:
                return rule
    return None


def _num(wb, sheet, text, r, c):
    """number from a bound typed in the dialog: literal, date text or =formula"""
    t = str(text).strip()
    if not t:
        return None
    if t.startswith("="):
        v = F.evaluate(t, F.Ctx(wb, sheet.name, r, c))
        return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None
    raw, _ = parse_input(t)
    return raw if isinstance(raw, (int, float)) and not isinstance(raw, bool) else None


def list_items(wb, sheet, rule, r=0, c=0):
    src = str(rule.get("list", "")).strip()
    if not src:
        return []
    s = src[1:] if src.startswith("=") else src
    sh_name = sheet.name
    ref = s
    if "!" in s:
        sh_name, _, ref = s.rpartition("!")
        sh_name = sh_name.strip("'")
    rg = parse_range(ref) if ref and not ("," in ref or '"' in ref) else None
    if rg and (src.startswith("=") or ":" in ref or ref.replace("$", "").isalnum() and any(ch.isdigit() for ch in ref)):
        target = wb.sheet(sh_name)
        if target is not None:
            r1, c1, r2, c2 = rg
            out = []
            for rr in range(r1, min(r2, r1 + 5000) + 1):
                for cc in range(c1, c2 + 1):
                    d = wb.display(sh_name, rr, cc)
                    d = d[0] if isinstance(d, tuple) else d
                    if d != "":
                        out.append(d)
            return out
    return [x.strip().strip('"') for x in src.strip('"').split(",") if x.strip()]


def check(wb, sheet, r, c, text):
    """None when the typed text is allowed, else the rule that rejects it"""
    rule = rule_at(sheet, r, c)
    if rule is None or rule["kind"] == "any":
        return None
    if text is None or str(text).strip() == "":
        return None if rule.get("blank", True) else rule
    t = str(text)
    kind = rule["kind"]
    try:
        if kind == "list":
            items = list_items(wb, sheet, rule, r, c)
            low = t.strip().lower()
            return None if any(low == i.lower() for i in items) else rule
        if kind == "custom":
            home = rects(rule)[0] if rects(rule) else (r, c, r, c)
            f = str(rule.get("a", "")).strip()
            if not f:
                return None
            if not f.startswith("="):
                f = "=" + f
            f = F.shift_formula(f, r - home[0], c - home[1])
            raw, _ = parse_input(t)
            # evaluate with the new value visible: temporary override
            key = (sheet.name, r, c)
            old = wb.sheet(sheet.name).cells.get((r, c))
            wb.sheet(sheet.name).cells[(r, c)] = (raw, old[1] if old else 0)
            wb._cache.clear()
            try:
                v = F.evaluate(f, F.Ctx(wb, sheet.name, r, c))
            finally:
                if old is None:
                    wb.sheet(sheet.name).cells.pop((r, c), None)
                else:
                    wb.sheet(sheet.name).cells[(r, c)] = old
                wb._cache.clear()
            return None if (v is True or (isinstance(v, (int, float)) and not isinstance(v, bool) and v != 0)) else rule
        raw, _ = parse_input(t)
        if kind == "length":
            val = len(t)
        else:
            if isinstance(raw, bool) or not isinstance(raw, (int, float)):
                return rule
            val = raw
            if kind == "whole" and val != int(val):
                return rule
        a = _num(wb, sheet, rule.get("a", ""), r, c)
        b = _num(wb, sheet, rule.get("b", ""), r, c)
        op = rule.get("op", "between")
        if a is None:
            return None
        if op in ("between", "notBetween"):
            if b is None:
                return None
            ok = min(a, b) <= val <= max(a, b)
            return None if ok == (op == "between") else rule
        ok = {"equal": val == a, "notEqual": val != a, "greater": val > a, "less": val < a, "ge": val >= a, "le": val <= a}[op]
        return None if ok else rule
    except Exception:
        return None


def default_message(rule):
    k = rule["kind"]
    if k == "list":
        return "Pick one of the values from the list."
    if k == "custom":
        return "This value does not meet the rule for this cell."
    what = {"whole": "a whole number", "decimal": "a number", "date": "a date", "length": "text with a length"}[k]
    op = rule.get("op", "between")
    if op in ("between", "notBetween"):
        return f"Enter {what} {OP_LABEL[op]} {rule.get('a')} and {rule.get('b')}."
    return f"Enter {what} {OP_LABEL[op]} {rule.get('a')}."


# ---------------------------------------------------------------- xlsx mapping
XL_TYPE = {"list": "list", "whole": "whole", "decimal": "decimal", "date": "date", "length": "textLength", "custom": "custom"}
XL_TYPE_BACK = {v: k for k, v in XL_TYPE.items()}
XL_OP = {"between": "between", "notBetween": "notBetween", "equal": "equal", "notEqual": "notEqual", "greater": "greaterThan", "less": "lessThan",
         "ge": "greaterThanOrEqual", "le": "lessThanOrEqual"}
XL_OP_BACK = {v: k for k, v in XL_OP.items()}
