"""Number / date formatting with Excel style format codes (General, 0.00, #,##0, 0%, yyyy-mm-dd, hh:mm ...)."""
import datetime
import re
from decimal import ROUND_HALF_UP, Decimal

EPOCH = datetime.datetime(1899, 12, 30)
MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]

FORMATS = [  # label, code
    ("General", "General"), ("Number  1234.50", "0.00"), ("Whole number  1235", "0"), ("Thousands  1,234.50", "#,##0.00"),
    ("Currency  $1,234.50", "$#,##0.00"), ("Percent  12.34%", "0.00%"), ("Percent  12%", "0%"),
    ("Scientific  1.23E+03", "0.00E+00"), ("Date  2026-03-05", "yyyy-mm-dd"), ("Date  3/5/2026", "m/d/yyyy"),
    ("Date  05/03/2026", "dd/mm/yyyy"), ("Long date", "dddd, mmmm d, yyyy"), ("Time  13:45:00", "hh:mm:ss"),
    ("Date and time", "yyyy-mm-dd hh:mm"), ("Text", "@"),
]


def serial_to_dt(v):
    return EPOCH + datetime.timedelta(days=float(v))


def dt_to_serial(d):
    if isinstance(d, datetime.datetime):
        delta = d - EPOCH
        return delta.days + delta.seconds / 86400 + delta.microseconds / 86400e6
    return (datetime.datetime(d.year, d.month, d.day) - EPOCH).days


def round_half_up(v, nd=0):
    try:
        q = Decimal(1).scaleb(-nd)
        return float(Decimal(repr(float(v))).quantize(q, rounding=ROUND_HALF_UP))
    except Exception:
        return v


def general(v):
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, int):
        return str(v)
    if v == int(v) and abs(v) < 1e15:
        return str(int(v))
    a = abs(v)
    if a >= 1e11 or a < 1e-9:
        s = "%.5E" % v
        m, e = s.split("E")
        m = m.rstrip("0").rstrip(".")
        return f"{m}E{e[0]}{e[1:].lstrip('0').rjust(2, '0')}"
    s = "%.10g" % v
    if "e" in s:
        s = ("%.10f" % v).rstrip("0").rstrip(".")
    return s


def _split_sections(code):
    out, cur, q = [], "", False
    i = 0
    while i < len(code):
        ch = code[i]
        if ch == '"':
            q = not q
        if ch == "\\" and i + 1 < len(code):
            cur += code[i:i + 2]
            i += 2
            continue
        if ch == ";" and not q:
            out.append(cur)
            cur = ""
        else:
            cur += ch
        i += 1
    out.append(cur)
    return out


def _tokens(sec):
    """list of (kind, text): lit | ph (placeholder char) | other"""
    toks, i = [], 0
    while i < len(sec):
        ch = sec[i]
        if ch == '"':
            j = sec.find('"', i + 1)
            j = len(sec) if j < 0 else j
            toks.append(("lit", sec[i + 1:j]))
            i = j + 1
        elif ch == "\\":
            toks.append(("lit", sec[i + 1:i + 2]))
            i += 2
        elif ch == "[":
            j = sec.find("]", i)
            j = len(sec) - 1 if j < 0 else j
            inner = sec[i + 1:j]
            if re.match(r"^[hms]+$", inner, re.I):
                toks.append(("el", inner.lower()))
            elif inner.startswith("$") and "-" in inner:
                toks.append(("lit", inner[1:inner.index("-")]))
            i = j + 1
        elif ch == "_":
            toks.append(("lit", " "))
            i += 2
        elif ch == "*":
            i += 2
        else:
            toks.append(("ch", ch))
            i += 1
    return toks


def _is_date(toks):
    s = "".join(t for k, t in toks if k == "ch").lower()
    s = re.sub(r"e[+-]", "", s)
    return bool(re.search(r"[ymdhs]", s)) or any(k == "el" for k, _ in toks)


def _fmt_date(v, toks):
    try:
        dt = serial_to_dt(v)
    except (OverflowError, ValueError):
        return "#####"
    s = "".join(t if k == "ch" else "\x00" + t.replace("\x00", "") + "\x01" if k == "lit" else "" for k, t in toks)
    ampm = bool(re.search(r"am/pm|a/p", s, re.I))
    out, i = [], 0
    low = s.lower()
    # secs rounding
    total_sec = round((dt - EPOCH).total_seconds()) if False else None
    last_hour = None
    parts = []
    while i < len(s):
        if s[i] == "\x00":
            j = s.index("\x01", i)
            parts.append(("lit", s[i + 1:j]))
            i = j + 1
            continue
        m = re.match(r"yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|am/pm|a/p|\.0+", low[i:])
        if m:
            parts.append(("tok", m.group(0), s[i:i + len(m.group(0))]))
            i += len(m.group(0))
        else:
            parts.append(("lit", s[i]))
            i += 1
    # decide minute vs month for 'm' / 'mm'
    res = []
    for idx, p in enumerate(parts):
        if p[0] == "lit":
            res.append(p[1])
            continue
        t = p[1]
        if t in ("m", "mm"):
            prev = next((q for q in reversed(parts[:idx]) if q[0] == "tok"), None)
            nxt = next((q for q in parts[idx + 1:] if q[0] == "tok"), None)
            if (prev and prev[1] in ("h", "hh")) or (nxt and nxt[1] in ("s", "ss")):
                res.append(("%02d" if t == "mm" else "%d") % dt.minute)
            else:
                res.append(("%02d" if t == "mm" else "%d") % dt.month)
        elif t == "yyyy":
            res.append("%04d" % dt.year)
        elif t == "yy":
            res.append("%02d" % (dt.year % 100))
        elif t == "mmm":
            res.append(MONTHS[dt.month - 1][:3])
        elif t == "mmmm":
            res.append(MONTHS[dt.month - 1])
        elif t == "mmmmm":
            res.append(MONTHS[dt.month - 1][0])
        elif t == "d":
            res.append(str(dt.day))
        elif t == "dd":
            res.append("%02d" % dt.day)
        elif t == "ddd":
            res.append(DAYS[dt.weekday()][:3])
        elif t == "dddd":
            res.append(DAYS[dt.weekday()])
        elif t in ("h", "hh"):
            h = dt.hour
            if ampm:
                h = h % 12 or 12
            res.append(("%02d" if t == "hh" else "%d") % h)
        elif t == "ss":
            res.append("%02d" % dt.second)
        elif t == "s":
            res.append(str(dt.second))
        elif t in ("am/pm", "a/p"):
            ap = "AM" if dt.hour < 12 else "PM"
            res.append(ap if t == "am/pm" else ap[0])
        elif t.startswith("."):
            res.append(("%.6f" % (dt.microsecond / 1e6))[1:1 + len(t)])
    return "".join(res)


def _fmt_number(v, toks, neg_auto):
    pct = sum(1 for k, t in toks if k == "ch" and t == "%")
    sci = None
    for i, (k, t) in enumerate(toks):
        if k == "ch" and t in "eE" and i + 1 < len(toks) and toks[i + 1][0] == "ch" and toks[i + 1][1] in "+-":
            sci = i
            break
    # placeholder run
    ph_idx = [i for i, (k, t) in enumerate(toks) if k == "ch" and t in "0#?" and (sci is None or i < sci)]
    if not ph_idx:
        txt = "".join(t for k, t in toks if k in ("lit", "ch"))
        return txt
    first, last = ph_idx[0], ph_idx[-1]
    dot = next((i for i in range(first, last + 1) if toks[i] == ("ch", ".")), None)
    int_part = [t for i, (k, t) in enumerate(toks[first:(dot if dot is not None else last + 1)], first) if k == "ch" and t in "0#?"]
    dec_part = [t for i, (k, t) in enumerate(toks[dot + 1:last + 1], dot + 1) if k == "ch" and t in "0#?"] if dot is not None else []
    thousands = any(k == "ch" and t == "," for k, t in toks[first:last + 1])
    trailing_commas = 0
    j = last + 1
    while j < len(toks) and toks[j] == ("ch", ","):
        trailing_commas += 1
        j += 1
    x = abs(v) * (100 ** pct) / (1000 ** trailing_commas)
    nd = len(dec_part)
    prefix = "".join(t for k, t in toks[:first] if k in ("lit", "ch") and t not in "%")
    prefix = "".join(("%" if (k == "ch" and t == "%") else t) for k, t in toks[:first] if k in ("lit", "ch"))
    if sci is not None:
        exp_digits = sum(1 for k, t in toks[sci + 2:] if k == "ch" and t == "0")
        sign_always = toks[sci + 1][1] == "+"
        e = 0
        if x != 0:
            ni = len(int_part) or 1
            e = int(("%e" % x).split("e")[1]) - (ni - 1) if ni > 1 else int(("%e" % x).split("e")[1])
        m = x / (10 ** e) if x else 0
        mt = f"{round_half_up(m, nd):.{nd}f}"
        if float(mt) >= 10 ** max(len(int_part), 1):
            e += 1
            mt = f"{round_half_up(x / (10 ** e), nd):.{nd}f}"
        es = ("+" if e >= 0 else "-") if (sign_always or e < 0) else ""
        body = f"{mt}E{es}{abs(e):0{max(exp_digits, 1)}d}"
        tail = "".join(t for k, t in toks[sci + 2:] if k == "lit")
        return ("-" if neg_auto and v < 0 else "") + prefix + body + tail
    r = round_half_up(x, nd)
    s = f"{r:,.{nd}f}" if thousands else f"{r:.{nd}f}"
    ip, _, dp = s.partition(".")
    min_int = sum(1 for t in int_part if t == "0")
    qs = sum(1 for t in int_part if t == "?")
    ip_digits = ip.replace(",", "")
    if min_int == 0 and ip_digits == "0" and r != 0 or (min_int == 0 and ip_digits == "0" and nd > 0):
        ip = ""
    elif len(ip_digits) < min_int:
        ip_digits = ip_digits.rjust(min_int, "0")
        ip = f"{int(ip_digits):,}".rjust(min_int, "0") if thousands and len(ip_digits) > 3 else ip_digits
    if min_int == 0 and ip_digits == "0" and nd == 0:
        ip = "" if not int_part else ip
        if r == 0 and not any(t == "0" for t in int_part + dec_part):
            ip = ""
    # optional decimals (#)
    if dp:
        keep = len(dp)
        while keep > 0 and dec_part[keep - 1] == "#" and dp[keep - 1] == "0":
            keep -= 1
        dp = dp[:keep]
    body = ip + ("." + dp if dp or (dec_part and all(t == "0" for t in dec_part) and False) else "")
    if dot is not None and not dp and not body:
        body = ""
    # literal chars interleaved inside the run (e.g. 000-0000) are rare; ignore
    suffix = "".join(("%" if (k == "ch" and t == "%") else t) for k, t in toks[last + 1:] if k in ("lit", "ch") and t != ",")
    sign = "-" if neg_auto and v < 0 and r != 0 else ""
    return sign + prefix + body + suffix


def format_value(v, code="General"):
    """text shown for a value with a format code"""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if not code or code.lower() == "general":
        return general(v) if isinstance(v, (int, float)) else str(v)
    secs = _split_sections(code)
    if isinstance(v, str):
        sec = secs[3] if len(secs) > 3 else (secs[0] if "@" in secs[0] else None)
        if sec is None:
            return v
        return "".join(v if (k == "ch" and t == "@") else t for k, t in _tokens(sec) if k in ("lit", "ch") or t)
    # numbers
    if len(secs) == 1:
        sec, neg_auto = secs[0], True
    elif len(secs) == 2:
        sec, neg_auto = (secs[0], False) if v >= 0 else (secs[1], False)
    else:
        sec, neg_auto = (secs[0] if v > 0 else secs[1] if v < 0 else secs[2]), False
    if sec.strip() == "@":
        return general(v)
    toks = _tokens(sec)
    if _is_date(toks):
        return _fmt_date(v, toks)
    return _fmt_number(v, toks, neg_auto)


def is_date_code(code):
    return _is_date(_tokens(_split_sections(code or "")[0])) if code and code.lower() != "general" else False
