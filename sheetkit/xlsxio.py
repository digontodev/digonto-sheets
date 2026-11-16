"""Read and write .xlsx (zip + XML, no extra libraries) and .csv"""
import csv
import io
import math
import re
import zipfile
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

from . import formula as F
from . import xlsxcharts as xcharts
from .model import DEFAULT, DEFAULT_H, DEFAULT_W, Sheet, Style, Workbook, parse_input

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PNS = "{http://schemas.openxmlformats.org/package/2006/relationships}"

BUILTIN = {0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00", 9: "0%", 10: "0.00%", 11: "0.00E+00", 12: "# ?/?", 13: "# ??/??",
           14: "m/d/yyyy", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy", 18: "h:mm AM/PM", 19: "h:mm:ss AM/PM", 20: "h:mm", 21: "h:mm:ss",
           22: "m/d/yyyy h:mm", 37: "#,##0;(#,##0)", 38: "#,##0;[Red](#,##0)", 39: "#,##0.00;(#,##0.00)", 40: "#,##0.00;[Red](#,##0.00)",
           45: "mm:ss", 46: "[h]:mm:ss", 47: "mm:ss.0", 49: "@"}
THEME = ["FFFFFF", "000000", "E7E6E6", "44546A", "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47", "0563C1", "954F72"]
INDEXED = ["000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
           "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080", "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
           "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF", "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
           "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696", "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333"]


def _tint(hexcol, tint):
    r, g, b = int(hexcol[0:2], 16), int(hexcol[2:4], 16), int(hexcol[4:6], 16)
    f = (lambda x: x * (1 + tint)) if tint < 0 else (lambda x: x + (255 - x) * tint)
    return "#%02X%02X%02X" % tuple(max(0, min(255, int(round(f(v))))) for v in (r, g, b))


def _color(el):
    if el is None:
        return None
    rgb = el.get("rgb")
    if rgb:
        return "#" + rgb[-6:].upper()
    th = el.get("theme")
    if th is not None and int(th) < len(THEME):
        return _tint(THEME[int(th)], float(el.get("tint", 0)))
    ix = el.get("indexed")
    if ix is not None and int(ix) < len(INDEXED):
        return "#" + INDEXED[int(ix)]
    return None


# ---------------------------------------------------------------- reading
def read_xlsx(path):
    z = zipfile.ZipFile(path)
    names = set(z.namelist())
    wb = Workbook()
    wb.sheets = []
    shared = []
    if "xl/sharedStrings.xml" in names:
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
        for si in root.findall(NS + "si"):
            shared.append("".join(t.text or "" for t in si.iter(NS + "t")))
    fonts, fills, borders, numfmts, xfs = [], [], [], {}, []
    if "xl/styles.xml" in names:
        st = ET.fromstring(z.read("xl/styles.xml"))
        for nf in st.iter(NS + "numFmt"):
            numfmts[int(nf.get("numFmtId"))] = nf.get("formatCode")
        for f in (st.find(NS + "fonts") or []):
            u = f.find(NS + "u")
            fonts.append(dict(b=f.find(NS + "b") is not None and f.find(NS + "b").get("val") != "0", i=f.find(NS + "i") is not None and f.find(NS + "i").get("val") != "0",
                              u=u is not None and u.get("val") != "none", s=f.find(NS + "strike") is not None,
                              name=(f.find(NS + "name").get("val") if f.find(NS + "name") is not None else None),
                              size=float(f.find(NS + "sz").get("val")) if f.find(NS + "sz") is not None else None, color=_color(f.find(NS + "color"))))
        for f in (st.find(NS + "fills") or []):
            pf = f.find(NS + "patternFill")
            col = None
            if pf is not None and pf.get("patternType") == "solid":
                col = _color(pf.find(NS + "fgColor"))
            fills.append(col)
        for b in (st.find(NS + "borders") or []):
            d = {}
            for side in ("left", "top", "right", "bottom"):
                e = b.find(NS + side)
                d[side] = (_color(e.find(NS + "color")) or "#000000") if e is not None and e.get("style") else None
            borders.append(d)
        for xf in (st.find(NS + "cellXfs") or []):
            xfs.append(xf)
    style_cache = {}

    def style_for(idx):
        if idx in style_cache:
            return style_cache[idx]
        if idx >= len(xfs):
            return 0
        xf = xfs[idx]
        nid = int(xf.get("numFmtId", 0))
        fmt = numfmts.get(nid) or BUILTIN.get(nid, "General")
        fo = fonts[int(xf.get("fontId", 0))] if fonts and int(xf.get("fontId", 0)) < len(fonts) else {}
        bg = fills[int(xf.get("fillId", 0))] if fills and int(xf.get("fillId", 0)) < len(fills) else None
        bd = borders[int(xf.get("borderId", 0))] if borders and int(xf.get("borderId", 0)) < len(borders) else {}
        al = xf.find(NS + "alignment")
        h = al.get("horizontal") if al is not None else None
        h = {"centerContinuous": "center", "general": None, "fill": None, "justify": "left", "distributed": "center"}.get(h, h)
        v = al.get("vertical") if al is not None else None
        v = {"center": "middle"}.get(v, v)
        s = Style(bold=fo.get("b", False), italic=fo.get("i", False), underline=fo.get("u", False), strike=fo.get("s", False),
                  color=fo.get("color") if fo.get("color") not in ("#000000",) else None, bg=bg,
                  font=fo.get("name") if fo.get("name") not in ("Calibri", "Arial") else None,
                  size=fo.get("size") if fo.get("size") not in (11.0, 10.0) else None, h=h, v=v,
                  wrap=bool(al is not None and al.get("wrapText") in ("1", "true")), fmt=fmt,
                  bl=bd.get("left"), bt=bd.get("top"), br=bd.get("right"), bb=bd.get("bottom"),
                  indent=int(al.get("indent", 0)) if al is not None else 0)
        style_cache[idx] = wb.style_id(s)
        return style_cache[idx]

    root = ET.fromstring(z.read("xl/workbook.xml"))
    rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels")).findall(PNS + "Relationship")}
    wb.date1904 = (root.find(NS + "workbookPr") is not None and root.find(NS + "workbookPr").get("date1904") in ("1", "true"))
    sheet_files = []
    for s in root.find(NS + "sheets"):
        tgt = rels.get(s.get(RNS + "id"), "")
        tgt = tgt.lstrip("/")
        if not tgt.startswith("xl/"):
            tgt = "xl/" + tgt
        sheet_files.append((s.get("name"), tgt))
    dn = root.find(NS + "definedNames")
    if dn is not None:
        for d in dn:
            if d.get("localSheetId") is None and not d.get("name").startswith("_xlnm") and d.text:
                wb.names[d.get("name").upper()] = d.text
    for name, tgt in sheet_files:
        sh = Sheet(name)
        wb.sheets.append(sh)
        if tgt not in names:
            continue
        sroot = ET.fromstring(z.read(tgt))
        pr = sroot.find(NS + "sheetPr")
        if pr is not None and pr.find(NS + "tabColor") is not None:
            sh.tab_color = _color(pr.find(NS + "tabColor"))
        sv = sroot.find(NS + "sheetViews")
        if sv is not None:
            v0 = sv.find(NS + "sheetView")
            if v0 is not None:
                sh.grid = v0.get("showGridLines") not in ("0", "false")
                pane = v0.find(NS + "pane")
                if pane is not None and pane.get("state", "") in ("frozen", "frozenSplit"):
                    sh.freeze = (int(float(pane.get("ySplit", 0))), int(float(pane.get("xSplit", 0))))
        fp = sroot.find(NS + "sheetFormatPr")
        defw = float(fp.get("defaultColWidth")) if fp is not None and fp.get("defaultColWidth") else None
        for col in (sroot.find(NS + "cols") or []):
            w = col.get("width")
            for c in range(int(col.get("min")) - 1, min(int(col.get("max")), 16384)):
                if col.get("hidden") in ("1", "true"):
                    sh.hidden_cols.add(c)
                if w and c < 2000:
                    sh.colw[c] = int(float(w) * 7 + 5)
        shared_f = {}
        for row in sroot.iter(NS + "row"):
            r = int(row.get("r")) - 1
            if row.get("ht") and row.get("customHeight", "1") not in ("0", "false"):
                sh.rowh[r] = int(float(row.get("ht")) * 96 / 72)
            if row.get("hidden") in ("1", "true"):
                sh.hidden_rows.add(r)
            for c_el in row.findall(NS + "c"):
                ref = c_el.get("r")
                m = re.match(r"([A-Z]+)(\d+)", ref)
                cc, rr = F.col_index(m.group(1)), int(m.group(2)) - 1
                t = c_el.get("t", "n")
                sid = style_for(int(c_el.get("s", 0)))
                f_el = c_el.find(NS + "f")
                v_el = c_el.find(NS + "v")
                raw = None
                if f_el is not None:
                    ftxt = f_el.text
                    if f_el.get("t") == "shared":
                        si = f_el.get("si")
                        if ftxt:
                            shared_f[si] = (ftxt, rr, cc)
                        elif si in shared_f:
                            base, br, bc = shared_f[si]
                            ftxt = F.shift_formula("=" + base, rr - br, cc - bc)[1:]
                    if ftxt:
                        raw = "=" + ftxt
                if raw is None:
                    if t == "s" and v_el is not None:
                        i = int(v_el.text)
                        raw = shared[i] if i < len(shared) else ""
                        if raw.startswith("=") or raw.startswith("'"):
                            raw = "'" + raw
                    elif t == "inlineStr":
                        raw = "".join(x.text or "" for x in c_el.iter(NS + "t"))
                    elif t == "b" and v_el is not None:
                        raw = v_el.text == "1"
                    elif t == "e" and v_el is not None:
                        raw = "=" + v_el.text if False else v_el.text
                    elif t == "str" and v_el is not None:
                        raw = v_el.text or ""
                    elif v_el is not None and v_el.text not in (None, ""):
                        try:
                            x = float(v_el.text)
                            raw = int(x) if x == int(x) and abs(x) < 1e15 and "." not in v_el.text and "E" not in v_el.text.upper() else x
                        except ValueError:
                            raw = v_el.text
                    if raw == "":
                        raw = None
                if raw is not None or sid:
                    sh.cells[(rr, cc)] = (raw, sid)
        for mc in (sroot.find(NS + "mergeCells") or []):
            a, b = mc.get("ref").split(":")
            r1, c1, *_ = F.parse_cell(a)
            r2, c2, *_ = F.parse_cell(b)
            sh.merges.append((r1, c1, r2, c2))
        af = sroot.find(NS + "autoFilter")
        if af is not None and af.get("ref") and ":" in af.get("ref"):
            a, b = af.get("ref").split(":")
            r1, c1, *_ = F.parse_cell(a)
            r2, c2, *_ = F.parse_cell(b)
            sh.filter = {"range": (r1, c1, r2, c2), "hidden": {}}
        for cf in sroot.findall(NS + "conditionalFormatting"):
            rng = cf.get("sqref", "")
            for rule in cf.findall(NS + "cfRule"):
                op = {"greaterThan": "greater than", "lessThan": "less than", "equal": "equal to", "notEqual": "not equal to", "between": "between"}.get(rule.get("operator"))
                kind = op if rule.get("type") == "cellIs" else ("text contains" if rule.get("type") == "containsText" else None)
                if kind:
                    fs = [f.text for f in rule.findall(NS + "formula")]
                    sh.cf.append({"kind": kind, "range": rng, "a": (fs[0] if fs else "") if kind != "text contains" else rule.get("text", ""), "b": fs[1] if len(fs) > 1 else "",
                                  "fg": "#9C0006", "bg": "#FFC7CE", "dxf": int(rule.get("dxfId", -1))})
        dvs = sroot.find(NS + "dataValidations")
        if dvs is not None:
            from . import validation as V
            for dv in dvs.findall(NS + "dataValidation"):
                kind = V.XL_TYPE_BACK.get(dv.get("type", ""))
                if not kind:
                    continue
                f1, f2 = dv.find(NS + "formula1"), dv.find(NS + "formula2")
                rule = V.new_rule(dv.get("sqref", ""))
                rule.update(kind=kind, op=V.XL_OP_BACK.get(dv.get("operator", "between"), "between"), blank=dv.get("allowBlank") == "1",
                            dropdown=dv.get("showDropDown") != "1", style=dv.get("errorStyle", "stop"), title=dv.get("errorTitle", ""), msg=dv.get("error", ""),
                            in_title=dv.get("promptTitle", ""), in_msg=dv.get("prompt", ""))
                if kind == "list":
                    t = (f1.text or "") if f1 is not None else ""
                    if t.startswith('"') and t.endswith('"'):
                        t = t[1:-1]
                    elif t and not t.startswith("="):
                        t = "=" + t
                    rule["list"] = t
                else:
                    rule["a"] = (f1.text or "") if f1 is not None else ""
                    rule["b"] = (f2.text or "") if f2 is not None else ""
                sh.dv.append(rule)
    if "xl/styles.xml" in names:
        dxfs = ET.fromstring(z.read("xl/styles.xml")).find(NS + "dxfs")
        if dxfs is not None:
            lst = list(dxfs)
            for sh in wb.sheets:
                for r in sh.cf:
                    i = r.pop("dxf", -1)
                    if 0 <= i < len(lst):
                        d = lst[i]
                        f = d.find(NS + "font")
                        if f is not None and f.find(NS + "color") is not None:
                            r["fg"] = _color(f.find(NS + "color")) or r["fg"]
                        fl = d.find(NS + "fill")
                        if fl is not None and fl.find(NS + "patternFill") is not None:
                            pf = fl.find(NS + "patternFill")
                            r["bg"] = _color(pf.find(NS + "bgColor")) or _color(pf.find(NS + "fgColor")) or r["bg"]
    if not wb.sheets:
        wb.sheets.append(Sheet("Sheet1"))
    wb.dirty = False
    return wb


# ---------------------------------------------------------------- writing
def _xml(s):
    return escape(s, {'"': "&quot;"})


def write_xlsx(wb, path, progress=None):
    strings, sidx = [], {}

    def sst(s):
        i = sidx.get(s)
        if i is None:
            i = sidx[s] = len(strings)
            strings.append(s)
        return i

    # styles
    fonts, fills, borders, nfs, xfs = [], [], [], {}, []
    fkey, flkey, bkey = {}, {}, {}

    def idx(table, key, val):
        if key not in table:
            table[key] = len(table)
            (fonts if table is fkey else fills if table is flkey else borders).append(val)
        return table[key]

    idx(fkey, ("Calibri", 11, 0, 0, 0, 0, None), '<font><sz val="11"/><name val="Calibri"/></font>')
    idx(flkey, None, '<fill><patternFill patternType="none"/></fill>')
    idx(flkey, "gray125", '<fill><patternFill patternType="gray125"/></fill>')
    idx(bkey, (None, None, None, None), "<border><left/><right/><top/><bottom/><diagonal/></border>")

    def argb(c):
        return "FF" + c.lstrip("#").upper()
    from .numfmt import MONTHS
    builtin_rev = {v: k for k, v in BUILTIN.items() if k not in (14, 22)}
    builtin_rev["m/d/yyyy"] = 14
    xf_of = {}

    def xf_for(sid):
        if sid in xf_of:
            return xf_of[sid]
        st = wb.styles[sid]
        fi = idx(fkey, (st.font or "Calibri", st.size or 11, st.bold, st.italic, st.underline, st.strike, st.color),
                 "<font>" + ("<b/>" if st.bold else "") + ("<i/>" if st.italic else "") + ("<strike/>" if st.strike else "") + ("<u/>" if st.underline else "")
                 + f'<sz val="{st.size or 11}"/>' + (f'<color rgb="{argb(st.color)}"/>' if st.color else "") + f'<name val="{_xml(st.font or "Calibri")}"/></font>')
        fl = idx(flkey, st.bg, f'<fill><patternFill patternType="solid"><fgColor rgb="{argb(st.bg)}"/><bgColor indexed="64"/></patternFill></fill>') if st.bg else 0
        bd = 0
        if any((st.bl, st.bt, st.br, st.bb)):
            def side(n, c):
                return f'<{n} style="thin"><color rgb="{argb(c)}"/></{n}>' if c else f"<{n}/>"
            bd = idx(bkey, (st.bl, st.bt, st.br, st.bb), "<border>" + side("left", st.bl) + side("right", st.br) + side("top", st.bt) + side("bottom", st.bb) + "<diagonal/></border>")
        nid = builtin_rev.get(st.fmt)
        if nid is None:
            if st.fmt not in nfs:
                nfs[st.fmt] = 164 + len(nfs)
            nid = nfs[st.fmt]
        al = ""
        if st.h or st.v or st.wrap or st.indent:
            al = "<alignment" + (f' horizontal="{st.h}"' if st.h else "") + (f' vertical="{ {"middle": "center"}.get(st.v, st.v) }"' if st.v else "") \
                 + (' wrapText="1"' if st.wrap else "") + (f' indent="{st.indent}"' if st.indent else "") + "/>"
        xml = f'<xf numFmtId="{nid}" fontId="{fi}" fillId="{fl}" borderId="{bd}" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"' + (' applyAlignment="1">' + al + "</xf>" if al else "/>")
        xfs.append(xml)
        xf_of[sid] = len(xfs) - 1
        return xf_of[sid]
    xf_for(0)

    dxfs, dxf_of = [], {}

    def dxf_for(rule):
        k = (rule.get("fg"), rule.get("bg"))
        if k not in dxf_of:
            dxf_of[k] = len(dxfs)
            dxfs.append(f'<dxf><font><color rgb="{argb(k[0] or "#9C0006")}"/></font><fill><patternFill><bgColor rgb="{argb(k[1] or "#FFC7CE")}"/></patternFill></fill></dxf>')
        return dxf_of[k]

    sheet_xml = []
    chart_parts = []
    for si, sh in enumerate(wb.sheets):
        out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">']
        if sh.tab_color:
            out.append(f'<sheetPr><tabColor rgb="{argb(sh.tab_color)}"/></sheetPr>')
        ur, uc = (max([r for r, _ in sh.cells] or [0]), max([c for _, c in sh.cells] or [0]))
        out.append(f'<dimension ref="A1:{F.cell_name(ur, uc)}"/>')
        view = f'<sheetViews><sheetView workbookViewId="0"' + ('' if sh.grid else ' showGridLines="0"') + (' tabSelected="1"' if si == 0 else "") + ">"
        fr, fc = sh.freeze
        if fr or fc:
            view += f'<pane' + (f' xSplit="{fc}"' if fc else "") + (f' ySplit="{fr}"' if fr else "") + f' topLeftCell="{F.cell_name(fr, fc)}" activePane="{"bottomRight" if fr and fc else "bottomLeft" if fr else "topRight"}" state="frozen"/>'
        out.append(view + "</sheetView></sheetViews>")
        out.append('<sheetFormatPr defaultRowHeight="15"/>')
        if sh.colw or sh.hidden_cols:
            out.append("<cols>")
            for c in sorted(set(sh.colw) | sh.hidden_cols):
                w = (sh.colw.get(c, DEFAULT_W) - 5) / 7
                out.append(f'<col min="{c + 1}" max="{c + 1}" width="{w:.2f}" customWidth="1"' + (' hidden="1"' if c in sh.hidden_cols else "") + "/>")
            out.append("</cols>")
        out.append("<sheetData>")
        rows = {}
        for (r, c), cell in sh.cells.items():
            rows.setdefault(r, []).append((c, cell))
        for r in sorted(set(rows) | set(sh.rowh) | sh.hidden_rows):
            attrs = f' r="{r + 1}"'
            if r in sh.rowh:
                attrs += f' ht="{sh.rowh[r] * 72 / 96:.2f}" customHeight="1"'
            if r in sh.hidden_rows:
                attrs += ' hidden="1"'
            out.append(f"<row{attrs}>")
            for c, (raw, sid) in sorted(rows.get(r, [])):
                ref = F.cell_name(r, c)
                s_attr = f' s="{xf_for(sid)}"' if sid else ""
                if raw is None:
                    out.append(f'<c r="{ref}"{s_attr}/>')
                elif isinstance(raw, bool):
                    out.append(f'<c r="{ref}"{s_attr} t="b"><v>{int(raw)}</v></c>')
                elif isinstance(raw, (int, float)):
                    out.append(f'<c r="{ref}"{s_attr}><v>{raw!r}</v></c>')
                elif raw.startswith("=") and len(raw) > 1:
                    v = wb.cell_value(sh.name, r, c)
                    f = _xml(raw[1:])
                    if isinstance(v, F.XErr):
                        out.append(f'<c r="{ref}"{s_attr} t="e"><f>{f}</f><v>{"#REF!" if v.code == "#CIRC!" else v.code}</v></c>')
                    elif isinstance(v, bool):
                        out.append(f'<c r="{ref}"{s_attr} t="b"><f>{f}</f><v>{int(v)}</v></c>')
                    elif isinstance(v, (int, float)):
                        out.append(f'<c r="{ref}"{s_attr}><f>{f}</f><v>{v!r}</v></c>')
                    elif v is None:
                        out.append(f'<c r="{ref}"{s_attr}><f>{f}</f><v>0</v></c>')
                    else:
                        out.append(f'<c r="{ref}"{s_attr} t="str"><f>{f}</f><v>{_xml(str(v))}</v></c>')
                else:
                    s = raw[1:] if raw.startswith("'") else raw
                    out.append(f'<c r="{ref}"{s_attr} t="s"><v>{sst(s)}</v></c>')
            out.append("</row>")
        out.append("</sheetData>")
        if sh.filter:
            r1, c1, r2, c2 = sh.filter["range"]
            out.append(f'<autoFilter ref="{F.cell_name(r1, c1)}:{F.cell_name(r2, c2)}"/>')
        if sh.merges:
            out.append(f'<mergeCells count="{len(sh.merges)}">' + "".join(f'<mergeCell ref="{F.cell_name(a, b)}:{F.cell_name(c, d)}"/>' for a, b, c, d in sh.merges) + "</mergeCells>")
        pr = 1
        opmap = {"greater than": "greaterThan", "less than": "lessThan", "equal to": "equal", "not equal to": "notEqual", "between": "between"}
        for rule in sh.cf:
            if rule["kind"] in opmap or rule["kind"] == "text contains":
                d = dxf_for(rule)
                if rule["kind"] == "text contains":
                    first = rule["range"].split(":")[0].split(" ")[0].replace("$", "")
                    out.append(f'<conditionalFormatting sqref="{_xml(rule["range"])}"><cfRule type="containsText" dxfId="{d}" priority="{pr}" operator="containsText" text="{_xml(rule["a"])}"><formula>NOT(ISERROR(SEARCH("{_xml(rule["a"])}",{first})))</formula></cfRule></conditionalFormatting>')
                else:
                    fs = f'<formula>{_xml(str(rule["a"]))}</formula>' + (f'<formula>{_xml(str(rule["b"]))}</formula>' if rule["kind"] == "between" else "")
                    out.append(f'<conditionalFormatting sqref="{_xml(rule["range"])}"><cfRule type="cellIs" dxfId="{d}" priority="{pr}" operator="{opmap[rule["kind"]]}">{fs}</cfRule></conditionalFormatting>')
                pr += 1
        if sh.dv:
            from . import validation as V
            items = []
            for rule in sh.dv:
                if rule["kind"] not in V.XL_TYPE:
                    continue
                at = f'type="{V.XL_TYPE[rule["kind"]]}"'
                if rule["kind"] not in ("list", "custom"):
                    at += f' operator="{V.XL_OP.get(rule.get("op", "between"), "between")}"'
                if rule.get("style", "stop") != "stop":
                    at += f' errorStyle="{rule["style"]}"'
                if rule.get("blank", True):
                    at += ' allowBlank="1"'
                if not rule.get("dropdown", True):
                    at += ' showDropDown="1"'
                if rule.get("in_title") or rule.get("in_msg"):
                    at += ' showInputMessage="1"' + (f' promptTitle="{_xml(rule["in_title"])}"' if rule.get("in_title") else "") + (f' prompt="{_xml(rule["in_msg"])}"' if rule.get("in_msg") else "")
                at += ' showErrorMessage="1"' + (f' errorTitle="{_xml(rule["title"])}"' if rule.get("title") else "") + (f' error="{_xml(rule["msg"])}"' if rule.get("msg") else "")
                if rule["kind"] == "list":
                    src = str(rule.get("list", ""))
                    f1 = _xml(src[1:] if src.startswith("=") else None) if src.startswith("=") else '"' + _xml(src) + '"'
                    f2 = ""
                elif rule["kind"] == "custom":
                    f1, f2 = _xml(str(rule.get("a", "")).lstrip("=")), ""
                else:
                    def _bound(t):
                        from .model import parse_input as _pi
                        v = _pi(str(t))[0] if rule["kind"] == "date" and not str(t).startswith("=") else None
                        return str(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else str(t).lstrip("=") if str(t).startswith("=") else str(t)
                    f1 = _xml(_bound(rule.get("a", "")))
                    f2 = f'<formula2>{_xml(_bound(rule.get("b", "")))}</formula2>' if rule.get("op") in ("between", "notBetween") else ""
                items.append(f'<dataValidation {at} sqref="{_xml(rule["range"])}"><formula1>{f1}</formula1>{f2}</dataValidation>')
            if items:
                out.append(f'<dataValidations count="{len(items)}">' + "".join(items) + "</dataValidations>")
        out.append('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>')
        pairs = [(ch, xcharts.chart_xml(wb, ch)) for ch in sh.charts]
        pairs = [(ch, cx) for ch, cx in pairs if cx]
        if pairs:
            out.append('<drawing r:id="rId1"/>')
            chart_parts.append((si, [p[0] for p in pairs], [p[1] for p in pairs]))
        out.append("</worksheet>")
        sheet_xml.append("".join(out))
        if progress:
            progress(si)

    styles = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">']
    if nfs:
        styles.append(f'<numFmts count="{len(nfs)}">' + "".join(f'<numFmt numFmtId="{i}" formatCode="{_xml(c)}"/>' for c, i in nfs.items()) + "</numFmts>")
    styles.append(f'<fonts count="{len(fonts)}">{"".join(fonts)}</fonts><fills count="{len(fills)}">{"".join(fills)}</fills><borders count="{len(borders)}">{"".join(borders)}</borders>')
    styles.append('<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>')
    styles.append(f'<cellXfs count="{len(xfs)}">{"".join(xfs)}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>')
    styles.append(f'<dxfs count="{len(dxfs)}">{"".join(dxfs)}</dxfs></styleSheet>')

    names_xml = ""
    if wb.names:
        names_xml = "<definedNames>" + "".join(f'<definedName name="{_xml(n)}">{_xml(t)}</definedName>' for n, t in wb.names.items()) + "</definedNames>"
    wbx = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets>'
           + "".join(f'<sheet name="{_xml(s.name)}" sheetId="{i + 1}" r:id="rId{i + 1}"/>' for i, s in enumerate(wb.sheets)) + "</sheets>" + names_xml + '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>')
    n = len(wb.sheets)
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + "".join(f'<Relationship Id="rId{i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i + 1}.xml"/>' for i in range(n))
            + f'<Relationship Id="rId{n + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            + f'<Relationship Id="rId{n + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>')
    chart_ct, nchart = "", 0
    for k, (si, chs, parts) in enumerate(chart_parts):
        chart_ct += f'<Override PartName="/xl/drawings/drawing{k + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
        for _ in parts:
            nchart += 1
            chart_ct += f'<Override PartName="/xl/charts/chart{nchart}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
          + "".join(f'<Override PartName="/xl/worksheets/sheet{i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for i in range(n))
          + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
          + chart_ct + "</Types>")
    sst_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="%d" uniqueCount="%d">' % (len(strings), len(strings))
               + "".join(f'<si><t xml:space="preserve">{_xml(s)}</t></si>' for s in strings) + "</sst>")
    tmp = path + ".part"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        z.writestr("xl/workbook.xml", wbx)
        z.writestr("xl/_rels/workbook.xml.rels", rels)
        z.writestr("xl/styles.xml", "".join(styles))
        z.writestr("xl/sharedStrings.xml", sst_xml)
        for i, x in enumerate(sheet_xml):
            z.writestr(f"xl/worksheets/sheet{i + 1}.xml", x)
        REL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        nchart = 0
        for k, (si, chs, parts) in enumerate(chart_parts):
            z.writestr(f"xl/worksheets/_rels/sheet{si + 1}.xml.rels", REL + f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing{k + 1}.xml"/></Relationships>')
            ids, items = [], []
            for ch, cx in zip(chs, parts):
                nchart += 1
                ids.append(nchart)
                items.append((f"rId{len(ids)}", ch))
                z.writestr(f"xl/charts/chart{nchart}.xml", cx)
            z.writestr(f"xl/drawings/drawing{k + 1}.xml", xcharts.drawing_xml(items))
            z.writestr(f"xl/drawings/_rels/drawing{k + 1}.xml.rels", REL + "".join(f'<Relationship Id="rId{j + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart{n}.xml"/>' for j, n in enumerate(ids)) + "</Relationships>")
    import os
    os.replace(tmp, path)


# ---------------------------------------------------------------- csv
def read_csv(path):
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "utf-16", "cp1252", "latin-1"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeError:
            continue
    try:
        dialect = csv.Sniffer().sniff(text[:20000], delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel
    wb = Workbook()
    sh = wb.sheets[0]
    import os
    sh.name = os.path.splitext(os.path.basename(path))[0][:31] or "Sheet1"
    for r, row in enumerate(csv.reader(io.StringIO(text, newline=""), dialect)):
        for c, val in enumerate(row):
            if val == "":
                continue
            v, fmt = parse_input(val)
            if isinstance(v, str) and v.startswith("="):
                v = "'" + v
            st = wb.style_id(DEFAULT._replace(fmt=fmt)) if fmt else 0
            sh.cells[(r, c)] = (v, st)
    wb.dirty = False
    return wb


def write_csv(wb, path, sheet=None, delimiter=","):
    sh = wb.sheet(sheet) if sheet else wb.sheets[0]
    ur, uc = wb.used_extent(sh.name)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=delimiter)
        for r in range(ur + 1):
            row = []
            for c in range(uc + 1):
                text, v, st = wb.display(sh, r, c)
                row.append(text if not isinstance(v, (int, float)) or isinstance(v, bool) or st.fmt != "General" else F_general(v))
            w.writerow(row)


def F_general(v):
    from .numfmt import general
    return general(v) if isinstance(v, float) else str(v)
