"""Write the floating charts of a sheet as real Excel charts (drawing + chart parts of the .xlsx package)."""
from xml.sax.saxutils import escape as _e

from . import formula as F
from .charts import PALETTE, read_data

NS = ('xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
EMU = 9525


def _fill(col):
    return f'<c:spPr><a:solidFill><a:srgbClr val="{col.lstrip("#").upper()}"/></a:solidFill></c:spPr>'


def _line(col, w=28575):
    return f'<c:spPr><a:ln w="{w}" cap="rnd"><a:solidFill><a:srgbClr val="{col.lstrip("#").upper()}"/></a:solidFill><a:round/></a:ln></c:spPr>'


def _str_cache(vals):
    return f'<c:strCache><c:ptCount val="{len(vals)}"/>' + "".join(f'<c:pt idx="{i}"><c:v>{_e(str(v))}</c:v></c:pt>' for i, v in enumerate(vals)) + "</c:strCache>"


def _num_cache(vals):
    return (f'<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="{len(vals)}"/>'
            + "".join(f'<c:pt idx="{i}"><c:v>{v!r}</c:v></c:pt>' for i, v in enumerate(vals) if isinstance(v, (int, float))) + "</c:numCache>")


def _title(text):
    if not text:
        return '<c:autoTitleDeleted val="1"/>'
    return (f'<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:pPr><a:defRPr sz="1400" b="1"/></a:pPr><a:r><a:rPr lang="en-US" sz="1400" b="1"/><a:t>{_e(text)}</a:t></a:r></a:p></c:rich></c:tx>'
            '<c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>')


def chart_xml(wb, ch):
    refs = {}
    cats, series = read_data(wb, ch, refs)
    if not series:
        return None
    kind = ch.get("kind", "column")
    rs = refs.get("series", [])
    cat_ref = refs.get("cat")

    def tx(i, name):
        r = rs[i]["name"] if i < len(rs) else None
        if r:
            return f'<c:tx><c:strRef><c:f>{_e(r)}</c:f>{_str_cache([name])}</c:strRef></c:tx>'
        return f'<c:tx><c:v>{_e(name)}</c:v></c:tx>'

    def cat(extra=""):
        if cat_ref:
            return f'<c:cat><c:strRef><c:f>{_e(cat_ref)}</c:f>{_str_cache(cats)}</c:strRef></c:cat>'
        return f'<c:cat><c:strLit><c:ptCount val="{len(cats)}"/>' + "".join(f'<c:pt idx="{i}"><c:v>{_e(str(v))}</c:v></c:pt>' for i, v in enumerate(cats)) + "</c:strLit></c:cat>"

    def val(i, vals):
        return f'<c:val><c:numRef><c:f>{_e(rs[i]["val"])}</c:f>{_num_cache(vals)}</c:numRef></c:val>'

    stacked = bool(ch.get("stacked")) and kind in ("column", "bar", "area")
    sers = []
    for i, (name, vals) in enumerate(series):
        col = PALETTE[i % len(PALETTE)]
        head = f'<c:idx val="{i}"/><c:order val="{i}"/>{tx(i, name)}'
        if kind in ("column", "bar"):
            sers.append(f'<c:ser>{head}{_fill(col)}<c:invertIfNegative val="0"/>{cat()}{val(i, vals)}</c:ser>')
        elif kind == "line":
            sers.append(f'<c:ser>{head}{_line(col)}<c:marker><c:symbol val="circle"/><c:size val="5"/>{_fill(col)}</c:marker>{cat()}{val(i, vals)}<c:smooth val="0"/></c:ser>')
        elif kind == "area":
            sers.append(f'<c:ser>{head}{_fill(col)}{cat()}{val(i, vals)}</c:ser>')
        elif kind == "scatter":
            xs = [float(c) if str(c).replace(".", "", 1).replace("-", "", 1).isdigit() else k + 1 for k, c in enumerate(cats)]
            xref = f'<c:xVal><c:numRef><c:f>{_e(cat_ref)}</c:f>{_num_cache(xs)}</c:numRef></c:xVal>' if cat_ref else f'<c:xVal><c:numLit><c:ptCount val="{len(xs)}"/>' + "".join(f'<c:pt idx="{k}"><c:v>{v!r}</c:v></c:pt>' for k, v in enumerate(xs)) + "</c:numLit></c:xVal>"
            sers.append(f'<c:ser>{head}<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="7"/>{_fill(col)}</c:marker>{xref}'
                        f'<c:yVal><c:numRef><c:f>{_e(rs[i]["val"])}</c:f>{_num_cache(vals)}</c:numRef></c:yVal><c:smooth val="0"/></c:ser>')
    ax = '<c:axId val="111"/><c:axId val="222"/>'
    cat_ax = '<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="%s"/><c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="222"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx>'
    val_ax = '<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="%s"/><c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="E0E0E0"/></a:solidFill></a:ln></c:spPr></c:majorGridlines><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="111"/><c:crosses val="autoZero"/><c:crossBetween val="%s"/></c:valAx>'
    if kind in ("column", "bar"):
        horiz = kind == "bar"
        plot = (f'<c:barChart><c:barDir val="{"bar" if horiz else "col"}"/><c:grouping val="{"stacked" if stacked else "clustered"}"/><c:varyColors val="0"/>{"".join(sers)}'
                f'<c:gapWidth val="80"/>{"<c:overlap val=\"100\"/>" if stacked else ""}{ax}</c:barChart>'
                + cat_ax % ("l" if horiz else "b") + val_ax % ("b" if horiz else "l", "between"))
    elif kind == "line":
        plot = f'<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>{"".join(sers)}<c:marker val="1"/>{ax}</c:lineChart>' + cat_ax % "b" + val_ax % ("l", "between")
    elif kind == "area":
        plot = f'<c:areaChart><c:grouping val="{"stacked" if stacked else "standard"}"/><c:varyColors val="0"/>{"".join(sers)}{ax}</c:areaChart>' + cat_ax % "b" + val_ax % ("l", "midCat")
    elif kind == "scatter":
        plot = f'<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>{"".join(sers)}{ax}</c:scatterChart>' + val_ax.replace('<c:axId val="222"/>', '<c:axId val="222"/>', 1) % ("l", "midCat")
        plot = plot.replace(f'<c:valAx><c:axId val="222"/>', '<c:valAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="222"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx><c:valAx><c:axId val="222"/>', 1)
    else:  # pie
        name, vals = series[0]
        dpts = "".join(f'<c:dPt><c:idx val="{i}"/><c:bubble3D val="0"/>{_fill(PALETTE[i % len(PALETTE)])}</c:dPt>' for i in range(len(vals)))
        plot = (f'<c:pieChart><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/>{tx(0, name)}{dpts}'
                f'<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="1"/><c:showBubbleSize val="0"/></c:dLbls>{cat()}{val(0, vals)}</c:ser><c:firstSliceAng val="0"/></c:pieChart>')
    legend = '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>' if ch.get("legend", True) and (len(series) > 1 or kind == "pie") else ""
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace {NS}><c:roundedCorners val="0"/><c:chart>{_title(ch.get("title", ""))}'
            f'<c:plotArea><c:layout/>{plot}</c:plotArea>{legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>')


def drawing_xml(items):
    """items: [(rId, chart dict)] -> drawing part with the same position and size as on screen"""
    out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
           'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
           'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">']
    for i, (rid, ch) in enumerate(items):
        x, y, w, h = (int(ch.get(k, d)) * EMU for k, d in (("x", 20), ("y", 20), ("w", 440), ("h", 280)))
        out.append(f'<xdr:absoluteAnchor><xdr:pos x="{x}" y="{y}"/><xdr:ext cx="{w}" cy="{h}"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="{i + 2}" name="Chart {i + 1}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>'
                   f'<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="{rid}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:absoluteAnchor>')
    out.append("</xdr:wsDr>")
    return "".join(out)
