"""Digonto Sheets' own file (.dsht): a small zip with the workbook (xlsx) and the chart settings, so everything comes back exactly."""
import json
import os
import tempfile
import zipfile

from . import xlsxio

MAGIC = "digonto-sheets"


def write_dsht(wb, path, progress=None):
    with tempfile.TemporaryDirectory() as td:
        tmp = os.path.join(td, "workbook.xlsx")
        xlsxio.write_xlsx(wb, tmp, progress)
        charts = {s.name: s.charts for s in wb.sheets if s.charts}
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("manifest.json", json.dumps({"app": MAGIC, "format": 1, "content": "workbook.xlsx"}))
            z.write(tmp, "workbook.xlsx")
            z.writestr("charts.json", json.dumps(charts))


def read_dsht(path):
    with zipfile.ZipFile(path) as z:
        try:
            man = json.loads(z.read("manifest.json"))
        except KeyError:
            raise ValueError("this is not a Digonto Sheets file")
        if man.get("app") != MAGIC:
            raise ValueError("this is not a Digonto Sheets file")
        with tempfile.TemporaryDirectory() as td:
            tmp = os.path.join(td, "workbook.xlsx")
            with open(tmp, "wb") as f:
                f.write(z.read(man.get("content", "workbook.xlsx")))
            wb = xlsxio.read_xlsx(tmp)
        try:
            charts = json.loads(z.read("charts.json"))
        except KeyError:
            charts = {}
    for s in wb.sheets:
        s.charts = [dict(c) for c in charts.get(s.name, [])]
    return wb
