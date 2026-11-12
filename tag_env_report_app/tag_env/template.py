"""Read the row list (and the divisors used per row) from the reference report.

The approved report's column A is the master row list: the generated report
keeps those rows, labels and order. New stock items that have no row there are
inserted in alphabetical position.
"""
from __future__ import annotations

import io
import re
from dataclasses import dataclass, field

import openpyxl

from .readers import norm

MAIN_SHEET = "Env & Tag Stock"


@dataclass
class MasterRow:
    label: str
    dens: dict[str, float] = field(default_factory=dict)   # kg_col -> divisor used in the reference


def read_master_rows(template, cfg: dict) -> list[MasterRow]:
    if hasattr(template, "read"):
        pos = template.tell()
        data = template.read()
        template.seek(pos)
    else:
        with open(template, "rb") as f:
            data = f.read()
    ws = openpyxl.load_workbook(io.BytesIO(data), read_only=True)[MAIN_SHEET]
    inputs = [i for b in cfg["blocks"] for i in b["inputs"]]
    rows = list(ws.iter_rows(min_row=5, values_only=False))
    out = []
    for cells in rows:
        by_col = {c.column_letter: c.value for c in cells if hasattr(c, "column_letter")}
        label = by_col.get("A")
        if label is None or not str(label).strip():
            if str(by_col.get("B") or "").upper().startswith("=SUM("):
                break           # total row
            continue
        m = MasterRow(label=str(label))
        for inp in inputs:
            f = str(by_col.get(inp["cases_col"]) or "")
            hit = re.search(r"/\s*([\d.]+)\s*$", f)
            if f.startswith("=") and hit and float(hit.group(1)) >= 1:
                m.dens[inp["kg_col"]] = float(hit.group(1))
        out.append(m)
    return out


def label_key(label: str) -> str:
    return norm(label)
