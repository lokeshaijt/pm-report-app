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
    prev_remarks: set[str] = field(default_factory=set)    # remarks this row had in the reference


def read_master_rows(template, cfg: dict) -> list[MasterRow]:
    if hasattr(template, "read"):
        pos = template.tell()
        data = template.read()
        template.seek(pos)
    else:
        with open(template, "rb") as f:
            data = f.read()
    prev = previous_remarks(openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True), cfg)
    ws = openpyxl.load_workbook(io.BytesIO(data), read_only=True)[MAIN_SHEET]   # formulas (divisors)
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
        m = MasterRow(label=str(label), prev_remarks=prev.get(norm(label), set()))
        for inp in inputs:
            f = str(by_col.get(inp["cases_col"]) or "")
            hit = re.search(r"/\s*([\d.]+)\s*$", f)
            if f.startswith("=") and hit and float(hit.group(1)) >= 1:
                m.dens[inp["kg_col"]] = float(hit.group(1))
        out.append(m)
    return out


def previous_remarks(wb, cfg: dict) -> dict[str, set[str]]:
    """Remarks per row label in the reference workbook.

    Read from every sheet that has a 'Remarks' header (the visible sheet's AM
    column and the 'Tag & Env Stock' sheets' column S), so a brand marked
    'Running' last time can be carried forward.
    """
    out: dict[str, set[str]] = {}
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(min_row=1, max_row=min(ws.max_row or 0, 2000), values_only=True))
        col = hdr = None
        for i, r in enumerate(rows[:5]):
            for j, v in enumerate(r):
                if str(v or "").strip().lower() == "remarks":
                    col, hdr = j, i
        if col is None:
            continue
        for r in rows[hdr + 1:]:
            if not r or r[0] is None or col >= len(r) or not str(r[col] or "").strip():
                continue
            out.setdefault(norm(r[0]), set()).add(str(r[col]).strip())
    return out


def label_key(label: str) -> str:
    return norm(label)
