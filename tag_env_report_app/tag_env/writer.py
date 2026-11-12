"""Fill the reference Tag & Env Stock workbook with freshly calculated data.

The reference workbook (templates/Tag_Env_Stock_Template.xlsx) is opened
as-is, so sheet names/order/visibility, header rows, fonts, fills, borders,
column widths, number formats, page setup and freeze panes all come from it.
Only the data rows are rewritten; each new row copies the cell styles of the
template's own data rows, and the total rows copy the template's total row.
"""
from __future__ import annotations

import io
import re
from collections import Counter
from copy import copy
from datetime import date, datetime

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

from .engine import ReportResult, cell_denominator

MAIN_SHEET = "Env & Tag Stock"
REGION_SHEET = "Ghana & AB & AG Stock"
FIRST_DATA_ROW = 5
LAST_COL = 39  # AM


def _col(letter: str) -> int:
    return column_index_from_string(letter)


def _snapshot(cell):
    """Copy everything that defines a cell's look."""
    return {
        "font": copy(cell.font), "fill": copy(cell.fill), "border": copy(cell.border),
        "alignment": copy(cell.alignment), "number_format": cell.number_format,
        "protection": copy(cell.protection),
    }


def _apply(cell, snap):
    cell.font = snap["font"]
    cell.fill = snap["fill"]
    cell.border = snap["border"]
    cell.alignment = snap["alignment"]
    cell.number_format = snap["number_format"]
    cell.protection = snap["protection"]


def _style_key(cell):
    return tuple(cell._style) if cell.has_style else ()  # font/fill/border/alignment/numfmt ids


def _unmerge_from(ws, first_row: int):
    for rng in list(ws.merged_cells.ranges):
        if rng.min_row >= first_row:
            ws.unmerge_cells(str(rng))


def _strip_external_formulas(wb, cached_wb):
    """Replace formulas that point to other workbooks ('[1]Sheet'!A1) by their
    last cached value, so the file opens without an 'update links' prompt."""
    for ws in wb.worksheets:
        cws = cached_wb[ws.title]
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("=") and re.search(r"\[\d+\]", c.value):
                    c.value = cws[c.coordinate].value


# --------------------------------------------------------------------------- #
# Main sheet
# --------------------------------------------------------------------------- #
def _fill_main(ws, result: ReportResult, cfg: dict):
    # --- capture prototype styles from the template ----------------------- #
    last_tpl = ws.max_row
    total_tpl_row = last_tpl
    for r in range(ws.max_row, FIRST_DATA_ROW, -1):
        v = ws.cell(r, 2).value
        if isinstance(v, str) and v.upper().startswith("=SUM("):
            total_tpl_row = r
            break
    data_rows = range(FIRST_DATA_ROW, total_tpl_row)

    row_style = {}
    for c in range(1, LAST_COL + 1):
        # the most common style of this column across the template data rows
        cells = [ws.cell(r, c) for r in data_rows if not isinstance(ws.cell(r, c), openpyxl.cell.cell.MergedCell)]
        best = Counter(_style_key(x) for x in cells).most_common(1)[0][0] if cells else None
        proto = next((x for x in cells if _style_key(x) == best), ws.cell(FIRST_DATA_ROW, c))
        row_style[c] = _snapshot(proto)
    # Total / remark columns: use the bordered style of the first data row
    # (the template fills them only on its first rows).
    for letter in (cfg["row_total"]["col"], cfg["remarks_col"]):
        row_style[_col(letter)] = _snapshot(ws.cell(FIRST_DATA_ROW, _col(letter)))
    total_style = {c: _snapshot(ws.cell(total_tpl_row, c)) for c in range(1, LAST_COL + 1)}
    row_height = ws.row_dimensions[FIRST_DATA_ROW + 1].height or 15.0

    # --- clear the old data area ------------------------------------------ #
    _unmerge_from(ws, FIRST_DATA_ROW)
    ws.delete_rows(FIRST_DATA_ROW, ws.max_row - FIRST_DATA_ROW + 1)

    # --- header: date + conversion factors -------------------------------- #
    if result.report_date:
        ws["A1"].value = datetime.combine(result.report_date, datetime.min.time())
    factors = result.factors
    for block in cfg["blocks"]:
        for inp in block["inputs"]:
            f = factors.get(inp["factor_key"])
            if f is not None:
                ws[f"{inp['cases_col']}4"].value = int(f) if float(f).is_integer() else f

    # --- data rows --------------------------------------------------------- #
    r = FIRST_DATA_ROW
    for row in result.rows:
        for c in range(1, LAST_COL + 1):
            _apply(ws.cell(r, c), row_style[c])
        ws.row_dimensions[r].height = row_height
        ws.cell(r, 1).value = row.label
        for block in cfg["blocks"]:
            for inp in block["inputs"]:
                kg_col, cs_col = inp["kg_col"], inp["cases_col"]
                cell = row.cells.get(kg_col)
                ws[f"{kg_col}{r}"].value = round(cell.kg, 4) if cell else 0
                den = cell_denominator(row, kg_col)
                if den:
                    d = int(den) if float(den).is_integer() else den
                    ws[f"{cs_col}{r}"].value = f"=({kg_col}{r}*${cs_col}$4)/{d}"
                else:
                    ws[f"{cs_col}{r}"].value = None   # BOM qty unknown -> no number
            expr = block["max_formula"]
            ws[f"{block['max_col']}{r}"].value = "=" + re.sub(r"\{(\w+)\}", lambda m: f"{m.group(1)}{r}", expr)
        ws[f"{cfg['row_total']['col']}{r}"].value = "=" + cfg["row_total"]["formula"].format(r=r)
        ws[f"{cfg['remarks_col']}{r}"].value = row.remark
        r += 1

    # --- shared tags: merge KG + cases cells across the brand's rows -------- #
    inp_by_kg = {i["kg_col"]: i for b in cfg["blocks"] for i in b["inputs"]}
    for i, row in enumerate(result.rows):
        for kg_col, n in row.merge_spans.items():
            if n < 2:
                continue
            top = FIRST_DATA_ROW + i
            for col in (kg_col, inp_by_kg[kg_col]["cases_col"]):
                c = ws[f"{col}{top}"]
                c.alignment = c.alignment.copy(horizontal="center", vertical="center")
                ws.merge_cells(f"{col}{top}:{col}{top + n - 1}")

    # --- total row ---------------------------------------------------------- #
    last = r - 1
    for c in range(1, LAST_COL + 1):
        _apply(ws.cell(r, c), total_style[c])
    for block in cfg["blocks"]:
        for inp in block["inputs"]:
            k = inp["kg_col"]
            ws[f"{k}{r}"].value = f"=SUM({k}{FIRST_DATA_ROW}:{k}{last})"
    ws.row_dimensions[r].height = row_height

    # --- view: frozen header (B4) scrolled to the top ----------------------- #
    ws.freeze_panes = None
    ws.freeze_panes = "B4"
    ws.sheet_view.selection[-1].activeCell = "A5"
    ws.sheet_view.selection[-1].sqref = "A5"
    return FIRST_DATA_ROW, last, r


# --------------------------------------------------------------------------- #
# Ghana / Abidjan / Angola sheet
# --------------------------------------------------------------------------- #
def _fill_regions(ws, result: ReportResult):
    # Prototype rows from the template layout:
    #   3-column section (GHANA): title r1, headers r2-3, item r5, total r14
    #   2-column section (ABIDJAN / ANGOLA): title r16, headers r17-18, item r25, total r26
    #   grand total r29, blank r15
    def snap_row(r, n=3):
        return [(_snapshot(ws.cell(r, c)), ws.cell(r, c).value) for c in range(1, n + 1)]

    proto3 = {"title": snap_row(1), "h1": snap_row(2), "h2": snap_row(3), "item": snap_row(5),
              "total": snap_row(14), "blank": snap_row(15)}
    proto2 = {"title": snap_row(16), "h1": snap_row(17), "h2": snap_row(18), "item": snap_row(25),
              "total": snap_row(26), "blank": snap_row(21)}
    grand = snap_row(29)
    heights = {"title": ws.row_dimensions[1].height, "h1": ws.row_dimensions[2].height,
               "item": ws.row_dimensions[5].height or 15.0, "grand": ws.row_dimensions[29].height}

    for rng in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(rng))
    ws.delete_rows(1, ws.max_row)

    def put(r, proto_row, values=None, keep_text=False):
        for i, (snap, tval) in enumerate(proto_row, start=1):
            c = ws.cell(r, i)
            _apply(c, snap)
            if values is not None and i - 1 < len(values):
                c.value = values[i - 1]
            elif keep_text and isinstance(tval, str) and not tval.startswith("="):
                c.value = tval

    r = 1
    totals = []
    for sec in result.regions:
        p = proto3 if sec.has_tag_col else proto2
        ncol = 3 if sec.has_tag_col else 2
        put(r, p["title"], [sec.title, None, None])
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncol)
        if heights["title"]:
            ws.row_dimensions[r].height = heights["title"]
        put(r + 1, p["h1"], keep_text=True)
        put(r + 2, p["h2"], keep_text=True)
        ws.merge_cells(start_row=r + 1, start_column=1, end_row=r + 2, end_column=1)
        if heights["h1"]:
            ws.row_dimensions[r + 1].height = heights["h1"]
        first = r + 3
        rr = first
        for item, env_kg, tag_kg in sec.rows:
            vals = [item, round(env_kg, 4)] + ([round(tag_kg, 4)] if sec.has_tag_col else [])
            put(rr, p["item"], vals + [None] * (3 - len(vals)))
            ws.row_dimensions[rr].height = heights["item"]
            rr += 1
        last = rr - 1
        label = "Total" if sec.has_tag_col else "TOTAL"
        if last >= first:
            vals = [label, f"=SUM(B{first}:B{last})"] + ([f"=SUM(C{first}:C{last})"] if sec.has_tag_col else [None])
        else:
            vals = [label, 0] + ([0] if sec.has_tag_col else [None])
        put(rr, p["total"], vals)
        totals.append((rr, sec.has_tag_col))
        put(rr + 1, p["blank"], [None, None, None])
        r = rr + 2

    # Grand total, two blank rows below the last section (as in the template).
    r += 1
    parts = []
    for tr, has_tag in totals:
        parts.append(f"B{tr}")
        if has_tag:
            parts.append(f"C{tr}")
    put(r, grand, ["GRAND TOTAL", "=" + "+".join(parts) if parts else 0, None])
    if heights["grand"]:
        ws.row_dimensions[r].height = heights["grand"]
    return r


# --------------------------------------------------------------------------- #
def write_report(result: ReportResult, cfg: dict, template_path) -> bytes:
    if hasattr(template_path, "read"):
        raw = template_path.read()
    else:
        with open(template_path, "rb") as f:
            raw = f.read()
    wb = openpyxl.load_workbook(io.BytesIO(raw), keep_links=False)
    cached = openpyxl.load_workbook(io.BytesIO(raw), data_only=True, keep_links=False)
    _strip_external_formulas(wb, cached)

    _fill_main(wb[MAIN_SHEET], result, cfg)
    if REGION_SHEET in wb.sheetnames:
        _fill_regions(wb[REGION_SHEET], result)

    # Open on the main sheet and let Excel recalculate every formula on load.
    wb.active = wb.sheetnames.index(MAIN_SHEET)
    for ws in wb.worksheets:
        ws.sheet_view.tabSelected = ws.title == MAIN_SHEET
    wb.calculation.fullCalcOnLoad = True

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()
