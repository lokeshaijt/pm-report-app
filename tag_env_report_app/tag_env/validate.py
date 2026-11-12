"""Compare a generated report with the reference template (structure, format, formulas)."""
from __future__ import annotations

import io
import re

import openpyxl
import pandas as pd
from openpyxl.cell.cell import MergedCell

from .engine import ReportResult

MAIN = "Env & Tag Stock"
REGION = "Ghana & AB & AG Stock"


def _sk(c):
    """Style signature used for comparisons."""
    f = c.font
    return (f.name, f.sz, f.b, f.i, str(f.color.rgb if f.color is not None else None),
            c.fill.fill_type, str(c.fill.fgColor.rgb), c.fill.fgColor.theme, round(c.fill.fgColor.tint or 0, 3),
            c.border.left.style, c.border.right.style, c.border.top.style, c.border.bottom.style,
            c.alignment.horizontal, c.alignment.vertical, c.number_format)


def validate_report(xlsx: bytes, template_path, result: ReportResult, cfg: dict) -> pd.DataFrame:
    out = openpyxl.load_workbook(io.BytesIO(xlsx))
    tpl = openpyxl.load_workbook(template_path)
    checks = []

    def add(name, ok, detail=""):
        checks.append({"Check": name, "Result": "PASS" if ok is True else ("INFO" if ok is None else "FAIL"),
                       "Details": detail})

    # 1. sheets
    add("Worksheet names & order", out.sheetnames == tpl.sheetnames, ", ".join(out.sheetnames))
    add("Worksheet visibility", [w.sheet_state for w in out.worksheets] == [w.sheet_state for w in tpl.worksheets],
        ", ".join(f"{w.title}: {w.sheet_state}" for w in out.worksheets))

    # 2. column widths on every sheet
    bad = []
    for ws in tpl.worksheets:
        o = out[ws.title]
        for k, d in ws.column_dimensions.items():
            if abs((o.column_dimensions[k].width or 0) - (d.width or 0)) > 0.01:
                bad.append(f"{ws.title}!{k}")
    add("Column widths", not bad, "identical" if not bad else "differs: " + ", ".join(bad[:10]))

    o, t = out[MAIN], tpl[MAIN]
    factor_cells = {f"{i['cases_col']}4" for b in cfg["blocks"] for i in b["inputs"]}

    # 3. header rows 1-4 (text + styles)
    diff_val, diff_sty = [], []
    for r in range(1, 5):
        for c in range(1, 40):
            tc, oc = t.cell(r, c), o.cell(r, c)
            if isinstance(tc, MergedCell):
                continue
            if tc.coordinate not in factor_cells | {"A1"} and tc.value != oc.value:
                diff_val.append(tc.coordinate)
            if _sk(tc) != _sk(oc):
                diff_sty.append(tc.coordinate)
    add("Header titles / column names (rows 2-3)", not diff_val, "identical" if not diff_val else ", ".join(diff_val))
    add("Header formatting (fonts, fills, borders, alignment)", not diff_sty,
        "identical" if not diff_sty else ", ".join(diff_sty))
    hm_t = sorted(str(m) for m in t.merged_cells.ranges if m.max_row <= 4)
    hm_o = sorted(str(m) for m in o.merged_cells.ranges if m.max_row <= 4)
    add("Header merged cells", hm_t == hm_o, ", ".join(hm_o))
    add("Row heights (header)", all(t.row_dimensions[r].height == o.row_dimensions[r].height for r in range(1, 5)), "")

    # 4. conversion factors in row 4
    fac = {f"{i['cases_col']}4": cfg["conversion_factors"][i["factor_key"]] for b in cfg["blocks"] for i in b["inputs"]}
    wrong = [f"{k}={o[k].value} (expected {v})" for k, v in fac.items() if o[k].value != v]
    add("Conversion factors in row 4 match config", not wrong,
        "; ".join(f"{k}={v}" for k, v in fac.items()) if not wrong else "; ".join(wrong))

    # 5. data rows: styles must be ones the template uses in the same column
    first, n = 5, len(result.rows)
    last = first + n - 1
    t_total = next(r for r in range(t.max_row, 4, -1) if str(t.cell(r, 2).value or "").upper().startswith("=SUM("))
    allowed = {c: {_sk(t.cell(r, c)) for r in range(5, t_total) if not isinstance(t.cell(r, c), MergedCell)}
               for c in range(1, 40)}
    bad = [o.cell(r, c).coordinate for r in range(first, last + 1) for c in range(1, 40)
           if _sk(o.cell(r, c)) not in allowed[c]]
    add("Data-row formatting copied from template rows", not bad,
        f"{n} rows checked" if not bad else f"{len(bad)} cells differ, e.g. {', '.join(bad[:8])}")
    bad = [o.cell(last + 1, c).coordinate for c in range(1, 40) if _sk(o.cell(last + 1, c)) != _sk(t.cell(t_total, c))]
    add("Total-row formatting", not bad, f"row {last + 1}" if not bad else ", ".join(bad))

    # 6. formulas
    kg_cols = [(i["kg_col"], i["cases_col"]) for b in cfg["blocks"] for i in b["inputs"]]
    bad, dens, blanks = [], set(), 0
    for r in range(first, last + 1):
        for kg, cs in kg_cols:
            v = o[f"{cs}{r}"].value
            if v is None:
                blanks += 1 if (o[f"{kg}{r}"].value or 0) else 0
                continue
            m = re.fullmatch(rf"=\({kg}{r}\*\${cs}\$4\)/([\d.]+)", str(v))
            if not m or float(m.group(1)) < 1:
                bad.append(f"{cs}{r}")
            else:
                dens.add(float(m.group(1)))
    add("Make Cases formulas =(KG x factor)/BOM qty", not bad,
        f"{len(dens)} distinct BOM denominators used: {', '.join(f'{d:g}' for d in sorted(dens))}"
        if not bad else ", ".join(bad[:10]))
    add("Cells with stock KG but no valid BOM qty (left blank, see Exceptions)", None, f"{blanks} cells")
    sums_ok = all(o[f"{kg}{last + 1}"].value == f"=SUM({kg}{first}:{kg}{last})" for kg, _ in kg_cols)
    add("Total row SUM covers every data row", sums_ok, f"rows {first}-{last}")
    ext = [f"{ws.title}!{c.coordinate}" for ws in out.worksheets for row in ws.iter_rows() for c in row
           if isinstance(c.value, str) and re.search(r"\[\d+\]", c.value)]
    add("No links to other workbooks", not ext and not out._external_links, ", ".join(ext[:5]))

    # 7. view / print settings
    add("Freeze panes", o.freeze_panes == "B4", f"{o.freeze_panes}")
    ps_t, ps_o = t.page_setup, o.page_setup
    add("Page setup (orientation / paper / scale)",
        (ps_t.orientation, ps_t.paperSize, ps_t.scale) == (ps_o.orientation, ps_o.paperSize, ps_o.scale),
        f"{ps_o.orientation}, paper {ps_o.paperSize}, scale {ps_o.scale}")

    # 8. KG reconciliation: KG written to the workbook == KG of the stock items in scope
    main_kg = sum(o[f"{kg}{r}"].value or 0 for r in range(first, last + 1) for kg, _ in kg_cols)
    add("KG reconciliation - main sheet vs Stock Summary", abs(main_kg - result.stock_kg_main) < 0.01,
        f"report {main_kg:,.2f} kg / stock {result.stock_kg_main:,.2f} kg")
    if REGION in out.sheetnames:
        ro = out[REGION]
        reg_kg = sum(c.value for row in ro.iter_rows(min_col=2, max_col=3) for c in row
                     if isinstance(c.value, (int, float)))
        add("KG reconciliation - regional sheet vs Stock Summary", abs(reg_kg - result.stock_kg_regions) < 0.01,
            f"report {reg_kg:,.2f} kg / stock {result.stock_kg_regions:,.2f} kg")

    # 9. regional sheet headers
    if REGION in out.sheetnames:
        ro, rt = out[REGION], tpl[REGION]
        ok = all(_sk(ro.cell(r, c)) == _sk(rt.cell(r, c)) and ro.cell(r, c).value == rt.cell(r, c).value
                 for r in (2, 3) for c in (1, 2, 3))
        add("Ghana / Abidjan / Angola sheet headers & formatting", ok, "")
    return pd.DataFrame(checks)
