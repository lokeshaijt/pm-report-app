"""
Writes every sheet of the final report, using data/Report_Template.xlsx as the
formatting source (freeze panes, colors, column widths, conditional formatting,
sheet visibility) so visual output never drifts from the approved format -
only the data changes.

Sheet visibility is intentionally left untouched wherever this module doesn't
explicitly set it: the template's own hidden/visible state for each sheet
(confirmed 19-Sep-2026: FG & PM STOCK, PM Monthly Summary, Orders, and MOQ
hidden; everything else visible; Stock Report Summary and Pending PO hidden)
carries straight through.
"""
import copy
import re
from datetime import date, datetime
import openpyxl
import openpyxl.worksheet.table
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter

from .reference_data import DATA_DIR
from .weeks import month_label as _month_label

RED_FONT = openpyxl.styles.Font(color="FF9C0006")
RED_FILL = openpyxl.styles.PatternFill(start_color="FFFFC7CE", end_color="FFFFC7CE", fill_type="solid")


def _clear_conditional_formatting(ws):
    for cf in list(ws.conditional_formatting):
        del ws.conditional_formatting._cf_rules[cf]


def _add_negative_highlight(ws, cols, first_row, last_row):
    for c in cols:
        col_letter = get_column_letter(c)
        rule = CellIsRule(operator="lessThan", formula=["0"], stopIfTrue=False, font=RED_FONT, fill=RED_FILL)
        ws.conditional_formatting.add(f"{col_letter}{first_row}:{col_letter}{last_row}", rule)


def _style_ref(ws, row, max_col):
    ref = {}
    for c in range(1, max_col + 1):
        cell = ws.cell(row, c)
        ref[c] = {
            "font": copy.copy(cell.font), "fill": copy.copy(cell.fill),
            "border": copy.copy(cell.border), "alignment": copy.copy(cell.alignment),
            "number_format": cell.number_format,
        }
    return ref


def load_template():
    return openpyxl.load_workbook(f"{DATA_DIR}/Report_Template.xlsx")


def _true_last_row(ws, check_col, start_row):
    last = start_row
    for r in range(start_row, ws.max_row + 1):
        if ws.cell(r, check_col).value not in (None, ""):
            last = r
    return last


def _true_last_col(ws, header_row):
    last = 1
    for c in range(1, ws.max_column + 1):
        if ws.cell(header_row, c).value not in (None, ""):
            last = c
    return last


def fix_autofilter(ws, header_row, check_col):
    """Set the AutoFilter to span the full header-through-last-row,
    first-through-last-column extent, so every column gets a working dropdown."""
    last_col = _true_last_col(ws, header_row)
    last_row = _true_last_row(ws, check_col, header_row + 1)
    ws.auto_filter.ref = f"A{header_row}:{get_column_letter(last_col)}{last_row}"


# ---------------------------------------------------------------------------
def write_canpack_sheet(wb, rows):
    ws = wb["FG & PM STOCK"]
    TEMPLATE_ROW, MAX_COL = 2, 14
    OLD_MAX_ROW = ws.max_row
    # Everything maps back onto the bundled template's original 9 styled
    # columns (FG, BRAND, ORDER STATUS, TOTAL AVAIL, PLANNED ORDERS [unused],
    # CFC STOCK, CTN STOCK, CAN PACK, SHORT/EXCESS) - the Pending PO/Total/
    # Leftover columns all reuse their respective CFC or CTN stock style.
    STYLE_SRC = {1: 1, 2: 2, 3: 3, 4: 4, 5: 6, 6: 6, 7: 6, 8: 7, 9: 7, 10: 7,
                 11: 8, 12: 9, 13: 6, 14: 7}
    style = _style_ref(ws, TEMPLATE_ROW, 9)
    header_style = _style_ref(ws, 1, 9)

    headers = ["FG NAME", "BRAND", "ORDER STATUS", "TOTAL AVAILABLE SALE ORDERS BC",
               "CFC STOCK (IN CASES)", "CFC PENDING PO(IN CASES)", "TOTAL CFC(IN CASES)",
               "CTN STOCK (IN CASES)", "CTN PENDING PO(IN CASES)", "TOTAL CTN(IN CASES)",
               "CAN PACK (BASIS PM STOCK)", "SHORT / EXCESS",
               "LEFTOVER CFC (IN CASES)", "LEFTOVER CTN (IN CASES)"]
    for col, h in enumerate(headers, start=1):
        cell = ws.cell(1, col); cell.value = h
        t = header_style[STYLE_SRC[col]]
        cell.font, cell.fill, cell.border, cell.alignment = t["font"], t["fill"], t["border"], t["alignment"]

    for r in range(TEMPLATE_ROW, OLD_MAX_ROW + 1):
        for c in range(1, MAX_COL + 1):
            ws.cell(r, c).value = None

    def apply(cell, col):
        t = style[STYLE_SRC[col]]
        cell.font, cell.fill, cell.border, cell.alignment = t["font"], t["fill"], t["border"], t["alignment"]
        cell.number_format = t["number_format"]

    row_idx = TEMPLATE_ROW
    for r in rows:
        vals = [r["fg"], r["brand"], r["order_flag"], r["total_orders_bc"],
                r["cfc_stock_cases"], r["cfc_po_cases"], r["total_cfc_cases"],
                r["ctn_stock_cases"], r["ctn_po_cases"], r["total_ctn_cases"],
                r["can_pack"], r["short_excess"],
                r["leftover_cfc"] if r["leftover_cfc"] is not None else "-",
                r["leftover_ctn"] if r["leftover_ctn"] is not None else "-"]
        for col, val in enumerate(vals, start=1):
            cell = ws.cell(row_idx, col); cell.value = val; apply(cell, col)
        row_idx += 1

    new_max_row = row_idx - 1
    for r in range(new_max_row + 1, OLD_MAX_ROW + 1):
        for c in range(1, MAX_COL + 1):
            cell = ws.cell(r, c); cell.value = None
            cell.fill = openpyxl.styles.PatternFill(fill_type=None)
            cell.font, cell.border = openpyxl.styles.Font(), openpyxl.styles.Border()

    _clear_conditional_formatting(ws)
    _add_negative_highlight(ws, [12], TEMPLATE_ROW, new_max_row)
    fix_autofilter(ws, 1, 1)
    for letter in ("F", "G", "I", "J", "M", "N"):
        ws.column_dimensions[letter].width = 14


# ---------------------------------------------------------------------------
def _write_weekly_or_monthly_summary(ws, results, buckets, bucket_label_fn, nav_map,
                                      title, generated_on):
    OLD_MAX_COL = ws.max_column
    OLD_MAX_ROW = ws.max_row
    header_style = _style_ref(ws, 3, OLD_MAX_COL)
    data_style = _style_ref(ws, 4, OLD_MAX_COL)

    for rng in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(rng))
    for r in range(1, max(OLD_MAX_ROW, 5) + 1):
        for c in range(1, OLD_MAX_COL + 5):
            ws.cell(r, c).value = None

    start_col = 9
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
    tcell = ws.cell(1, 1); tcell.value = title
    tcell.font = openpyxl.styles.Font(bold=True, color="FFCC0000", size=14)
    ws.cell(2, 1).value = f"Report Generated on {generated_on}"
    ws.cell(2, 1).font = openpyxl.styles.Font(bold=True, italic=True, size=9)

    # Weeks alternate between two color pairs so adjacent weeks are easy to
    # tell apart: even weeks use blue header / maroon subheader, odd weeks
    # use green header / purple subheader (theme-based, "Darker 25%" tint,
    # matching the reference report's manually-applied palette).
    week_font = openpyxl.styles.Font(name="Calibri", size=10, bold=True, color="FFFFFFFF")
    subheader_font = openpyxl.styles.Font(name="Calibri", size=10, bold=True, color="FFFFFFFF")
    week_align = openpyxl.styles.Alignment(horizontal="center", vertical="center", wrap_text=True)
    WEEK_FILLS = [
        openpyxl.styles.PatternFill("solid", fgColor="FF2E75B6"),
        openpyxl.styles.PatternFill("solid", fgColor=openpyxl.styles.colors.Color(theme=6, tint=-0.249977111117893)),
    ]
    SUBHEADER_FILLS = [
        openpyxl.styles.PatternFill("solid", fgColor=openpyxl.styles.colors.Color(theme=5, tint=-0.249977111117893)),
        openpyxl.styles.PatternFill("solid", fgColor=openpyxl.styles.colors.Color(theme=7, tint=-0.249977111117893)),
    ]
    for i, b in enumerate(buckets):
        c0 = start_col + i * 4
        week_fill = WEEK_FILLS[i % 2]
        ws.merge_cells(start_row=2, start_column=c0, end_row=2, end_column=c0 + 3)
        cell = ws.cell(2, c0); cell.value = bucket_label_fn(b)
        cell.font, cell.fill, cell.alignment = week_font, week_fill, week_align
        for cc in range(c0, c0 + 4):
            ws.cell(2, cc).fill = week_fill

    headers = ["Nav Item Code", "Item Name", "Item Type", "ITEM CATEGORY", "Order Qunantity",
               "Current Stock", "Total Pending Orders", "Short / Excess"]
    for i, h in enumerate(headers, start=1):
        cell = ws.cell(3, i); cell.value = h
        src = header_style[i]
        cell.font, cell.fill, cell.border, cell.alignment = src["font"], src["fill"], src["border"], src["alignment"]
    subheaders = ["Opening Stock", "Pending Delivery", "Planned Consum", "Short / Excess"]
    for i in range(len(buckets)):
        c0 = start_col + i * 4
        subheader_fill = SUBHEADER_FILLS[i % 2]
        for j, sh in enumerate(subheaders):
            cell = ws.cell(3, c0 + j); cell.value = sh
            src = header_style[10 + j]
            cell.font, cell.fill = subheader_font, subheader_fill
            cell.border, cell.alignment = src["border"], src["alignment"]

    def apply_data(cell, col):
        t = data_style[col]
        cell.font, cell.fill, cell.border, cell.alignment = t["font"], t["fill"], t["border"], t["alignment"]
        cell.number_format = t["number_format"]

    items_sorted = sorted(results.values(), key=lambda v: v["name"])
    row_idx = 4
    for item in items_sorted:
        navcode = nav_map.get(item["name"].upper())
        ws.cell(row_idx, 1).value = navcode if navcode else "-"; apply_data(ws.cell(row_idx, 1), 1)
        ws.cell(row_idx, 2).value = item["name"]; apply_data(ws.cell(row_idx, 2), 2)
        ws.cell(row_idx, 3).value = item["item_type"]; apply_data(ws.cell(row_idx, 3), 3)
        ws.cell(row_idx, 4).value = item["category"]; apply_data(ws.cell(row_idx, 4), 4)
        ws.cell(row_idx, 5).value = round(item["order_qty"], 4); apply_data(ws.cell(row_idx, 5), 5)
        ws.cell(row_idx, 6).value = round(item["current_stock"], 4); apply_data(ws.cell(row_idx, 6), 6)
        ws.cell(row_idx, 7).value = round(item["total_pending_po"], 4); apply_data(ws.cell(row_idx, 7), 7)
        se_cell = ws.cell(row_idx, 8)
        se_cell.value = f"=(F{row_idx}+G{row_idx})-E{row_idx}"
        apply_data(se_cell, 8)

        col = start_col
        for i, b in enumerate(buckets):
            if i == 0:
                oc = ws.cell(row_idx, col); oc.value = round(item["opening"][b], 4); apply_data(oc, 10)
            else:
                prev_col = get_column_letter(col - 1)
                oc = ws.cell(row_idx, col); oc.value = f"=+{prev_col}{row_idx}"; apply_data(oc, 10)
            pd_cell = ws.cell(row_idx, col + 1); pd_cell.value = round(item["pending_delivery"][b], 4); apply_data(pd_cell, 11)
            pc_cell = ws.cell(row_idx, col + 2); pc_cell.value = round(item["planned_consum"][b], 4); apply_data(pc_cell, 12)
            se_cell = ws.cell(row_idx, col + 3)
            se_cell.value = f"=({get_column_letter(col)}{row_idx}+{get_column_letter(col+1)}{row_idx})-{get_column_letter(col+2)}{row_idx}"
            apply_data(se_cell, 13)
            col += 4
        row_idx += 1

    new_max_row = row_idx - 1
    new_max_col = start_col + len(buckets) * 4 - 1
    for r in range(new_max_row + 1, OLD_MAX_ROW + 1):
        for c in range(1, new_max_col + 1):
            cell = ws.cell(r, c); cell.value = None
            cell.fill = openpyxl.styles.PatternFill(fill_type=None)
            cell.font, cell.border = openpyxl.styles.Font(), openpyxl.styles.Border()

    _clear_conditional_formatting(ws)
    opening_cols = [start_col + 4 * i for i in range(len(buckets))]
    se_cols = [8] + [12 + 4 * i for i in range(len(buckets))]
    _add_negative_highlight(ws, opening_cols + se_cols, 4, new_max_row)
    fix_autofilter(ws, 3, 2)


def write_week_summary_sheet(wb, results, WEEKS, WEEK_BOUNDS, nav_map, generated_on):
    ws = wb["PM Week Wise Summary"]
    fmt = lambda d: d.strftime("%d-%b-%y")
    title = f"PM Requirements Week {WEEKS[0]}-{WEEKS[-1]}({fmt(WEEK_BOUNDS[WEEKS[0]][0])} to {fmt(WEEK_BOUNDS[WEEKS[-1]][1])})"
    label_fn = lambda w: f"Week {w}({fmt(WEEK_BOUNDS[w][0])} to {fmt(WEEK_BOUNDS[w][1])})"
    _write_weekly_or_monthly_summary(ws, results, WEEKS, label_fn, nav_map, title, generated_on)


def write_monthly_summary_sheet(wb, results, MONTHS, nav_map, generated_on):
    ws = wb["PM Monthly Summary"]
    title = f"PM Requirements (Monthly) - {_month_label(MONTHS[0])} to {_month_label(MONTHS[-1])}"
    label_fn = lambda m: _month_label(m)
    _write_weekly_or_monthly_summary(ws, results, MONTHS, label_fn, nav_map, title, generated_on)


# ---------------------------------------------------------------------------
def write_shortfall_sheet(wb, consolidated, WEEKS, WEEK_BOUNDS, nav_map):
    ws = wb["Shortfall"]
    fmt = lambda d: d.strftime("%d-%b-%y")
    OLD_MAX_ROW = ws.max_row
    OLD_MAX_COL = ws.max_column

    name_style = _style_ref(ws, 3, 1)[1]
    type_style = _style_ref(ws, 3, 3)[3]
    qty_style = _style_ref(ws, 3, 5)[5]
    dash_style = _style_ref(ws, 3, 6)[6]

    for rng in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(rng))
    for r in range(1, OLD_MAX_ROW + 1):
        for c in range(1, OLD_MAX_COL + 3):
            cell = ws.cell(r, c)
            cell.value = None
            cell.fill = openpyxl.styles.PatternFill(fill_type=None)
            cell.font, cell.border = openpyxl.styles.Font(), openpyxl.styles.Border()

    NW = len(WEEKS)
    SF_START = 7
    ISSUE_START = SF_START + NW
    ARRIVE_START = ISSUE_START + NW
    NEW_MAX_COL = ARRIVE_START + NW - 1

    # Each block gets its own solid color (theme-based, "Darker 25%" tint) so
    # the three week-blocks are distinguishable at a glance, matching the
    # reference report's palette.
    BLOCK_THEMES = {
        "sf":     {"title": "FF1F4E78", "week": openpyxl.styles.colors.Color(theme=5, tint=-0.249977111117893)},
        "issue":  {"title": "FF375623", "week": openpyxl.styles.colors.Color(theme=7, tint=-0.249977111117893)},
        "arrive": {"title": "FFC00000", "week": openpyxl.styles.colors.Color(theme=8, tint=-0.249977111117893)},
    }
    head_fill = openpyxl.styles.PatternFill("solid", fgColor="FF1F4E78")
    head_font = openpyxl.styles.Font(bold=True, color="FFFFFFFF")
    title_font = openpyxl.styles.Font(bold=True, color="FFFFFFFF")
    center = openpyxl.styles.Alignment(horizontal="center", vertical="center")

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
    ws.cell(1, 1).value = (f"PM Shortfall & PO Plan (weekly-bucketed Pending Delivery) - "
                            f"{fmt(WEEK_BOUNDS[WEEKS[0]][0])} to {fmt(WEEK_BOUNDS[WEEKS[-1]][1])}")

    def block_title(t, start, end, theme_key):
        ws.merge_cells(start_row=1, start_column=start, end_row=1, end_column=end)
        c = ws.cell(1, start); c.value = t
        c.fill = openpyxl.styles.PatternFill("solid", fgColor=BLOCK_THEMES[theme_key]["title"])
        c.font, c.alignment = title_font, center

    block_title("Shortfall Weeks", SF_START, SF_START + NW - 1, "sf")
    block_title("PO to Issue by", ISSUE_START, ISSUE_START + NW - 1, "issue")
    block_title("PO to Arrive", ARRIVE_START, ARRIVE_START + NW - 1, "arrive")

    for i, h in enumerate(["Nav Item Code", "Item Name", "Item Type", "Item Category",
                            "Shortfall Quantity", "Suggested MOQ"], start=1):
        c = ws.cell(2, i); c.value = h; c.fill, c.font, c.alignment = head_fill, head_font, center
    for i, w in enumerate(WEEKS):
        for base, theme_key in ((SF_START, "sf"), (ISSUE_START, "issue"), (ARRIVE_START, "arrive")):
            fill = openpyxl.styles.PatternFill("solid", fgColor=BLOCK_THEMES[theme_key]["week"])
            c = ws.cell(2, base + i); c.value = f"Week {w}"; c.fill, c.font, c.alignment = fill, head_font, center

    ws.column_dimensions[get_column_letter(SF_START)].width = 7.21875

    SF_COLS = {w: SF_START + i for i, w in enumerate(WEEKS)}
    ISSUE_COLS = {w: ISSUE_START + i for i, w in enumerate(WEEKS)}
    ARRIVE_COLS = {w: ARRIVE_START + i for i, w in enumerate(WEEKS)}

    def apply(cell, style):
        cell.font, cell.fill, cell.border, cell.alignment = style["font"], style["fill"], style["border"], style["alignment"]
        cell.number_format = style["number_format"]

    def set_dash(cell):
        cell.value = 0; apply(cell, dash_style)

    row_idx = 3
    for r in consolidated:
        navcode = nav_map.get(r["name"].upper())
        c0 = ws.cell(row_idx, 1); c0.value = navcode if navcode else "-"; apply(c0, name_style)
        c1 = ws.cell(row_idx, 2); c1.value = r["name"]; apply(c1, name_style)
        c2 = ws.cell(row_idx, 3); c2.value = r["item_type"]; apply(c2, type_style)
        c3 = ws.cell(row_idx, 4); c3.value = r["category"]; apply(c3, type_style)
        c4 = ws.cell(row_idx, 5); c4.value = round(r["total_shortfall"], 4); apply(c4, qty_style)
        c5 = ws.cell(row_idx, 6); c5.value = r["suggested_moq"] if r["suggested_moq"] is not None else 0
        apply(c5, dash_style)

        for w in WEEKS:
            cell = ws.cell(row_idx, SF_COLS[w])
            qty = r["bucket_shortfalls"].get(w)
            if qty is not None:
                cell.value = round(qty, 4); apply(cell, qty_style)
            else:
                set_dash(cell)

            cell = ws.cell(row_idx, ISSUE_COLS[w])
            qty = r["issue_by"].get(w)
            if qty is not None:
                cell.value = round(qty, 4); apply(cell, qty_style)
            else:
                set_dash(cell)

            cell = ws.cell(row_idx, ARRIVE_COLS[w])
            qty = r["arrive_by"].get(w)
            if qty is not None:
                cell.value = round(qty, 4); apply(cell, qty_style)
            else:
                set_dash(cell)
        row_idx += 1

    new_max_row = row_idx - 1
    for r in range(new_max_row + 1, OLD_MAX_ROW + 1):
        for c in range(1, NEW_MAX_COL + 1):
            cell = ws.cell(r, c); cell.value = None
            cell.fill = openpyxl.styles.PatternFill(fill_type=None)
            cell.font, cell.border = openpyxl.styles.Font(), openpyxl.styles.Border()

    fix_autofilter(ws, 2, 2)


# ---------------------------------------------------------------------------
def write_laminates_sheet(wb, results, MONTHS, nav_map):
    ws = wb["Laminates"]
    TEMPLATE_ROW = 4
    MAX_COL = ws.max_column
    OLD_MAX_ROW = ws.max_row
    style = _style_ref(ws, TEMPLATE_ROW, MAX_COL)

    def apply(cell, col):
        t = style[col]
        cell.fill, cell.border, cell.alignment = t["fill"], t["border"], t["alignment"]
        cell.number_format, cell.font = t["number_format"], t["font"]

    row_idx = TEMPLATE_ROW
    for itemk, v in sorted(results.items(), key=lambda kv: kv[1]["name"]):
        navcode = nav_map.get(v["name"].upper())
        vals = [navcode if navcode else "-", v["name"], v["uom"],
                round(v["order_qty"], 4), round(v["current_stock"], 4), round(v["total_pending_po"], 4),
                round(v["summary_short_excess"], 4), v["suggested_moq"] if v["suggested_moq"] is not None else 0]
        for col, val in enumerate(vals, start=1):
            cell = ws.cell(row_idx, col); cell.value = val; apply(cell, col)
        col = 9
        for i, m in enumerate(MONTHS):
            if i == 0:
                cell = ws.cell(row_idx, col); cell.value = round(v["opening"][m], 4); apply(cell, col)
            else:
                prev_col = get_column_letter(col - 1)
                cell = ws.cell(row_idx, col); cell.value = f"=+{prev_col}{row_idx}"; apply(cell, col)
            c1 = ws.cell(row_idx, col + 1); c1.value = round(v["pending_delivery"][m], 4); apply(c1, col + 1)
            c2 = ws.cell(row_idx, col + 2); c2.value = round(v["planned_consum"][m], 4); apply(c2, col + 2)
            c3 = ws.cell(row_idx, col + 3)
            c3.value = f"=({get_column_letter(col)}{row_idx}+{get_column_letter(col+1)}{row_idx})-{get_column_letter(col+2)}{row_idx}"
            apply(c3, col + 3)
            col += 4
        row_idx += 1

    new_max_row = row_idx - 1
    for r in range(new_max_row + 1, OLD_MAX_ROW + 1):
        for c in range(1, MAX_COL + 1):
            cell = ws.cell(r, c); cell.value = None
            cell.fill = openpyxl.styles.PatternFill(fill_type=None)
            cell.font, cell.border = openpyxl.styles.Font(), openpyxl.styles.Border()

    _clear_conditional_formatting(ws)
    opening_cols = [9 + 4 * i for i in range(len(MONTHS))]
    se_cols = [7] + [12 + 4 * i for i in range(len(MONTHS))]
    _add_negative_highlight(ws, opening_cols + se_cols, TEMPLATE_ROW, new_max_row)
    fix_autofilter(ws, 3, 2)


# ---------------------------------------------------------------------------
def write_possible_error_sheet(wb, rows):
    ws = wb["Possible Error"]
    OLD_MAX_ROW = ws.max_row
    OLD_MAX_COL = ws.max_column
    header_style = _style_ref(ws, 1, OLD_MAX_COL)
    data_border = copy.copy(ws.cell(2, 1).border) if OLD_MAX_ROW >= 2 else openpyxl.styles.Border()

    for r in range(1, OLD_MAX_ROW + 1):
        for c in range(1, OLD_MAX_COL + 1):
            ws.cell(r, c).value = None

    headers = ["Nav Doc No", "Buyer requested shipment date", "FG name", "Order Qty", "Issue"]
    for i, h in enumerate(headers, start=1):
        c = ws.cell(1, i, h)
        src = header_style.get(i, header_style[1])
        c.font, c.fill, c.alignment, c.border = src["font"], src["fill"], src["alignment"], src["border"]

    row_idx = 2
    for r in rows:
        ws.cell(row_idx, 1, r["nav_doc_no"]).border = data_border
        d_cell = ws.cell(row_idx, 2, r["ship_date"])
        d_cell.border = data_border
        if r["ship_date"] is not None:
            d_cell.number_format = "DD-MM-YYYY"
        ws.cell(row_idx, 3, r["fg_name"]).border = data_border
        ws.cell(row_idx, 4, round(r["order_qty"], 4)).border = data_border
        ws.cell(row_idx, 5, r["issue"]).border = data_border
        row_idx += 1

    new_max_row = max(row_idx - 1, 1)
    fix_autofilter(ws, 1, 3)
    return new_max_row - 1


# ---------------------------------------------------------------------------
_DATE_RE = re.compile(r"^(\d{1,2})-(\d{1,2})-(\d{4})$")


def _parse_date_str(s):
    m = _DATE_RE.match(s.strip())
    if not m:
        return None
    d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    try:
        return date(y, mo, d)
    except ValueError:
        return None


def write_m4_sheet(wb, sheet_name, src_file):
    """Rebuilds an Export/Domestic 'orders in M4' sheet from scratch (avoids
    leftover data from a differently-sized prior run), preserving whichever
    columns are hidden in the source file that day."""
    idx = wb.sheetnames.index(sheet_name)
    del wb[sheet_name]
    ws = wb.create_sheet(sheet_name, idx)

    src_wb_formula = openpyxl.load_workbook(src_file)
    src_file.seek(0)
    src_wb_value = openpyxl.load_workbook(src_file, data_only=True)
    src_ws_f = src_wb_formula["Order Status-By shipment Date"]
    src_ws_v = src_wb_value["Order Status-By shipment Date"]

    DATE_COL = 6
    for row in src_ws_f.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            v = cell.value
            if isinstance(v, str) and v.startswith("="):
                v = src_ws_v.cell(cell.row, cell.column).value
            if isinstance(v, str) and "#VALUE" in v:
                v = None
            tgt = ws.cell(cell.row, cell.column)
            if cell.column == DATE_COL and isinstance(v, str):
                dt = _parse_date_str(v)
                if dt is not None:
                    v = dt
            tgt.value = v
            tgt.font = copy.copy(cell.font)
            tgt.fill = copy.copy(cell.fill)
            tgt.border = copy.copy(cell.border)
            tgt.alignment = copy.copy(cell.alignment)
            tgt.number_format = "DD-MM-YYYY" if (cell.column == DATE_COL and isinstance(v, date)) else cell.number_format

    for rng in src_ws_f.merged_cells.ranges:
        ws.merge_cells(str(rng))
    for k, v in src_ws_f.column_dimensions.items():
        ws.column_dimensions[k].width = v.width
        if v.hidden:
            ws.column_dimensions[k].hidden = True

    max_row = ws.max_row
    rule = CellIsRule(operator="lessThan", formula=["$O7"], stopIfTrue=False, font=RED_FONT, fill=RED_FILL)
    ws.conditional_formatting.add(f"P7:P{max_row}", rule)
    fix_autofilter(ws, 6, 10)


def write_orders_sheet(wb, src_file, generated_on):
    ws = wb["Orders"]
    for rng in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(rng))
    max_r, max_c = ws.max_row, ws.max_column
    for r in range(1, max_r + 1):
        for c in range(1, max_c + 1):
            ws.cell(r, c).value = None

    header_fill = openpyxl.styles.PatternFill("solid", fgColor="FF1F4E78")
    header_font = openpyxl.styles.Font(name="Calibri", size=10, bold=True, color="FFFFFFFF")
    title_font = openpyxl.styles.Font(name="Calibri", size=12, bold=True, color="FFCC0000")
    thin = openpyxl.styles.Side(style="thin")
    border = openpyxl.styles.Border(left=thin, right=thin, top=thin, bottom=thin)
    center = openpyxl.styles.Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws["A1"] = f"Order Status-Shipment Date Wise On {generated_on}"
    ws["A1"].font = title_font

    headers = ["Buyer Name", "NAV Doc No", "Buyer Contact No", "Buyer Requested Shipment Date",
               "Plan To Ship", "Delay by Days", "Prod Id", "Prod Name", "Pending Ord Qty", "Pending Prod"]
    for i, h in enumerate(headers, start=1):
        c = ws.cell(3, i, h); c.font, c.fill, c.alignment, c.border = header_font, header_fill, center, border

    def fmt_date(v):
        if v is None:
            return None
        if isinstance(v, (datetime, date)):
            return v.strftime("%d-%m-%Y")
        return str(v).strip()

    wb_src = openpyxl.load_workbook(src_file, data_only=True, read_only=True)
    ws_src = wb_src["Order Status-By shipment Date"]

    row_idx = 4
    last_date = None
    for row in ws_src.iter_rows(min_row=7, values_only=True):
        ship_raw = row[5]
        if ship_raw not in (None, " "):
            last_date = ship_raw
        prod_name = row[9]
        if prod_name is None or (isinstance(prod_name, str) and prod_name.strip() == ""):
            continue
        vals = [row[2], row[3], row[4], fmt_date(last_date), row[6], row[7], row[8], row[9], row[10], row[14]]
        for i, v in enumerate(vals, start=1):
            v = None if (isinstance(v, str) and "#VALUE" in v) else v
            cell = ws.cell(row_idx, i, v)
            cell.border = border
            cell.font = openpyxl.styles.Font(name="Calibri", size=10)
        row_idx += 1

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["H"].width = 45
    fix_autofilter(ws, 3, 8)


# ---------------------------------------------------------------------------
def write_transaction_report_sheet(wb, items, by_item_week, weeks):
    """
    Optional "Wastage Summary" sheet, only written when a Transaction Report
    was uploaded. Week-wise pivot: one row per item, with a 7-column block
    per week (Consumption, Floor Wastage, Floor Wastage %, Store Wastage,
    Store Wastage %, Writeoff, Writeoff %), oldest week first, plus a
    trailing summary block (Consumed Qty, Floor Wastage %, Store Wastage %,
    Writeoff %) across the whole window. Each week's own Floor/Store % is
    that week's own Floor/Store divided by that week's own Consumption;
    Writeoff % is always /100. The summary block's Floor/Store % divide the
    4-week sums by the 4-week Consumed Qty; its Writeoff % is also /100. The
    week-block header (row 1) carries that week's own date range and
    alternates two fill colors so adjacent weeks are easy to tell apart; the
    summary block gets a third color. Every %-column turns red (font + fill)
    wherever it's strictly greater than 2%.
    """
    from openpyxl.formatting.rule import CellIsRule
    from .transaction_report import week_label

    sheet_name = "Wastage Summary"
    for old_name in ("Wastage Summary", "Transaction Report"):
        if old_name in wb.sheetnames:
            del wb[old_name]
    ws = wb.create_sheet(sheet_name)

    THIN = openpyxl.styles.Side(style="thin")
    BORDER = openpyxl.styles.Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    HEADER_FONT = openpyxl.styles.Font(name="Calibri", size=9, bold=True, color="FFFFFFFF")
    DATA_FONT = openpyxl.styles.Font(name="Calibri", size=9)
    CENTER = openpyxl.styles.Alignment(horizontal="center", vertical="center", wrap_text=True)
    BASE_FILL = openpyxl.styles.PatternFill("solid", fgColor="FF1F4E78")
    SUBHEADER_FILL = openpyxl.styles.PatternFill(
        "solid", fgColor=openpyxl.styles.colors.Color(theme=2, tint=-0.749992370372631)
    )
    WEEK_FILLS = [
        openpyxl.styles.PatternFill("solid", fgColor=openpyxl.styles.colors.Color(theme=5, tint=-0.249977111117893)),
        openpyxl.styles.PatternFill("solid", fgColor=openpyxl.styles.colors.Color(theme=9, tint=-0.249977111117893)),
    ]
    SUMMARY_TITLE_FILL = openpyxl.styles.PatternFill(
        "solid", fgColor=openpyxl.styles.colors.Color(theme=8, tint=-0.249977111117893)
    )
    SUMMARY_HEADER_FILL = openpyxl.styles.PatternFill(
        "solid", fgColor=openpyxl.styles.colors.Color(theme=5, tint=-0.249977111117893)
    )
    NUM_FMT = '_ * #,##0_ ;_ * \\-#,##0_ ;_ * "-"??_ ;_ @_ '
    PCT_FMT = "0%"
    RED_FONT = openpyxl.styles.Font(name="Calibri", size=9, color="FF9C0006")
    RED_FILL = openpyxl.styles.PatternFill("solid", fgColor="FFFFC7CE")

    HEADER_ROW = 2
    START_COL = 4  # D
    WEEK_BLOCK_WIDTH = 7
    SUMMARY_COL = START_COL + len(weeks) * WEEK_BLOCK_WIDTH

    for i, w in enumerate(weeks):
        c0 = START_COL + i * WEEK_BLOCK_WIDTH
        ws.merge_cells(start_row=1, start_column=c0, end_row=1, end_column=c0 + WEEK_BLOCK_WIDTH - 1)
        cell = ws.cell(1, c0, week_label(w))
        cell.font, cell.alignment = HEADER_FONT, CENTER
        fill = WEEK_FILLS[i % 2]
        for cc in range(c0, c0 + WEEK_BLOCK_WIDTH):
            ws.cell(1, cc).fill = fill
            ws.cell(1, cc).border = BORDER

    ws.merge_cells(start_row=1, start_column=SUMMARY_COL, end_row=1, end_column=SUMMARY_COL + 3)
    title_cell = ws.cell(1, SUMMARY_COL, f"Week {weeks[0]} - Week {weeks[-1]}")
    title_cell.font, title_cell.alignment = HEADER_FONT, CENTER
    for cc in range(SUMMARY_COL, SUMMARY_COL + 4):
        ws.cell(1, cc).fill = SUMMARY_TITLE_FILL
        ws.cell(1, cc).border = BORDER

    for col, h in enumerate(["Item Name", "UOM", "Item Type"], start=1):
        cell = ws.cell(HEADER_ROW, col, h)
        cell.font, cell.fill, cell.alignment, cell.border = HEADER_FONT, BASE_FILL, CENTER, BORDER
    WEEK_SUBHEADERS = ["Consumption", "Floor Wastage", "Floor Wastage %",
                       "Store Wastage", "Store Wastage %", "Writeoff", "Writeoff %"]
    for i, w in enumerate(weeks):
        c0 = START_COL + i * WEEK_BLOCK_WIDTH
        for j, h in enumerate(WEEK_SUBHEADERS):
            cell = ws.cell(HEADER_ROW, c0 + j, h)
            cell.font, cell.fill, cell.alignment, cell.border = HEADER_FONT, SUBHEADER_FILL, CENTER, BORDER
    for j, h in enumerate(["Consumed Qty", "Floor Wastage %", "Store Wastage %", "Writeoff %"]):
        cell = ws.cell(HEADER_ROW, SUMMARY_COL + j, h)
        cell.font, cell.fill, cell.alignment, cell.border = HEADER_FONT, SUMMARY_HEADER_FILL, CENTER, BORDER
    ws.row_dimensions[HEADER_ROW].height = 24

    pct_col_letters = []  # every %-column, for conditional formatting

    row_idx = HEADER_ROW + 1
    for item in items:
        ws.cell(row_idx, 1, item["item_name"])
        ws.cell(row_idx, 2, item["uom"])
        ws.cell(row_idx, 3, item["item_type"])
        for c in (1, 2, 3):
            cell = ws.cell(row_idx, c)
            cell.font, cell.border = DATA_FONT, BORDER

        week_cons_letters, week_floor_letters, week_store_letters, week_wo_letters = [], [], [], []
        total_cons_v = 0.0
        for i, w in enumerate(weeks):
            c0 = START_COL + i * WEEK_BLOCK_WIDTH
            v = by_item_week.get((item["item_name"], w))
            cons_v = round(v["consumption"], 4) if v and v["consumption"] else None
            total_cons_v += cons_v or 0.0
            floor_v = round(v["floor"], 4) if v and v["floor"] else None
            store_v = round(v["store"], 4) if v and v["store"] else None
            wo_v = round(v["writeoff"], 4) if v and v["writeoff"] else None

            cons_cell = ws.cell(row_idx, c0, cons_v)
            floor_cell = ws.cell(row_idx, c0 + 1, floor_v)
            # avoid #DIV/0! when there's no Consumption this week - a 0
            # denominator means 0% wastage-of-consumption, not an error
            floor_pct = ws.cell(row_idx, c0 + 2,
                                 f"={floor_cell.coordinate}/{cons_cell.coordinate}" if cons_v else "=0/100%")
            store_cell = ws.cell(row_idx, c0 + 3, store_v)
            store_pct = ws.cell(row_idx, c0 + 4,
                                 f"={store_cell.coordinate}/{cons_cell.coordinate}" if cons_v else "=0/100%")
            wo_cell = ws.cell(row_idx, c0 + 5, wo_v)
            wo_pct = ws.cell(row_idx, c0 + 6, f"={wo_cell.coordinate}/100")

            for cell in (cons_cell, floor_cell, store_cell, wo_cell):
                cell.font, cell.border, cell.number_format = DATA_FONT, BORDER, NUM_FMT
            for cell in (floor_pct, store_pct, wo_pct):
                cell.font, cell.border, cell.number_format = DATA_FONT, BORDER, PCT_FMT
            if row_idx == HEADER_ROW + 1:
                pct_col_letters += [floor_pct.column_letter, store_pct.column_letter, wo_pct.column_letter]

            week_cons_letters.append(cons_cell.coordinate)
            week_floor_letters.append(floor_cell.coordinate)
            week_store_letters.append(store_cell.coordinate)
            week_wo_letters.append(wo_cell.coordinate)

        cons_sum = "+".join(week_cons_letters)
        floor_sum = "+".join(week_floor_letters)
        store_sum = "+".join(week_store_letters)
        wo_sum = "+".join(week_wo_letters)
        total_cons_cell = ws.cell(row_idx, SUMMARY_COL, f"={cons_sum}")
        # same 0% fallback as the per-week columns: no Consumption across the
        # whole window means 0% wastage-of-consumption, not a div/0 error
        floor_pct_cell = ws.cell(row_idx, SUMMARY_COL + 1,
                                  f"=({floor_sum})/{total_cons_cell.coordinate}" if total_cons_v else "=0/100%")
        store_pct_cell = ws.cell(row_idx, SUMMARY_COL + 2,
                                  f"=({store_sum})/{total_cons_cell.coordinate}" if total_cons_v else "=0/100%")
        wo_pct_cell = ws.cell(row_idx, SUMMARY_COL + 3, f"=({wo_sum})/100")
        total_cons_cell.font, total_cons_cell.border, total_cons_cell.number_format = DATA_FONT, BORDER, NUM_FMT
        for cell in (floor_pct_cell, store_pct_cell, wo_pct_cell):
            cell.font, cell.border, cell.number_format = DATA_FONT, BORDER, PCT_FMT
        if row_idx == HEADER_ROW + 1:
            pct_col_letters += [floor_pct_cell.column_letter, store_pct_cell.column_letter, wo_pct_cell.column_letter]
        row_idx += 1

    last_row = row_idx - 1
    last_col_letter = get_column_letter(SUMMARY_COL + 3)
    ws.auto_filter.ref = f"A{HEADER_ROW}:{last_col_letter}{last_row}"

    for col_letter in pct_col_letters:
        rule = CellIsRule(operator="greaterThan", formula=["0.02"], stopIfTrue=False,
                           font=RED_FONT, fill=RED_FILL)
        ws.conditional_formatting.add(f"{col_letter}{HEADER_ROW + 1}:{col_letter}{last_row}", rule)

    ws.column_dimensions["A"].width = 55
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 14
    for col_letter in pct_col_letters:
        ws.column_dimensions[col_letter].width = 13
    ws.column_dimensions[get_column_letter(SUMMARY_COL)].width = 13
    ws.freeze_panes = f"D{HEADER_ROW + 1}"


# ---------------------------------------------------------------------------
def write_raw_hidden_sheet(wb, dest_sheet_name, src_path):
    """Adds a hidden, full raw copy of a source workbook (Stock Report Summary
    or Pending PO) - visible only if someone explicitly unhides it in Excel."""
    src_wb = openpyxl.load_workbook(src_path, data_only=True)
    created = []
    for sn in src_wb.sheetnames:
        src_ws = src_wb[sn]
        tab_name = dest_sheet_name if len(src_wb.sheetnames) == 1 else f"{dest_sheet_name} ({sn})"
        tab_name = tab_name[:31]
        if tab_name in wb.sheetnames:
            del wb[tab_name]
        ws = wb.create_sheet(tab_name)
        for row in src_ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                tgt = ws.cell(cell.row, cell.column)
                tgt.value = cell.value
                tgt.font = copy.copy(cell.font)
                tgt.fill = copy.copy(cell.fill)
                tgt.border = copy.copy(cell.border)
                tgt.alignment = copy.copy(cell.alignment)
                tgt.number_format = cell.number_format
        for rng in src_ws.merged_cells.ranges:
            ws.merge_cells(str(rng))
        for k, v in src_ws.column_dimensions.items():
            ws.column_dimensions[k].width = v.width
        ws.sheet_state = "hidden"
        created.append(tab_name)
    return created
