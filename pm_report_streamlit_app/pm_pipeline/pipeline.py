"""
Ties every module together: parse the 4 required files (plus an optional
supplemental BOM), combine with the bundled reference data, run the weekly +
monthly cascades, and write the final formatted workbook.
"""
import io
from datetime import date

from . import reference_data as ref
from .stock import parse_stock
from .pending_po import parse_pending_po
from .orders import combine_demand
from .pm_universe import build_pm_universe
from .possible_error import build_possible_error
from .demand import aggregate_weekly, aggregate_monthly
from .cascade import explode_demand, run_cascade, consolidate_shortfall
from .canpack import build_canpack
from .weeks import build_week_window, build_month_window
from . import report_writer as rw


def generate_report(export_file, domestic_file, pending_po_file, stock_file,
                     supplemental_bom_file=None, today: date = None, progress_cb=None):
    """
    Each *_file argument is a file-like object (e.g. from st.file_uploader).
    supplemental_bom_file is optional: a one-off BOM (same 6-column layout as
    the bundled Exploded_BOM_Supplemental.xlsx) covering FGs still showing on
    the Possible Error sheet. It is merged in for this run only - it is not
    saved into the bundled reference data.
    progress_cb(pct: float, message: str) is called periodically if provided.
    Returns: (bytes of the finished .xlsx, a dict of summary stats for the UI)
    """
    def prog(pct, msg):
        if progress_cb:
            progress_cb(pct, msg)

    if today is None:
        today = date.today()
    generated_on = today.strftime("%d/%m/%Y")

    # ---- 1. Reference data (bundled, static) ----
    prog(0.05, "Loading reference data (BOM, Can Pack master, Nav codes)...")
    raw_bom_rows, item_meta, fg_brand, uom_map, fg_cfc, fg_ctn = ref.load_bom()
    supp_bundled_rows, supp_bundled_brand = ref.load_supplemental_bom(
        f"{ref.DATA_DIR}/Exploded_BOM_Supplemental.xlsx"
    )
    fg_brand.update(supp_bundled_brand)
    canpack_master = ref.load_canpack_master()
    nav_map = ref.load_nav_map()

    supplemental_lists = [supp_bundled_rows]
    if supplemental_bom_file is not None:
        prog(0.10, "Merging uploaded supplemental BOM...")
        extra_rows, extra_brand = ref.load_supplemental_bom(supplemental_bom_file)
        fg_brand.update(extra_brand)
        supplemental_lists.append(extra_rows)

    # ---- 2. Parse the 4 uploaded files ----
    prog(0.15, "Parsing Stock Report...")
    stock = parse_stock(stock_file)

    prog(0.20, "Parsing Pending PO...")
    total_pending_po, po_rows, po_format, po_excluded = parse_pending_po(pending_po_file)

    prog(0.25, "Parsing Export and Domestic order status...")
    export_file.seek(0); domestic_file.seek(0)
    demand_rows, canpack_rows, possible_error_rows, order_active_fgs = combine_demand(export_file, domestic_file)

    # ---- 3. PM item universe: every FG in the merged BOM (no Active FG list) ----
    prog(0.35, "Building PM item universe...")
    pm_items, fg_names_in_bom = build_pm_universe(raw_bom_rows, item_meta, supplemental_lists)

    # ---- 4. Possible Error: order-active FGs still missing from the BOM ----
    possible_error = build_possible_error(possible_error_rows, fg_names_in_bom)

    # ---- 5. Weekly cascade (8-week rolling window, starting today) ----
    prog(0.45, "Running the 8-week cascade...")
    WEEKS, WEEK_BOUNDS, RANGE_START, RANGE_END = build_week_window(today, num_weeks=8)
    fg_weekly_demand = aggregate_weekly(demand_rows, WEEKS, WEEK_BOUNDS, RANGE_START, RANGE_END)
    weekly_pending_delivery = aggregate_weekly(
        [(k, d, q) for k, d, q in po_rows], WEEKS, WEEK_BOUNDS, RANGE_START, RANGE_END
    )
    weekly_planned_consum = explode_demand(pm_items, fg_weekly_demand, WEEKS)
    weekly_results = run_cascade(pm_items, stock, total_pending_po,
                                  weekly_pending_delivery, weekly_planned_consum, WEEKS)
    shortfall_consolidated = consolidate_shortfall(weekly_results, WEEKS)

    # ---- 6. Monthly cascade (3-month rolling window, full item universe) ----
    prog(0.60, "Running the 3-month cascade...")
    MONTHS = build_month_window(today, num_months=3)
    fg_monthly_demand = aggregate_monthly(demand_rows, MONTHS)
    monthly_pending_delivery = aggregate_monthly([(k, d, q) for k, d, q in po_rows], MONTHS)
    monthly_planned_consum = explode_demand(pm_items, fg_monthly_demand, MONTHS)
    monthly_results = run_cascade(pm_items, stock, total_pending_po,
                                   monthly_pending_delivery, monthly_planned_consum, MONTHS)

    # ---- 7. Laminates sheet: fixed item list, 6-month window ----
    prog(0.68, "Building the Laminates sheet...")
    LAM_MONTHS = build_month_window(today, num_months=6)
    lam_items = {k: v for k, v in pm_items.items() if k in ref.LAMINATED_SHEET_ITEMS}
    fg_lam_demand = aggregate_monthly(demand_rows, LAM_MONTHS)
    lam_planned_consum = explode_demand(lam_items, fg_lam_demand, LAM_MONTHS)
    lam_pending_delivery = aggregate_monthly([(k, d, q) for k, d, q in po_rows], LAM_MONTHS)
    lam_results = run_cascade(lam_items, stock, total_pending_po,
                               lam_pending_delivery, lam_planned_consum, LAM_MONTHS)
    for itemk, v in lam_results.items():
        v["uom"] = uom_map.get(itemk, "PCS")

    # ---- 8. Can Pack sheet ----
    prog(0.75, "Building the Can Pack sheet...")
    canpack_out = build_canpack(canpack_master, fg_cfc, fg_ctn, fg_brand, stock, canpack_rows)

    # ---- 9. Write the workbook ----
    prog(0.85, "Writing the formatted workbook...")
    wb = rw.load_template()
    rw.write_canpack_sheet(wb, canpack_out)
    rw.write_week_summary_sheet(wb, weekly_results, WEEKS, WEEK_BOUNDS, nav_map, generated_on)
    rw.write_monthly_summary_sheet(wb, monthly_results, MONTHS, nav_map, generated_on)
    rw.write_shortfall_sheet(wb, shortfall_consolidated, WEEKS, WEEK_BOUNDS, nav_map)
    rw.write_laminates_sheet(wb, lam_results, LAM_MONTHS, nav_map)
    rw.write_possible_error_sheet(wb, possible_error)

    export_file.seek(0)
    rw.write_orders_sheet(wb, export_file, today.strftime("%d-%m-%Y"))
    export_file.seek(0)
    rw.write_m4_sheet(wb, "Export orders in M4", export_file)
    domestic_file.seek(0)
    rw.write_m4_sheet(wb, "Domestic orders in M4", domestic_file)

    stock_file.seek(0)
    rw.write_raw_hidden_sheet(wb, "Stock Report Summary", stock_file)
    pending_po_file.seek(0)
    rw.write_raw_hidden_sheet(wb, "Pending PO", pending_po_file)

    prog(0.95, "Finalizing...")
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    stats = {
        "today": today,
        "week_range": (WEEKS[0], WEEKS[-1], WEEK_BOUNDS[WEEKS[0]][0], WEEK_BOUNDS[WEEKS[-1]][1]),
        "pm_item_count": len(pm_items),
        "weekly_shortfall_items": sum(1 for v in weekly_results.values() if v["summary_short_excess"] < -0.001),
        "consolidated_shortfall_rows": len(shortfall_consolidated),
        "possible_error_rows": len(possible_error),
        "possible_error_distinct_fgs": len({r["fg_name"] for r in possible_error}),
        "po_format": po_format,
        "po_excluded_other_location": po_excluded,
    }
    prog(1.0, "Done.")
    return buf.getvalue(), stats
