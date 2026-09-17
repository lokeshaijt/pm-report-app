"""
Order Status report parser (used for both Export and Domestic files - same layout).

Column layout on the 'Order Status-By shipment Date' sheet (0-indexed):
  5  Buyer Requested Shipment Date  (forward-filled down blank rows of the same order)
  9  Prod Name
  10 Pending Ord Qty
  14 Pending Prod   <- the actual production-need quantity used for demand
"""
import openpyxl
from collections import defaultdict
from .utils import norm, to_float
from .weeks import parse_date

SHEET_NAME = "Order Status-By shipment Date"


def parse_order_rows(file_obj):
    """
    Returns a list of (fg_key, ship_date, pending_prod_qty) for every real
    order line with nonzero Pending Prod. Also returns raw rows needed for
    Total Available Orders / Planned Orders (Can Pack sheet).
    """
    wb = openpyxl.load_workbook(file_obj, data_only=True, read_only=True)
    ws = wb[SHEET_NAME]

    demand_rows = []       # (fg_key, date, pending_prod)   - for weekly/monthly cascade
    canpack_rows = []      # (fg_key, pending_ord_qty, pending_prod) - for Can Pack sheet
    last_date = None

    for row in ws.iter_rows(min_row=7, values_only=True):
        ship_raw = row[5]
        if ship_raw not in (None, " "):
            last_date = ship_raw
        prod_name = row[9]
        if prod_name is None or (isinstance(prod_name, str) and prod_name.strip() == ""):
            continue
        pending_ord_qty = to_float(row[10])
        pending_prod = to_float(row[14])
        fgk = norm(prod_name)

        if pending_ord_qty != 0 or pending_prod != 0:
            canpack_rows.append((fgk, pending_ord_qty, pending_prod))

        if pending_prod != 0:
            d = parse_date(last_date)
            demand_rows.append((fgk, d, pending_prod))

    return demand_rows, canpack_rows


def combine_demand(*order_files):
    """
    order_files: list of file objects (Export, Domestic, ...).
    Returns:
      all_demand_rows: combined list of (fg_key, date, qty) across every file
      all_canpack_rows: combined list of (fg_key, pending_ord_qty, pending_prod)
      order_active_fgs: set of fg_keys with a real order somewhere
    """
    all_demand_rows = []
    all_canpack_rows = []
    order_active_fgs = set()
    for f in order_files:
        demand_rows, canpack_rows = parse_order_rows(f)
        all_demand_rows.extend(demand_rows)
        all_canpack_rows.extend(canpack_rows)
        for fgk, poq, pp in canpack_rows:
            if poq != 0 or pp != 0:
                order_active_fgs.add(fgk)
    return all_demand_rows, all_canpack_rows, order_active_fgs
