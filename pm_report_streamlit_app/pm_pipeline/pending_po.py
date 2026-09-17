"""
Pending Purchase Order parser.

Always uses the sheet with header row: Document No, Order Date, Customer Code,
Customer Name, Item No, Item Description, Delivery Date, Quantity,
Outstanding Quantity, Unit Cost - this sheet is the curated "Outstanding Qty"
view (previously seen as Sheet2), never the full raw PO extract.
"""
import openpyxl
from collections import defaultdict
from .utils import norm, to_float
from .weeks import parse_date


def _find_po_sheet(wb):
    for name in wb.sheetnames:
        ws = wb[name]
        header = [ws.cell(1, c).value for c in range(1, 11)]
        if header[5] == "Item Description" and header[6] == "Delivery Date":
            return ws
    # fall back to the last sheet if the exact header match fails
    return wb[wb.sheetnames[-1]]


def parse_pending_po(file_obj):
    """
    Returns:
      total_pending_po: {item_key: total_outstanding_qty}
      rows: list of (item_key, delivery_date, qty)   for weekly/monthly bucketing
    """
    wb = openpyxl.load_workbook(file_obj, data_only=True, read_only=True)
    ws = _find_po_sheet(wb)

    total_pending_po = defaultdict(float)
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row is None or row[5] is None:
            continue
        item_desc, delivery_date, outstanding = row[5], row[6], row[8]
        if outstanding is None:
            continue
        qty = to_float(outstanding)
        if qty == 0:
            continue
        key = norm(item_desc)
        total_pending_po[key] += qty
        d = parse_date(delivery_date)
        rows.append((key, d, qty))

    return dict(total_pending_po), rows
