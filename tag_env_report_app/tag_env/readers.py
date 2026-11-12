"""Readers for the three uploaded input files.

Each reader locates its header row by column *names* (not fixed positions),
so small layout changes in the exported reports do not break parsing.
"""
from __future__ import annotations

import io
import re
from dataclasses import dataclass
from datetime import date, datetime

import openpyxl
import pandas as pd


def norm(text) -> str:
    """Upper-case, collapse whitespace - used for every name comparison."""
    if text is None:
        return ""
    return re.sub(r"\s+", " ", str(text).replace("\xa0", " ")).strip().upper()


def _load_workbook(data: bytes, read_only: bool = True):
    """Open an Excel file by its content.

    The Stock Summary export is named ``.xls`` but is really an ``.xlsx``
    (zip) file, so detect the format from the bytes, not the extension.
    """
    if data[:2] == b"PK":
        return openpyxl.load_workbook(io.BytesIO(data), read_only=read_only, data_only=True)
    if data[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":  # legacy BIFF .xls
        import xlrd  # optional dependency, only needed for true .xls files

        book = xlrd.open_workbook(file_contents=data)
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        for sh in book.sheets():
            ws = wb.create_sheet(sh.name[:31])
            for r in range(sh.nrows):
                ws.append(sh.row_values(r))
        return wb
    raise ValueError("Unrecognised file format - please upload an Excel (.xlsx / .xls) file.")


def _find_header(rows, required: list[str], max_scan: int = 30):
    """Return (row_index, {header_name: col_index}) for the first row containing all required headers."""
    for i, row in enumerate(rows[:max_scan]):
        cells = [norm(c) for c in row]
        if all(norm(h) in cells for h in required):
            return i, {c: j for j, c in enumerate(cells) if c}
    raise ValueError(f"Could not find a header row containing: {', '.join(required)}")


# --------------------------------------------------------------------------- #
# Stock Summary Report
# --------------------------------------------------------------------------- #
@dataclass
class StockData:
    items: pd.DataFrame          # columns: item, item_norm, machine_line, qty
    report_date: date | None
    sheet_name: str


def read_stock_summary(data: bytes) -> StockData:
    wb = _load_workbook(data)
    candidates = []
    for ws in wb.worksheets:
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        try:
            h, cols = _find_header(rows, ["Item Name", "Machine Line Name", "Qty"])
        except ValueError:
            continue
        state = getattr(ws, "sheet_state", "visible")
        candidates.append((state != "visible", -len(rows), ws.title, rows, h, cols))
    if not candidates:
        raise ValueError("Stock Summary: no sheet with 'Item Name / Machine Line Name / Qty' headers found.")
    # Prefer a visible sheet. The ERP export has a single sheet; if the user
    # added a sorted copy ("... (2)") pick the original (no " (" suffix).
    candidates.sort(key=lambda c: (c[0], "(" in c[2], c[1]))
    _, _, title, rows, h, cols = candidates[0]

    ci, cm, cq = cols["ITEM NAME"], cols["MACHINE LINE NAME"], cols["QTY"]
    recs = []
    for row in rows[h + 1:]:
        if ci >= len(row) or row[ci] in (None, ""):
            continue
        qty = pd.to_numeric(row[cq] if cq < len(row) else None, errors="coerce")
        recs.append({
            "item": re.sub(r"\s+", " ", str(row[ci]).replace("\xa0", " ")).strip(),
            "item_norm": norm(row[ci]),
            "machine_line": (str(row[cm]).strip() if cm < len(row) and row[cm] not in (None, "") else ""),
            "qty": 0.0 if pd.isna(qty) else float(qty),
        })
    items = pd.DataFrame(recs, columns=["item", "item_norm", "machine_line", "qty"])
    # The same item can appear twice in an export - sum it.
    if not items.empty:
        items = (items.groupby(["item_norm", "machine_line"], as_index=False, sort=False)
                 .agg(item=("item", "first"), qty=("qty", "sum")))[["item", "item_norm", "machine_line", "qty"]]

    report_date = None
    for row in rows[:h + 1]:
        for c in row:
            m = re.search(r"(\d{4}-\d{2}-\d{2})", str(c or ""))
            if m:
                report_date = datetime.strptime(m.group(1), "%Y-%m-%d").date()
                break
        if report_date:
            break
    return StockData(items=items, report_date=report_date, sheet_name=title)


# --------------------------------------------------------------------------- #
# Order Status Report
# --------------------------------------------------------------------------- #
def read_order_status(data: bytes) -> pd.DataFrame:
    """Pending orders per FG: columns prod_id, prod_norm, pending_qty.

    Uses the 'Order Status-By Buyer' sheet (falls back to any sheet having
    Prod Id / Prod Name / Pending Ord Qty). In that sheet the Order Status /
    buyer columns are filled only on the first line of each order, so every
    product line is treated as a pending order line.
    """
    wb = _load_workbook(data)
    order = sorted(wb.worksheets, key=lambda ws: (0 if "BY BUYER" in ws.title.upper() else 1))
    for ws in order:
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        try:
            h, cols = _find_header(rows, ["Prod Id", "Prod Name", "Pending Ord Qty"])
        except ValueError:
            continue
        pi, pn, pq = cols["PROD ID"], cols["PROD NAME"], cols["PENDING ORD QTY"]
        recs = []
        for row in rows[h + 1:]:
            pid = pd.to_numeric(row[pi] if pi < len(row) else None, errors="coerce")
            name = row[pn] if pn < len(row) else None
            if pd.isna(pid) or not str(name or "").strip():
                continue
            qty = pd.to_numeric(row[pq] if pq < len(row) else None, errors="coerce")
            recs.append({"prod_id": int(pid), "prod_name": str(name).strip(), "prod_norm": norm(name),
                         "pending_qty": 0.0 if pd.isna(qty) else float(qty)})
        df = pd.DataFrame(recs, columns=["prod_id", "prod_name", "prod_norm", "pending_qty"])
        df = df[df.pending_qty > 0]
        return (df.groupby(["prod_id", "prod_norm"], as_index=False)
                .agg(prod_name=("prod_name", "first"), pending_qty=("pending_qty", "sum")))
    raise ValueError("Order Status: no sheet with 'Prod Id / Prod Name / Pending Ord Qty' headers found.")


# --------------------------------------------------------------------------- #
# Exploded BOM
# --------------------------------------------------------------------------- #
def read_exploded_bom(data: bytes) -> pd.DataFrame:
    """FG -> component rows: fg_id, fg_name, fg_norm, item, item_norm, item_type, qty, uom, fg_uom."""
    wb = _load_workbook(data)
    sheets = sorted(wb.worksheets, key=lambda ws: (0 if "EXPLODED" in ws.title.upper() else 1))
    for ws in sheets:
        it = ws.iter_rows(values_only=True)
        head = []
        for _ in range(10):
            try:
                head.append(list(next(it)))
            except StopIteration:
                break
        try:
            h, _ = _find_header(head, ["FG Name", "Item Name", "Qty", "FG Uom"], max_scan=10)
        except ValueError:
            continue
        header = [norm(c) for c in head[h]]
        # 'ITEM ID' appears twice (FG id, component id) - take by position.
        idx_item_id = [j for j, c in enumerate(header) if c == "ITEM ID"]
        c_fg = header.index("FG NAME")
        c_item = header.index("ITEM NAME")
        c_type = header.index("ITEM TYPE") if "ITEM TYPE" in header else None
        c_qty = header.index("QTY")
        c_uom = header.index("UOM") if "UOM" in header else None
        c_fguom = header.index("FG UOM")
        c_fgid = idx_item_id[0] if idx_item_id and idx_item_id[0] < c_fg else None

        recs = []
        rest = head[h + 1:]
        for row in _chain(rest, it):
            if c_item >= len(row) or row[c_item] in (None, "") or row[c_fg] in (None, ""):
                continue
            qty = pd.to_numeric(row[c_qty], errors="coerce")
            fgid = pd.to_numeric(row[c_fgid], errors="coerce") if c_fgid is not None else float("nan")
            recs.append((
                None if pd.isna(fgid) else int(fgid),
                _clean(row[c_fg]), norm(row[c_fg]),
                _clean(row[c_item]), norm(row[c_item]),
                norm(row[c_type]) if c_type is not None else "",
                None if pd.isna(qty) else float(qty),
                norm(row[c_uom]) if c_uom is not None else "",
                norm(row[c_fguom]),
            ))
        return pd.DataFrame(recs, columns=["fg_id", "fg_name", "fg_norm", "item", "item_norm",
                                           "item_type", "qty", "uom", "fg_uom"])
    raise ValueError("Exploded BOM: no sheet with 'FG Name / Item Name / Qty / FG Uom' headers found.")


def _clean(text) -> str:
    return re.sub(r"\s+", " ", str(text).replace("\xa0", " ")).strip()


def _chain(first, second):
    yield from first
    for r in second:
        yield list(r)
