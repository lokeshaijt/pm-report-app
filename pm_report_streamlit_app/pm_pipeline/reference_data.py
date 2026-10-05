"""
Loads the reference data that ships *inside* the app repo and rarely changes:
  - Exploded BOM (new format: ITEM TYPE / ITEM GROUP built into the sheet) -
    the FG universe is EVERY FG in this file; there is no Active FG list filter.
  - Exploded_BOM_Supplemental.xlsx - a standing merge of extra brands whose main-BOM
    entries were missing (GV, INDUS, JUNGLE KING, LALKUMBH, PREMIER, added 19-Sep-2026).
    This file has no ITEM TYPE/ITEM GROUP columns, so new component names get their
    category/type inferred from the naming convention (see infer_meta below).
  - Can Pack master FG+Brand list
  - Nav Item Code mapping
  - The fixed Laminated sheet list (40 items as of 21-Sep-2026: 27 LAMINATED ROLL +
    11 POUCH TEA INDIA CHAI MMNTS + POUCH LAMINATED TEA INDIA 3 LB + 2 POLY TEA INDIA)
  - MOQ tier table

If the Can Pack master or Nav mapping change, just replace the corresponding file in
data/ and redeploy - no code changes needed. If the main BOM changes, replace
Exploded_BOM.xlsx. The person can also upload a one-off supplemental BOM at runtime
through the app (for new/missing FGs flagged on the Possible Error sheet) without
needing a redeploy - see pm_universe.py.
"""
import os

import openpyxl
from .utils import norm, norm_disp

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")

MOQ_TIERS = {
    "CFC": [250, 500, 1000, 3000, 5000, 10000],
    "CTN": [5000, 10000, 25000, 50000, 75000, 100000, 150000, 200000],
    "TRAY": [250, 500, 1000, 3000, 5000, 10000],
    "T-SHIRT": [1000],
    "ALUMINIUM WIRE": [1000],
    "AL-WIRE": [1000],
    "BOPP FILM": [500],
    "BOPP": [500],
    "ANGLE": [500],
    "SLIP SHEET": [250],
    "SLIP-SHEET": [250],
    "THREAD": [500],
    "POUCH GARANT": [5000, 10000, 25000, 50000],
}

LAMINATED_SHEET_ITEMS_DISPLAY = [
    "LAMINATED ROLL AKOUNA COFFEE 3 IN 1", "LAMINATED ROLL AKOUNA INSTANT COFFEE (1.5GMS)",
    "LAMINATED ROLL AKOUNA PREMIUM CHOCO GRANULE", "LAMINATED ROLL AL FINA CAPPUCCINO 3 IN 1",
    "LAMINATED ROLL AL FINA COFFEE 3 IN 1 20G", "LAMINATED ROLL AL FINA PREMIUM CHOCO GRANULE",
    "LAMINATED ROLL ALIA INSTANT COFFEE (2GMS)", "LAMINATED ROLL JAAGO 100 GM",
    "LAMINATED ROLL JAAGO 250 GM", "LAMINATED ROLL PADI HOT CHOCOLATE 3 IN 1 (30GMS)",
    "LAMINATED ROLL TEZ CAFE 3 IN 1 COFFEE 35G", "LAMINATED ROLL TEZ CAFE 3 IN 1 HOT CHOCOLATE 30G",
    "LAMINATED ROLL TEZ CAFE CAPPUCCINO 3 IN 1", "LAMINATED ROLL TEZ CAFE INSTANT COFFEE (2GMS)",
    "LAMINATED ROLL TEZ CAFE PREMIUM CHOCO GRANULE", "LAMINATED ROLL TEZ ELAICHI 100 GM",
    "LAMINATED ROLL TEZ ELAICHI Rs.10/-", "LAMINATED ROLL TEZ ELAICHI Rs.5/-",
    "LAMINATED ROLL TEZ ELAICHI Rs.5/- (3 TRACK)", "LAMINATED ROLL TEZ F&S 250 GM",
    "LAMINATED ROLL TEZ PREMIUM 15 GM", "LAMINATED ROLL TEZ PREMIUM 30 GM",
    "LAMINATED ROLL TEZ RED 100 GM", "LAMINATED ROLL TEZ RED 250 GM",
    "LAMINATED ROLL TEZ RED 250 GM ELAICHI", "LAMINATED ROLL TEZ RED 35 TO 45 GMS",
    "POUCH LAMINATED TEA INDIA 3 LB",
    "POUCH TEA INDIA 10 CHAI MMNTS GNGR TRMRIC CL", "POUCH TEA INDIA 10 CHAI MMNTS INST CINN CL",
    "POUCH TEA INDIA 10 CHAI MMNTS INST CRDM CL (V2)", "POUCH TEA INDIA 10 CHAI MMNTS INST GNGR CL (V2)",
    "POUCH TEA INDIA 10 CHAI MMNTS INST MASALA CL (V2)", "POUCH TEA INDIA 10 CHAI MMNTS INST MILK CL",
    "POUCH TEA INDIA 10 CHAI MMNTS INST UNSWTND CRDM", "POUCH TEA INDIA 10 CHAI MMNTS INST UNSWTND GNGR",
    "POUCH TEA INDIA 10 CHAI MMNTS INST UNSWTND M", "POUCH TEA INDIA 10 CHAI MMNTS LEMON GRASS CL",
    "POUCH TEA INDIA 10 CHAI MMNTS SAFFRON CL",
    "POLY TEA INDIA 1 LBS", "POLY TEA INDIA 2 LBS",
]
LAMINATED_SHEET_ITEMS = {norm(n) for n in LAMINATED_SHEET_ITEMS_DISPLAY}


def load_bom(path=f"{DATA_DIR}/Exploded_BOM.xlsx"):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb["Exploded BOM"]
    raw_rows = []
    item_meta = {}
    fg_brand = {}
    uom_map = {}
    fg_cfc, fg_ctn = {}, {}

    it = ws.iter_rows(min_row=2, values_only=True)
    next(it)
    for row in it:
        fg, item, item_type, item_group, qty, uom, brand = (
            row[1], row[3], row[4], row[5], row[6], row[7], row[8]
        )
        if fg is None:
            continue
        fgk = norm(fg)
        if brand and fgk not in fg_brand:
            fg_brand[fgk] = brand
        if item is None:
            continue
        itemk = norm(item)
        if itemk not in item_meta:
            item_meta[itemk] = (item_group, item_type)
        if itemk not in uom_map and uom:
            uom_map[itemk] = uom
        raw_rows.append((fg, item, qty))
        # TRAY shares the Can Pack sheet's CFC columns (stock, cascade, Leftover) -
        # it isn't a separate bucket, just another CFC-like item type.
        if item_type in ("CFC", "TRAY"):
            fg_cfc.setdefault(fgk, []).append((norm_disp(item), qty))
        elif item_type == "CTN":
            fg_ctn.setdefault(fgk, []).append((norm_disp(item), qty))

    return raw_rows, item_meta, fg_brand, uom_map, fg_cfc, fg_ctn


def infer_meta(item_name):
    """
    Best-effort (item_group, item_type) inference for a component name with no
    explicit ITEM TYPE/ITEM GROUP columns (used for the bundled supplemental BOM,
    and any one-off supplemental BOM uploaded at runtime). Matches the exact
    conventions used in the main BOM. Returns (None, None) when unrecognized -
    the item then shows as "(unmatched)" rather than being misclassified.
    """
    n = item_name.strip().upper()
    if n.startswith("CFC "):
        return ("PM-SPECIFIC", "CFC")
    if n.startswith("CTN "):
        return ("PM-SPECIFIC", "CTN")
    if n.startswith("ENV "):
        return ("PM-SPECIFIC", "ENV")
    if n.startswith("TAG "):
        return ("PM-SPECIFIC", "TAG")
    if n.startswith("BLEND "):
        return ("BLEND", "BLENDB")
    if n.startswith("LAMINATED ROLL"):
        return ("PM-SPECIFIC", "LAMINATED-ROLLS")
    if n.startswith("LABEL "):
        return ("PM-GENERIC", "LABEL")
    return (None, None)


def _load_supplemental_bom_curated(ws):
    """6-column layout: FG Name, Item Name, Qty, Uom, Brand, FG Uom (no ITEM
    TYPE/ITEM GROUP columns) - title on row 1, header on row 2, data from row 3.
    This is the bundled standing-merge format."""
    raw_rows = []
    fg_brand = {}
    for row in ws.iter_rows(min_row=3, values_only=True):
        if row is None or len(row) < 3:
            continue
        fg, item, qty = row[0], row[1], row[2]
        brand = row[4] if len(row) > 4 else None
        if fg is None or item is None:
            continue
        fgk = norm(fg)
        if brand and fgk not in fg_brand:
            fg_brand[fgk] = brand
        raw_rows.append((fg, item, qty))
    return raw_rows, fg_brand


def _load_supplemental_bom_routing_grid(ws):
    """4-column raw ERP export: BOM_Name, Item_Name, Qty, UOM_Name - header on
    row 1, data from row 2. UOM_Name doubles as a loose item-type hint (CFC,
    CTN, ...) rather than a real unit, and each FG typically carries one
    redundant row where Item_Name just repeats the FG name with UOM_Name=CFC
    alongside a properly-named "CFC <name>" row for the same thing - drop the
    duplicate. CTN-type rows are routinely missing the "CTN " prefix the rest
    of the BOM and Stock Report rely on to resolve against real stock, so it's
    added here when missing."""
    by_fg = {}
    order = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row is None or len(row) < 3:
            continue
        fg, item, qty = row[0], row[1], row[2]
        uom = row[3] if len(row) > 3 else None
        if fg is None or item is None:
            continue
        fg, item = str(fg).strip(), str(item).strip()
        uom = str(uom).strip().upper() if uom else ""
        if fg not in by_fg:
            by_fg[fg] = []
            order.append(fg)
        by_fg[fg].append((item, qty, uom))

    raw_rows = []
    for fg in order:
        items = by_fg[fg]
        has_proper_cfc = any(i.upper().startswith("CFC ") for i, _, _ in items)
        for item, qty, uom in items:
            if uom == "CFC" and item.upper() == fg.upper() and has_proper_cfc:
                continue
            if uom == "CTN" and not item.upper().startswith("CTN "):
                item = f"CTN {item}"
            raw_rows.append((fg, item, qty))
    return raw_rows, {}


def load_supplemental_bom(path):
    """
    Parses a one-off supplemental BOM, auto-detecting which of two layouts it
    is by header content: the curated 6-column format, or a raw 4-column
    "Routing Grid" ERP export. Used both for the bundled standing merge and
    for any file uploaded at runtime.
    """
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    header2 = [ws.cell(2, c).value for c in range(1, 7)]
    if header2[:2] == ["FG Name", "Item Name"]:
        return _load_supplemental_bom_curated(ws)
    header1 = [ws.cell(1, c).value for c in range(1, 5)]
    if header1[:2] == ["BOM_Name", "Item_Name"]:
        return _load_supplemental_bom_routing_grid(ws)
    # fall back to the curated layout's row positions, best-effort
    return _load_supplemental_bom_curated(ws)


def load_canpack_master(path=f"{DATA_DIR}/Can_Pack_Master.xlsx"):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb["Sheet1"]
    master_fgs = []
    for row in ws.iter_rows(min_row=3, values_only=True):
        if row[0]:
            master_fgs.append((norm(row[0]), norm_disp(row[0]), row[1]))
    return master_fgs


def load_nav_map(path=f"{DATA_DIR}/Nav_Mapping.xlsx"):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb["Default"]
    nav_map = {}
    for row in ws.iter_rows(min_row=6, values_only=True):
        name, nav = row[1], row[3]
        if not name:
            continue
        k = norm(name)
        if k not in nav_map:
            nav_map[k] = nav
    return nav_map


def fmt_item_type(t):
    if t is None:
        return ""
    if t == "PM-SPECIFIC":
        return "PM - SPECIFIC"
    if t == "PM-GENERIC":
        return "PM - GENERIC"
    return t


def suggested_moq(category, item_name, shortfall_amt):
    if shortfall_amt is None or shortfall_amt <= 0:
        return None
    name_u = norm(item_name)
    if name_u in LAMINATED_SHEET_ITEMS:
        tiers = [300]
    elif category == "POUCH" and "GARANT" in name_u:
        tiers = MOQ_TIERS.get("POUCH GARANT")
    else:
        tiers = MOQ_TIERS.get(category)
    if not tiers:
        return None
    for t in tiers:
        if shortfall_amt <= t:
            return t
    return round(shortfall_amt * 1.02)
