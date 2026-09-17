"""
Loads the reference data that ships *inside* the app repo and rarely changes:
  - Active FG list
  - Exploded BOM (new format: ITEM TYPE / ITEM GROUP built into the sheet)
  - Can Pack master FG+Brand list
  - Nav Item Code mapping
  - The fixed 38-item Laminated sheet list
  - MOQ tier table

If your Active FG list, BOM, Can Pack master, or Nav mapping change, just replace
the corresponding file in data/ and redeploy - no code changes needed.
"""
import openpyxl
from .utils import norm, norm_disp

DATA_DIR = "data"

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

# Fixed list of items shown on the Laminates sheet, using a flat 300-unit MOQ
# tier regardless of their normal category.
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
]
LAMINATED_SHEET_ITEMS = {norm(n) for n in LAMINATED_SHEET_ITEMS_DISPLAY}


def load_active_fgs(path=f"{DATA_DIR}/Active_FGs.xlsx"):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb["Sheet1"]
    active = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0]:
            active.add(norm(row[0]))
    return active


def load_bom(path=f"{DATA_DIR}/Exploded_BOM.xlsx"):
    """
    Returns:
      raw_rows: list of (fg_name, item_name, qty)   [one per BOM line, every FG]
      item_meta: {item_key: (item_group, item_type)}   e.g. (PM-SPECIFIC, CFC)
      fg_brand: {fg_key: brand}
      uom_map: {item_key: uom}
      fg_cfc / fg_ctn: {fg_key: [(item_name, qty), ...]}  (for Can Pack, every FG)
    """
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb["Exploded BOM"]
    raw_rows = []
    item_meta = {}
    fg_brand = {}
    uom_map = {}
    fg_cfc, fg_ctn = {}, {}

    it = ws.iter_rows(min_row=2, values_only=True)
    next(it)  # header row
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
            item_meta[itemk] = (item_group, item_type)  # (ITEM GROUP, ITEM TYPE)
        if itemk not in uom_map and uom:
            uom_map[itemk] = uom
        raw_rows.append((fg, item, qty))
        if item_type == "CFC":
            fg_cfc.setdefault(fgk, []).append((norm_disp(item), qty))
        elif item_type == "CTN":
            fg_ctn.setdefault(fgk, []).append((norm_disp(item), qty))

    return raw_rows, item_meta, fg_brand, uom_map, fg_cfc, fg_ctn


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
    """PM-SPECIFIC -> PM - SPECIFIC, PM-GENERIC -> PM - GENERIC, else unchanged."""
    if t is None:
        return ""
    if t == "PM-SPECIFIC":
        return "PM - SPECIFIC"
    if t == "PM-GENERIC":
        return "PM - GENERIC"
    return t


def suggested_moq(category, item_name, shortfall_amt):
    """MOQ tier lookup, with the Laminated-sheet and Pouch-Garant overrides."""
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
