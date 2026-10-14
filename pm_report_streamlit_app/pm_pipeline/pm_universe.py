"""
Builds the PM item universe used by every report sheet.

Standing rule (changed 18-Sep-2026): the FG universe is EVERY FG present in
the Exploded BOM - there is no Active FG list filter anymore. Any new FG
added to the BOM automatically appears in the report with no manual list
update ever needed.

Two BOM sources are combined:
  1. The main Exploded BOM (data/Exploded_BOM.xlsx) - full ITEM TYPE/ITEM
     GROUP metadata per component.
  2. A supplemental BOM - the bundled standing merge
     (data/Exploded_BOM_Supplemental.xlsx, covering brands whose main-BOM
     entries were missing) PLUS, when the person uploads one, a one-off
     supplemental file for whatever new FGs are still showing on the
     Possible Error sheet. Supplemental files have no ITEM TYPE/ITEM GROUP
     columns, so new component names get inferred metadata (reference_data.infer_meta).

Each contributing FG's demand is exploded using *its own* BOM qty-per-case
ratio for that component - never borrowed from a different FG's ratio, even
when two FGs share the same physical packaging item under different names.
"""
from .utils import norm, norm_disp
from .reference_data import fmt_item_type, infer_meta


def merge_bom(main_raw_rows, item_meta, supplemental_raw_rows_list=None):
    """
    Combines the main BOM's raw rows with however many supplemental BOM row
    lists (the bundled standing merge, plus optionally a runtime-uploaded
    one), inferring (item_group, item_type) for any component name the main
    BOM doesn't already have metadata for. Shared by build_pm_universe (the
    weekly/monthly/Laminates item universe) and build_canpack (the CFC/CTN/
    TRAY/Laminates/Pouch pools), so a component added only via supplemental
    BOM is recognized everywhere, not just in the cascade sheets.

    Returns: (all_raw_rows, merged_item_meta) - merged_item_meta is a new
    dict, safe to mutate without affecting the caller's main-BOM item_meta.
    """
    item_meta = dict(item_meta)
    all_raw_rows = list(main_raw_rows)

    for supp_rows in (supplemental_raw_rows_list or []):
        for fg, item, qty in supp_rows:
            itemk = norm(item)
            if itemk not in item_meta:
                grp, typ = infer_meta(item)
                item_meta[itemk] = (grp, typ)
            all_raw_rows.append((fg, item, qty))

    return all_raw_rows, item_meta


def build_pm_universe(main_raw_rows, item_meta, supplemental_raw_rows_list=None):
    """
    main_raw_rows: list of (fg, item, qty) from the main BOM.
    item_meta: {item_key: (item_group, item_type)} from the main BOM.
    supplemental_raw_rows_list: list of raw_rows lists (each from a supplemental
      BOM file - the bundled one, plus optionally a runtime-uploaded one).
      New component names not already in item_meta get inferred metadata.

    Returns: pm_items, fg_names_in_bom (the full merged FG universe, used by
    the Possible Error sheet to know which order-active FGs are still missing).
    """
    all_raw_rows, item_meta = merge_bom(main_raw_rows, item_meta, supplemental_raw_rows_list)

    fg_names_in_bom = {norm(fg) for fg, item, qty in all_raw_rows}

    seen = set()
    pm_items = {}
    for fg, item, qty in all_raw_rows:
        fgk = norm(fg)
        itemk = norm(item)
        qtyk = round(float(qty), 8) if qty is not None else 0
        dedupe_key = (fgk, itemk, qtyk)
        if dedupe_key in seen:
            continue
        seen.add(dedupe_key)

        meta = item_meta.get(itemk)
        if meta is None or meta == (None, None):
            item_type_raw, category = None, "(unmatched)"
        else:
            item_type_raw, category = meta

        if itemk not in pm_items:
            pm_items[itemk] = {
                "name": norm_disp(item),
                "category": category or "(blank)",
                "item_type": fmt_item_type(item_type_raw),
                "bom_pairs": [],
            }
        pm_items[itemk]["bom_pairs"].append((fgk, qtyk))

    return pm_items, fg_names_in_bom
