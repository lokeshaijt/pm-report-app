"""
Builds the PM item universe used by every report sheet.

Rule (established after tracing several "why is this item's demand wrong"
questions): the FG universe for demand purposes is

    Active FG list  UNION  any FG with a real order in Export or Domestic

Every category/type from the BOM is included as-is - there is no filtering
by ITEM TYPE or Category (that filter existed early on and was explicitly
removed: "take all item category like blends, blendb, tag, ... as per BOM").

Each contributing FG's demand is exploded using *its own* BOM qty-per-case
ratio for that component - never borrowed from a different FG's ratio, even
when two FGs share the same physical packaging item under different names.
"""
from .utils import norm, norm_disp
from .reference_data import fmt_item_type


def build_pm_universe(raw_bom_rows, item_meta, active_fgs, order_active_fgs):
    demand_fg_universe = active_fgs | order_active_fgs

    seen = set()
    pm_items = {}
    for fg, item, qty in raw_bom_rows:
        fgk = norm(fg)
        if fgk not in demand_fg_universe:
            continue
        itemk = norm(item)
        qtyk = round(float(qty), 8) if qty is not None else 0
        dedupe_key = (fgk, itemk, qtyk)
        if dedupe_key in seen:
            continue
        seen.add(dedupe_key)

        meta = item_meta.get(itemk)
        if meta is None:
            item_type_raw, category = None, "(unmatched)"
        else:
            item_type_raw, category = meta

        if itemk not in pm_items:
            pm_items[itemk] = {
                "name": norm_disp(item),
                "category": category or "(blank)",
                "item_type": fmt_item_type(item_type_raw),
                "bom_pairs": [],  # list of (fg_key, qty_per_case)
            }
        pm_items[itemk]["bom_pairs"].append((fgk, qtyk))

    return pm_items
