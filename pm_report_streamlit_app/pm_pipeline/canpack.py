"""
Can Pack sheet: for every FG in the Can Pack master list, how many cases can
be packed right now given CFC/CTN stock and Pending PO, versus how many are
actually on order.

CFC/CTN stock is shared: a carton can be the exact same physical item used by
several FGs (flavor variants of one line, pack-size variants, a combo pack
drawing on several flavor cartons at once - never assume a 1:1 CFC/FG or
CTN/FG mapping, and never assume two FGs showing the same case count share a
qty-per-case ratio). Processing FGs in the Can Pack master's own row order,
each one draws from whatever is LEFT in the shared pool after every FG above
it has already taken its share - not the full original stock - and what it
actually packs is deducted (in pieces, via its own qty-per-case ratio) before
the next FG down sees the pool.

Pending PO runs through the exact same row-order cascade, but as its own
separate pool - a pending carton isn't physically the same piece as a stock
carton, so the two pools deplete independently. CFC/CTN (in cases) show each
pool's own cascaded capacity; Total CFC/CTN (in cases) is simply their sum,
and Can Pack is based on the Total, since Pending PO can cover part of what
Stock alone falls short of. When Can Pack draws on Pending PO for a bucket,
the Stock pool for that bucket is only ever deducted up to what Stock alone
actually has (it never goes negative) - the rest comes out of the Pending PO
pool instead.

For a shared item, CFC/CTN Stock, Pending PO and Total show only what that
FG actually draws toward its own Can Pack (from_stock/from_po), not the raw
pool size - showing the full pool on every FG sharing it would make it look
like far more is available than really is, once everyone's claim is
accounted for. Only the true last FG (in sheet order) still drawing on that
item shows the real remaining pool, since it's the only one whose number
won't be made stale by someone claiming more after it. An item not shared
with anyone is trivially its own "last user", so this never changes its
numbers. Leftover CFC/CTN is the actual PIECES left over once the whole
chain is done with a shared item - Stock pool and Pending PO pool combined,
since both are real pieces nobody's Can Pack claimed - not a case count,
since a material can be the binding constraint (0 cases of case-level
"slack") while still leaving a handful of pieces that don't add up to one
more case. It's likewise shown only on that same true last FG.

This sheet deliberately ignores date/week windowing - Total Available Sale
Orders BC is summed across the *entire* order book, no cutoff.

fg_cfc/fg_ctn (build_fg_component_maps below) are built from the *full*
merged BOM - main BOM plus every supplemental source - not just the main
BOM, so a component defined only in a supplemental BOM still shows up here.
The CFC bucket already folds in TRAY items (an FG packs one or the other,
never both); the CTN bucket falls back to that FG's LAMINATED-ROLLS or
POUCH items when it has no real CTN entry, since some lines are packed into
laminated rolls or pouches instead of cartons.
"""
import math

from .utils import norm, norm_disp

CFC_LIKE_GROUPS = {"CFC", "TRAY"}
CTN_GROUPS = {"CTN"}
CTN_FALLBACK_GROUPS = {"LAMINATED-ROLLS", "POUCH"}


def build_fg_component_maps(all_raw_rows, merged_item_meta):
    """
    Builds fg_cfc/fg_ctn from the *full* merged BOM (main + every
    supplemental source), not just the main BOM - a component defined only
    in a supplemental BOM (e.g. a newly added brand's CFC/CTN/TRAY items)
    is otherwise invisible to Can Pack even though it's in the FG universe
    everywhere else.

    CFC bucket: item_group CFC or TRAY - an FG packs either one or the
    other, never both, so merging them unconditionally is equivalent to
    "use TRAY when there's no CFC entry" for every real case.

    CTN bucket: item_group CTN when the FG has one. When it doesn't, falls
    back to that FG's LAMINATED-ROLLS or POUCH items instead - some FG
    lines are packed into laminated rolls or pouches rather than cartons,
    and without this fallback those FGs show no CTN capacity at all.
    """
    fg_cfc, fg_ctn, fg_ctn_fallback = {}, {}, {}
    seen = set()
    for fg, item, qty in all_raw_rows:
        itemk = norm(item)
        meta = merged_item_meta.get(itemk)
        if not meta:
            continue
        # item_meta tuples are (item_type, item_group) despite load_bom's own
        # local variable names - verified against real BOM data, where e.g.
        # "ENV AFRICA CHOICE..." stores ('PM-SPECIFIC', 'ENV').
        _item_type, item_group = meta
        if not item_group:
            continue
        fgk = norm(fg)
        qtyk = round(float(qty), 8) if qty is not None else 0
        # the same (FG, item, qty) combo can appear in both the main and a
        # supplemental BOM - count it once, same dedup build_pm_universe uses
        dedupe_key = (fgk, itemk, qtyk)
        if dedupe_key in seen:
            continue
        seen.add(dedupe_key)
        if item_group in CFC_LIKE_GROUPS:
            fg_cfc.setdefault(fgk, []).append((norm_disp(item), qty))
        elif item_group in CTN_GROUPS:
            fg_ctn.setdefault(fgk, []).append((norm_disp(item), qty))
        elif item_group in CTN_FALLBACK_GROUPS:
            fg_ctn_fallback.setdefault(fgk, []).append((norm_disp(item), qty))

    for fgk, items in fg_ctn_fallback.items():
        fg_ctn.setdefault(fgk, items)

    return fg_cfc, fg_ctn


def _cap_from_pool(items, pool):
    """items: list of (item_name, qty_per_case). Returns min in cases, or None."""
    if not items:
        return None
    caps = []
    for item, qty in items:
        s = pool.get(item.upper(), 0.0)
        q = qty if qty else 0
        caps.append(s / q if q > 0 else float("inf"))
    return min(caps)


def _floor_or_zero(cap):
    return 0 if cap in (None, float("inf")) else math.floor(cap)


def _is_last_user(fgk, items, last_user_map):
    if not items:
        return False
    return all(last_user_map.get(item.upper()) == fgk for item, _ in items)


def build_canpack(master_fgs, fg_cfc, fg_ctn, fg_brand, stock, canpack_rows, pending_po=None):
    """
    master_fgs: list of (fg_key, fg_display, brand_from_master), in the exact
    row order the FG & PM STOCK sheet is written in - this is also the order
    the shared CFC/CTN pools cascade through.
    canpack_rows: combined list of (fg_key, pending_ord_qty, pending_prod) across all order files
    pending_po: {item_key: total_outstanding_qty} in pieces (from parse_pending_po), optional
    """
    pending_po = pending_po or {}

    total_orders_bc, planned_prod = {}, {}
    has_order = set()
    for fgk, poq, pp in canpack_rows:
        if poq != 0 or pp != 0:
            has_order.add(fgk)
        total_orders_bc[fgk] = total_orders_bc.get(fgk, 0.0) + poq
        planned_prod[fgk] = planned_prod.get(fgk, 0.0) + pp

    # last sheet-order FG that still draws on a given shared item - used so
    # the true remaining Leftover is shown exactly once per shared material,
    # on whichever FG is actually last in line for it.
    last_cfc_user, last_ctn_user = {}, {}
    for fgk, _, _ in master_fgs:
        for item, _ in fg_cfc.get(fgk, []):
            last_cfc_user[item.upper()] = fgk
        for item, _ in fg_ctn.get(fgk, []):
            last_ctn_user[item.upper()] = fgk

    remaining_stock = {k.upper(): v for k, v in stock.items()}
    remaining_po = {k.upper(): v for k, v in pending_po.items()}

    rows_out = []
    for fgk, fg_disp, brand_master in master_fgs:
        brand = brand_master or fg_brand.get(fgk, "")
        is_order = fgk in has_order
        order_flag = "ORDER" if is_order else "NO ORDER"

        cfc_items = fg_cfc.get(fgk, [])
        ctn_items = fg_ctn.get(fgk, [])

        own_cfc_stock = _floor_or_zero(_cap_from_pool(cfc_items, remaining_stock))
        own_ctn_stock = _floor_or_zero(_cap_from_pool(ctn_items, remaining_stock))
        own_cfc_po = _floor_or_zero(_cap_from_pool(cfc_items, remaining_po))
        own_ctn_po = _floor_or_zero(_cap_from_pool(ctn_items, remaining_po))

        total_cfc = own_cfc_stock + own_cfc_po
        total_ctn = own_ctn_stock + own_ctn_po

        totals = [t for t, items in ((total_cfc, cfc_items), (total_ctn, ctn_items)) if items]
        can_pack = min(totals) if totals else 0

        # Stock is only ever drawn down to what it actually has; whatever
        # portion of Can Pack it can't cover comes out of Pending PO instead.
        cfc_from_stock = min(can_pack, own_cfc_stock) if cfc_items else 0
        cfc_from_po = can_pack - cfc_from_stock if cfc_items else 0
        ctn_from_stock = min(can_pack, own_ctn_stock) if ctn_items else 0
        ctn_from_po = can_pack - ctn_from_stock if ctn_items else 0

        is_last_cfc = _is_last_user(fgk, cfc_items, last_cfc_user)
        is_last_ctn = _is_last_user(fgk, ctn_items, last_ctn_user)

        # Every FG but the true last user of a shared item shows only its
        # own claim on it (from_stock/from_po); the last user shows the
        # real remaining pool, so Total there can exceed its own Can Pack.
        cfc_stock_cases = (own_cfc_stock if is_last_cfc else cfc_from_stock) if cfc_items else "-"
        cfc_po_cases = (own_cfc_po if is_last_cfc else cfc_from_po) if cfc_items else "-"
        ctn_stock_cases = (own_ctn_stock if is_last_ctn else ctn_from_stock) if ctn_items else "-"
        ctn_po_cases = (own_ctn_po if is_last_ctn else ctn_from_po) if ctn_items else "-"
        total_cfc_cases = (cfc_stock_cases + cfc_po_cases) if cfc_items else "-"
        total_ctn_cases = (ctn_stock_cases + ctn_po_cases) if ctn_items else "-"

        for item, qty in cfc_items:
            k = item.upper()
            remaining_stock[k] = remaining_stock.get(k, 0.0) - cfc_from_stock * (qty or 0)
            remaining_po[k] = remaining_po.get(k, 0.0) - cfc_from_po * (qty or 0)
        for item, qty in ctn_items:
            k = item.upper()
            remaining_stock[k] = remaining_stock.get(k, 0.0) - ctn_from_stock * (qty or 0)
            remaining_po[k] = remaining_po.get(k, 0.0) - ctn_from_po * (qty or 0)

        # Leftover = actual pieces left in the Stock pool *and* the Pending
        # PO pool, read right after this row's own deduction from both -
        # correct for the true last user of a shared item, since nothing
        # after it will deduct any further.
        leftover_cfc = round(sum(remaining_stock.get(i.upper(), 0.0) + remaining_po.get(i.upper(), 0.0)
                                  for i, _ in cfc_items)) if cfc_items else None
        leftover_ctn = round(sum(remaining_stock.get(i.upper(), 0.0) + remaining_po.get(i.upper(), 0.0)
                                  for i, _ in ctn_items)) if ctn_items else None
        if leftover_cfc is not None and not is_last_cfc:
            leftover_cfc = None
        if leftover_ctn is not None and not is_last_ctn:
            leftover_ctn = None

        toab = total_orders_bc.get(fgk, 0.0) if is_order else 0.0
        short_excess = can_pack - toab

        rows_out.append({
            "fg": fg_disp, "brand": brand, "order_flag": order_flag,
            "total_orders_bc": toab,
            "cfc_stock_cases": cfc_stock_cases, "cfc_po_cases": cfc_po_cases, "total_cfc_cases": total_cfc_cases,
            "ctn_stock_cases": ctn_stock_cases, "ctn_po_cases": ctn_po_cases, "total_ctn_cases": total_ctn_cases,
            "leftover_cfc": leftover_cfc, "leftover_ctn": leftover_ctn,
            "can_pack": can_pack, "short_excess": short_excess,
        })
    return rows_out
