"""
Can Pack sheet: for every FG in the Can Pack master list, how many cases can
be packed right now given CFC/CTN stock, versus how many are actually on
order.

CFC/CTN stock is shared: a carton can be the exact same physical item used by
several FGs (flavor variants of one line, pack-size variants, a combo pack
drawing on several flavor cartons at once - never assume a 1:1 CFC/FG or
CTN/FG mapping, and never assume two FGs showing the same case count share a
qty-per-case ratio). Processing FGs in the Can Pack master's own row order,
each one draws from whatever is LEFT in the shared pool after every FG above
it has already taken its share - not the full original stock - and what it
actually packs is deducted (in pieces, via its own qty-per-case ratio) before
the next FG down sees the pool.

CFC (in cases), CTN (in cases) and Can Pack all show the same binding
minimum whenever an FG uses both materials, since Can Pack can never exceed
whichever material runs out first. Leftover CFC/CTN instead compares each
material's own independent capacity (ignoring the other material) to what
actually got packed, so a non-binding material's slack is still visible -
but only on the last FG (in sheet order) still drawing on that shared item,
so an only-partially-depleted pool isn't reported as "leftover" over and
over on every FG still ahead of it in the queue.

This sheet deliberately ignores date/week windowing - Total Available Sale
Orders BC is summed across the *entire* order book, no cutoff.
"""
import math


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


def _is_last_user(fgk, items, last_user_map):
    if not items:
        return False
    return all(last_user_map.get(item.upper()) == fgk for item, _ in items)


def build_canpack(master_fgs, fg_cfc, fg_ctn, fg_brand, stock, canpack_rows):
    """
    master_fgs: list of (fg_key, fg_display, brand_from_master), in the exact
    row order the FG & PM STOCK sheet is written in - this is also the order
    the shared CFC/CTN pool cascades through.
    canpack_rows: combined list of (fg_key, pending_ord_qty, pending_prod) across all order files
    """
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

    remaining = {k.upper(): v for k, v in stock.items()}

    rows_out = []
    for fgk, fg_disp, brand_master in master_fgs:
        brand = brand_master or fg_brand.get(fgk, "")
        is_order = fgk in has_order
        order_flag = "ORDER" if is_order else "NO ORDER"

        cfc_items = fg_cfc.get(fgk, [])
        ctn_items = fg_ctn.get(fgk, [])

        cfc_cap = _cap_from_pool(cfc_items, remaining)
        ctn_cap = _cap_from_pool(ctn_items, remaining)
        own_cfc_cases = 0 if cfc_cap in (None, float("inf")) else math.floor(cfc_cap)
        own_ctn_cases = 0 if ctn_cap in (None, float("inf")) else math.floor(ctn_cap)

        caps = [c for c in (cfc_cap, ctn_cap) if c is not None]
        can_pack = math.floor(min(caps)) if caps else 0

        cfc_cases = can_pack if cfc_items else "-"
        ctn_cases = can_pack if ctn_items else "-"

        leftover_cfc = (own_cfc_cases - can_pack) if cfc_items else None
        leftover_ctn = (own_ctn_cases - can_pack) if ctn_items else None
        if leftover_cfc is not None and not _is_last_user(fgk, cfc_items, last_cfc_user):
            leftover_cfc = None
        if leftover_ctn is not None and not _is_last_user(fgk, ctn_items, last_ctn_user):
            leftover_ctn = None

        # what actually gets packed is deducted from the shared pool, in
        # pieces, via each item's own qty-per-case ratio - so the next FG
        # down sharing this item sees the reduced amount.
        for item, qty in cfc_items:
            k = item.upper()
            remaining[k] = remaining.get(k, 0.0) - can_pack * (qty or 0)
        for item, qty in ctn_items:
            k = item.upper()
            remaining[k] = remaining.get(k, 0.0) - can_pack * (qty or 0)

        toab = total_orders_bc.get(fgk, 0.0) if is_order else 0.0
        short_excess = can_pack - toab

        rows_out.append({
            "fg": fg_disp, "brand": brand, "order_flag": order_flag,
            "total_orders_bc": toab,
            "cfc_cases": cfc_cases, "ctn_cases": ctn_cases,
            "leftover_cfc": leftover_cfc, "leftover_ctn": leftover_ctn,
            "can_pack": can_pack, "short_excess": short_excess,
        })
    return rows_out
