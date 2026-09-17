"""
Can Pack sheet: for every FG in the Can Pack master list, how many cases can
be packed right now given CFC/CTN stock, versus how many are actually on order.

Key correction baked in here: CFC/CTN stock is converted to CASE-equivalent
by dividing each item's raw piece stock by *that specific FG's* qty-per-case
ratio for that item (a shared carton can need 1 piece per case for one FG and
6 pieces per case for a sibling pack-size FG - never assume 1:1).

This sheet deliberately ignores date/week windowing - Total Available Orders
and Planned Orders are summed across the *entire* order book, no cutoff.
"""
import math


def _case_capacity(rows, stock):
    """rows: list of (item_name, qty_per_case). Returns min in cases, or None."""
    if not rows:
        return None
    caps = []
    for item, qty in rows:
        s = stock.get(item.upper(), 0.0)
        q = qty if qty else 0
        caps.append(s / q if q > 0 else float("inf"))
    return min(caps)


def build_canpack(master_fgs, fg_cfc, fg_ctn, fg_brand, stock, canpack_rows):
    """
    master_fgs: list of (fg_key, fg_display, brand_from_master)
    canpack_rows: combined list of (fg_key, pending_ord_qty, pending_prod) across all order files
    """
    total_orders_bc, planned_orders = {}, {}
    has_order = set()
    for fgk, poq, pp in canpack_rows:
        if poq != 0 or pp != 0:
            has_order.add(fgk)
        total_orders_bc[fgk] = total_orders_bc.get(fgk, 0.0) + poq
        planned_orders[fgk] = planned_orders.get(fgk, 0.0) + pp

    rows_out = []
    for fgk, fg_disp, brand_master in master_fgs:
        brand = brand_master or fg_brand.get(fgk, "")
        is_order = fgk in has_order
        order_flag = "ORDER" if is_order else "NO ORDER"

        cfc_cap = _case_capacity(fg_cfc.get(fgk, []), stock)
        ctn_cap = _case_capacity(fg_ctn.get(fgk, []), stock)
        cfc_cases = 0 if cfc_cap in (None, float("inf")) else math.floor(cfc_cap)
        ctn_cases = 0 if ctn_cap in (None, float("inf")) else math.floor(ctn_cap)

        caps = [c for c in (cfc_cap, ctn_cap) if c is not None]
        can_pack = math.floor(min(caps)) if caps else 0

        toab = total_orders_bc.get(fgk, 0.0) if is_order else 0.0
        plo = planned_orders.get(fgk, 0.0) if is_order else 0.0
        short_excess = (can_pack - plo) if is_order else 0.0

        rows_out.append({
            "fg": fg_disp, "brand": brand, "order_flag": order_flag,
            "total_orders_bc": toab, "planned_orders": plo,
            "cfc_cases": cfc_cases, "ctn_cases": ctn_cases,
            "can_pack": can_pack, "short_excess": short_excess,
        })
    return rows_out
