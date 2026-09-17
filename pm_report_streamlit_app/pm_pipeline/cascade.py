"""
The core requirement cascade - identical logic for weekly and monthly views,
parameterized by a generic list of "buckets" (week numbers or (year, month) tuples).

Cascade rule (no floor):
    Opening(bucket 1)        = Current Stock
    Short/Excess(bucket)     = Opening(bucket) + Pending Delivery(bucket) - Planned Consum(bucket)
    Opening(next bucket)     = Short/Excess(bucket)      <-- NOT capped at 0

A negative balance is allowed to carry forward and compound if nothing arrives
to offset it - this was a deliberate correction (previously the balance was
floored at 0 each bucket, which understated true cumulative shortfall).
"""
from .reference_data import suggested_moq


def explode_demand(pm_items, demand_by_fg_and_bucket, buckets):
    """
    demand_by_fg_and_bucket: {fg_key: {bucket: qty}}
    Returns {item_key: {bucket: planned_consum}}
    """
    planned_consum = {}
    for itemk, info in pm_items.items():
        pc = {b: 0.0 for b in buckets}
        for fgk, qty_per_case in info["bom_pairs"]:
            fg_dem = demand_by_fg_and_bucket.get(fgk)
            if not fg_dem:
                continue
            for b, dem_qty in fg_dem.items():
                if b in pc:
                    pc[b] += dem_qty * qty_per_case
        planned_consum[itemk] = pc
    return planned_consum


def run_cascade(pm_items, stock, total_pending, pending_delivery_by_item, planned_consum_by_item, buckets):
    """
    Returns {item_key: {..., 'opening':{}, 'pending_delivery':{}, 'planned_consum':{},
                         'short_excess':{}, 'order_qty':, 'current_stock':,
                         'total_pending_po':, 'summary_short_excess':, 'suggested_moq':}}
    """
    results = {}
    for itemk, info in pm_items.items():
        current_stock = stock.get(itemk, 0.0)
        tpo = total_pending.get(itemk, 0.0)
        pending_delivery = pending_delivery_by_item.get(itemk, {})
        planned_consum = planned_consum_by_item.get(itemk, {b: 0.0 for b in buckets})

        opening, short_excess = {}, {}
        op = current_stock
        for b in buckets:
            opening[b] = op
            pd = pending_delivery.get(b, 0.0)
            pc = planned_consum.get(b, 0.0)
            se = op + pd - pc
            short_excess[b] = se
            op = se  # no floor

        order_qty = sum(planned_consum.get(b, 0.0) for b in buckets)
        summary_se = current_stock + tpo - order_qty
        moq = suggested_moq(info["category"], info["name"], -summary_se) if summary_se < -0.001 else None

        results[itemk] = {
            "name": info["name"], "item_type": info["item_type"], "category": info["category"],
            "order_qty": order_qty, "current_stock": current_stock, "total_pending_po": tpo,
            "summary_short_excess": summary_se, "suggested_moq": moq,
            "opening": opening, "pending_delivery": {b: pending_delivery.get(b, 0.0) for b in buckets},
            "planned_consum": planned_consum, "short_excess": short_excess,
        }
    return results


def consolidate_shortfall(results, buckets, issue_lead=4, arrive_lead=2):
    """
    One row per item that has a shortfall somewhere in the window.

    Each bucket's value is the INCREMENTAL new shortfall introduced that bucket
    (this bucket's uncovered consumption), not the raw running balance - this
    avoids double-counting a compounding no-floor deficit across buckets.
    The sum of increments always reconciles with summary_short_excess whenever
    nothing arrives mid-window.
    """
    consolidated = []
    for itemk, v in sorted(results.items(), key=lambda kv: kv[1]["name"]):
        bucket_shortfalls = {}
        prev_magnitude = 0.0
        for b in buckets:
            se = v["short_excess"][b]
            magnitude = -se if se < -0.001 else 0.0
            increment = magnitude - prev_magnitude
            if increment > 0.001:
                bucket_shortfalls[b] = increment
            prev_magnitude = magnitude
        if not bucket_shortfalls:
            continue
        total_shortfall = sum(bucket_shortfalls.values())
        moq = suggested_moq(v["category"], v["name"], total_shortfall)

        # only meaningful for week-number buckets (integers), not month tuples
        issue_by, arrive_by = {}, {}
        if buckets and isinstance(buckets[0], int):
            for b, qty in bucket_shortfalls.items():
                issue_w = b - issue_lead
                if issue_w >= buckets[0]:
                    issue_by[issue_w] = qty
                    arrive_by[issue_w + arrive_lead] = qty

        consolidated.append({
            "name": v["name"], "item_type": v["item_type"], "category": v["category"],
            "total_shortfall": total_shortfall, "suggested_moq": moq,
            "bucket_shortfalls": bucket_shortfalls,
            "issue_by": issue_by, "arrive_by": arrive_by,
        })
    return consolidated
