"""
Turns raw (key, date, qty) rows into {key: {bucket: total_qty}}, using either
the weekly or monthly bucketing rule - overdue rows roll into the first
bucket, rows after the window are dropped (out of scope), exactly matching
the same overdue-rollover convention used throughout this report.
"""
from collections import defaultdict
from .weeks import bucket_week, bucket_month


def aggregate_weekly(rows, WEEKS, WEEK_BOUNDS, RANGE_START, RANGE_END):
    """rows: list of (key, date, qty). Returns {key: {week: qty}}."""
    out = defaultdict(lambda: defaultdict(float))
    for key, d, qty in rows:
        if d is None:
            continue
        wk = bucket_week(d, WEEKS, WEEK_BOUNDS, RANGE_START, RANGE_END)
        if wk == "overdue":
            out[key][WEEKS[0]] += qty
        elif wk == "after":
            continue
        else:
            out[key][wk] += qty
    return {k: dict(v) for k, v in out.items()}


def aggregate_monthly(rows, MONTHS):
    """rows: list of (key, date, qty). Returns {key: {(year,month): qty}}."""
    out = defaultdict(lambda: defaultdict(float))
    for key, d, qty in rows:
        if d is None:
            continue
        m = bucket_month(d, MONTHS)
        if m == "overdue":
            out[key][MONTHS[0]] += qty
        elif m == "after":
            continue
        else:
            out[key][m] += qty
    return {k: dict(v) for k, v in out.items()}
