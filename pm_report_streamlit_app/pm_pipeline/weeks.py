"""
Stateless week-window computation.

Weeks are numbered from a fixed global anchor (Week 1 = the Sunday-starting
week containing 04-Jan-2026), so the same calendar week always gets the same
number no matter which day the report happens to be generated on.

The report always shows a *rolling 8-week window* starting today:
  - Week N = whichever anchor-week contains today's date, truncated to run
    from today through that week's normal Saturday end.
  - The following 7 weeks keep their normal Sunday-Saturday boundaries.
This removes the need to manually track "which week did we start last time" -
every run derives the window fresh from the current date.
"""
from datetime import date, timedelta, datetime

ANCHOR = date(2025, 12, 28)  # a Sunday; Week 1 start
# Calibrated against the source system's own "Planned in NN" columns, which
# is the ground truth for week numbering (confirmed: 07/11-Sep-2026 -> Week 37,
# 14/16-Sep-2026 -> Week 38). Do not "correct" this to a rounder-looking date
# without re-checking against a fresh file's Planned-in-NN header first.


def week_number_for(d: date) -> int:
    """Return the anchor week number (1-indexed) containing date d."""
    days_since_anchor = (d - ANCHOR).days
    return days_since_anchor // 7 + 1


def week_bounds(n: int):
    start = ANCHOR + timedelta(days=(n - 1) * 7)
    end = start + timedelta(days=6)
    return start, end


def build_week_window(today: date, num_weeks: int = 8):
    """
    Returns (WEEKS, WEEK_BOUNDS, RANGE_START, RANGE_END).
    WEEKS is a list of week numbers, WEEK_BOUNDS maps week number -> (start, end).
    The first week is truncated to start on `today`.
    """
    start_week_num = week_number_for(today)
    WEEKS = list(range(start_week_num, start_week_num + num_weeks))
    WEEK_BOUNDS = {}
    for i, w in enumerate(WEEKS):
        s, e = week_bounds(w)
        if i == 0:
            s = today  # truncate the starting week to begin today
        WEEK_BOUNDS[w] = (s, e)
    RANGE_START = WEEK_BOUNDS[WEEKS[0]][0]
    RANGE_END = WEEK_BOUNDS[WEEKS[-1]][1]
    return WEEKS, WEEK_BOUNDS, RANGE_START, RANGE_END


def bucket_week(d: date, WEEKS, WEEK_BOUNDS, RANGE_START, RANGE_END):
    if d is None:
        return None
    if d < RANGE_START:
        return "overdue"
    if d > RANGE_END:
        return "after"
    for n in WEEKS:
        s, e = WEEK_BOUNDS[n]
        if s <= d <= e:
            return n
    return "after"


def month_key(d: date):
    return (d.year, d.month)


def month_label(m):
    names = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
             "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    return f"{names[m[1]-1]}-{str(m[0])[2:]}"


def build_month_window(today: date, num_months: int = 3):
    first = month_key(today)
    months = []
    y, mo = first
    for _ in range(num_months):
        months.append((y, mo))
        mo += 1
        if mo > 12:
            mo = 1
            y += 1
    return months


def bucket_month(d: date, months):
    if d is None:
        return None
    first, last = months[0], months[-1]
    m = month_key(d)
    if m < first:
        return "overdue"
    if m > last:
        return "after"
    return m


def parse_date(v):
    """Parse a date that may arrive as a datetime, date, or common string formats."""
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%m/%d/%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None
