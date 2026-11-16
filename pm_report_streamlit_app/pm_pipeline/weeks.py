"""
Stateless week-window computation.

Weeks are numbered from an anchor (Week 1 = the Sunday-starting week
containing 04-Jan-2026), so the same calendar week always gets the same
number no matter which day the report happens to be generated on - except
that the source system's own numbering resets to Week 1 again at the start
of each new year, rather than counting up forever. ANCHORS holds every such
reset point found so far, each one the real Sunday that starts that year's
Week 1; week_number_for always resolves a date against whichever anchor's
epoch it actually falls in, so numbering is correct on both sides of a
reset. When the source system resets again for a later year, add that
year's own Week-1 Sunday to ANCHORS - don't just add a generic "Jan 1 of
every year" rule, since (like the original 2025-12-28 anchor) the exact
Sunday has to be confirmed against that year's own "Planned in NN" column,
not assumed from the calendar alone.

The report always shows a *rolling N-week window* starting today:
  - Week N = whichever anchor-week contains today's date, truncated to run
    from today through that week's normal Saturday end.
  - Every following week keeps its normal Sunday-Saturday boundaries, with
    its own number re-resolved from its start date (so a window that spans
    a reset correctly drops back to Week 1 partway through, rather than
    continuing to count up past the old year's last week number).
This removes the need to manually track "which week did we start last time" -
every run derives the window fresh from the current date.
"""
from datetime import date, timedelta, datetime

ANCHORS = [
    date(2025, 12, 28),  # Week 1 of the pre-2027 cycle - a Sunday
    # Calibrated against the source system's own "Planned in NN" columns,
    # which is the ground truth for week numbering (confirmed:
    # 07/11-Sep-2026 -> Week 37, 14/16-Sep-2026 -> Week 38). Do not
    # "correct" this to a rounder-looking date without re-checking against
    # a fresh file's Planned-in-NN header first.
    date(2027, 1, 3),  # Week 1 resets here for the 2027 cycle - a Sunday
]


def _anchor_for(d: date) -> date:
    """The latest ANCHORS entry that is on or before d."""
    anchor = ANCHORS[0]
    for a in ANCHORS:
        if a <= d:
            anchor = a
        else:
            break
    return anchor


def week_number_for(d: date) -> int:
    """Return the anchor week number (1-indexed) containing date d."""
    anchor = _anchor_for(d)
    days_since_anchor = (d - anchor).days
    return days_since_anchor // 7 + 1


def week_start_containing(d: date) -> date:
    """The Sunday starting the week that contains date d."""
    anchor = _anchor_for(d)
    days_since_anchor = (d - anchor).days
    return anchor + timedelta(days=(days_since_anchor // 7) * 7)


def week_bounds(n: int):
    """
    Start/end dates for week number n. Week numbers repeat across a reset
    (e.g. there's a Week 1 in both the pre-2027 and 2027 cycles), so this
    can only disambiguate by always resolving against the most recent
    anchor - correct for the current/upcoming cycle (the normal case this
    is used for), but not for a week number from an older, already-reset
    cycle. Callers that need to tell cycles apart (like build_week_window,
    which can roll across a reset mid-window) resolve bounds from an actual
    date via week_start_containing instead of calling this.
    """
    anchor = ANCHORS[-1]
    start = anchor + timedelta(days=(n - 1) * 7)
    end = start + timedelta(days=6)
    return start, end


def build_week_window(today: date, num_weeks: int = 8):
    """
    Returns (WEEKS, WEEK_BOUNDS, RANGE_START, RANGE_END).
    WEEKS is a list of week numbers, WEEK_BOUNDS maps week number -> (start, end).
    The first week is truncated to start on `today`.
    """
    WEEKS = []
    WEEK_BOUNDS = {}
    cur_start = week_start_containing(today)
    for i in range(num_weeks):
        w = week_number_for(cur_start)
        s, e = cur_start, cur_start + timedelta(days=6)
        if i == 0:
            s = today  # truncate the starting week to begin today
        WEEKS.append(w)
        WEEK_BOUNDS[w] = (s, e)
        cur_start += timedelta(days=7)
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
