import re


def norm(s):
    """Normalize an item/FG name for matching: collapse whitespace, uppercase."""
    if s is None:
        return ""
    return re.sub(r"\s+", " ", str(s).replace("\xa0", " ")).strip().upper()


def norm_disp(s):
    """Normalize for display: collapse whitespace but keep original case."""
    if s is None:
        return ""
    return re.sub(r"\s+", " ", str(s).replace("\xa0", " ")).strip()


def to_float(v, default=0.0):
    if v in (None, ""):
        return default
    try:
        return float(v)
    except (TypeError, ValueError):
        return default
