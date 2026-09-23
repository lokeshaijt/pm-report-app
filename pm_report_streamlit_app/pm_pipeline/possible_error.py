"""
Possible Error sheet: every order line for an FG that has no BOM entry at all,
after excluding internal material-tracking placeholders (FG-prefixed and
STD-prefixed names, e.g. "FG-ALUMINIUM WIRE", "STD 786 BULK, 32 KGS") which
are not real finished goods and don't need a BOM fix.
"""


def _excluded_prefix(fg_name):
    n = fg_name.strip().upper()
    return n.startswith("FG") or n.startswith("STD")


def build_possible_error(possible_error_rows, fg_names_in_bom):
    """
    possible_error_rows: list of (fg_key, fg_display_name, date, pending_prod, nav_doc_no)
    fg_names_in_bom: set of normalized FG names present anywhere in the merged BOM.

    Returns a list of dicts, one per order line for a genuinely missing FG:
      {nav_doc_no, ship_date, fg_name, order_qty, issue}
    """
    out = []
    for fgk, fg_name, ship_date, qty, nav_doc_no in possible_error_rows:
        if fgk in fg_names_in_bom:
            continue
        if _excluded_prefix(fg_name):
            continue
        out.append({
            "nav_doc_no": nav_doc_no if nav_doc_no else "-",
            "ship_date": ship_date,
            "fg_name": fg_name,
            "order_qty": qty,
            "issue": "Not found in BOM",
        })
    out.sort(key=lambda r: (r["fg_name"], r["ship_date"] or ""))
    return out
