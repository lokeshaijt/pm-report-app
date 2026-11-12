"""Core business logic for the Tag & Env Stock report.

Make Cases = (Available KG x Conversion Factor) / Exploded-BOM qty

* Available KG       - Stock Summary 'Qty' of the ENV / TAG item.
* Conversion Factor  - pieces per KG of the item's Machine Line (config).
* Exploded-BOM qty   - ENV (or TAG) pieces per case (CFC) of the FG the item
                       is used in, summed over all ENV (or TAG) lines of that
                       FG (= bags per case; e.g. 12/20 -> 240, also for a
                       variety pack of 4 x 60). When the item is used in
                       several FGs, the FG with the largest pending order
                       qty (Order Status Report) is used, single-flavour FGs
                       before variety packs; with no pending order, the most
                       common pack size of its single-flavour FGs (ties ->
                       smaller pack).
"""
from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, field

import pandas as pd

from .readers import StockData, norm
from .template import MasterRow


# --------------------------------------------------------------------------- #
# Data classes
# --------------------------------------------------------------------------- #
@dataclass
class ItemCalc:
    item: str
    machine_line: str
    kg: float
    kg_col: str
    cases_col: str
    factor_key: str
    factor: float | None
    fg_name: str | None = None
    bom_qty: float | None = None
    pending_qty: float = 0.0
    has_order: bool = False
    bom_found: bool = False


@dataclass
class Cell:
    kg: float = 0.0
    items: list[ItemCalc] = field(default_factory=list)
    denominator: float | None = None
    den_source: str = ""        # "Exploded BOM" / "Reference report" / ""


@dataclass
class ReportRow:
    key: str
    label: str
    cells: dict[str, Cell] = field(default_factory=dict)   # by kg column
    remark: str = ""
    row_denominator: float | None = None                   # fallback for empty cells
    in_template: bool = False                              # row exists in the reference report
    template_dens: dict[str, float] = field(default_factory=dict)  # kg_col -> reference divisor
    merge_spans: dict[str, int] = field(default_factory=dict)      # kg_col -> rows a shared tag spans


@dataclass
class RegionSection:
    title: str
    has_tag_col: bool
    rows: list[tuple[str, float, float]]   # (item, env_kg, tag_kg)


@dataclass
class ReportResult:
    rows: list[ReportRow]
    regions: list[RegionSection]
    details: pd.DataFrame
    exceptions: pd.DataFrame
    report_date: object
    factors: dict[str, float]
    stock_kg_main: float = 0.0       # KG of stock items routed to the main sheet
    stock_kg_regions: float = 0.0    # KG of stock items routed to the regional sheet


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def _words(name: str) -> set[str]:
    return set(re.split(r"[\s()\-/]+", norm(name)))


def row_key(item: str, cfg: dict) -> tuple[str, str]:
    """Return (prefix, key) for grouping an item into a report row.

    'ENV BOM DIA BLACK TEA DC SF C250 TBGS (MATADI)' -> ('ENV', 'BOM DIA BLACK TEA (MATADI)')
    'ENV GARANT CHAI MASALA DC SF ENV TBGS-FT'        -> ('ENV', 'GARANT CHAI MASALA -FT')
    'TAG ELDORADO C250 TAG TBGS'                       -> ('TAG', 'ELDORADO')
    """
    g = cfg["grouping"]
    name = re.sub(r"\s+", " ", item).strip()
    for q in g["ignore_qualifiers"]:
        name = re.sub(q, " ", name, flags=re.I)
    name = re.sub(r"\s+", " ", name).strip()
    parts = name.split(" ", 1)
    prefix = parts[0].upper()
    rest = parts[1] if len(parts) > 1 else ""
    tokens = "|".join(re.escape(t) for t in g["machine_tokens"])
    m = re.search(rf"\s(?:{tokens})(?=\s|$)", " " + rest, flags=re.I)
    if m:
        base = (" " + rest)[: m.start()].strip()
        tail_src = (" " + rest)[m.start():]
        t = re.search(r"TBGS(.*)$", tail_src, flags=re.I)
        tail = t.group(1).strip() if t else ""
        if not tail:
            # e.g. 'ENV VICTORIA DC ENV (GHANA)' - keep a trailing (...) qualifier
            p = re.search(r"(\([^)]*\))\s*$", tail_src)
            tail = p.group(1) if p else ""
        key = f"{base} {tail}".strip() if tail else base
    else:
        key = rest
    return prefix, re.sub(r"\s+", " ", key).strip().upper()


def _column_for(item: str, machine_line: str, cfg: dict):
    """Find the (block, input) a stock item feeds, by its Machine Line Name."""
    ml = norm(machine_line)
    is_hs = "HS" in _words(item)
    for block in cfg["blocks"]:
        for inp in block["inputs"]:
            if ml not in {norm(x) for x in inp["machine_lines"]}:
                continue
            if inp.get("hs_only") and not is_hs:
                continue
            if inp.get("hs_excluded") and is_hs:
                continue
            return block, inp
    return None, None


class BomLookup:
    """Per-item denominator selection from the Exploded BOM + pending orders."""

    def __init__(self, bom: pd.DataFrame, orders: pd.DataFrame, cfg: dict):
        b = cfg["bom"]
        case_uoms = {norm(u) for u in b["case_uoms"]}
        comp_uom = norm(b.get("component_uom", ""))
        min_q = float(b.get("min_pcs_per_case", 1))
        sel = bom[bom.fg_uom.isin(case_uoms) & bom.qty.notna()]
        if comp_uom:
            sel = sel[(sel.uom == comp_uom) | (sel.uom == "")]
        # e.g. a TAG listed as 0.0072 PCS per case is a BOM data error
        self.invalid_items = set(sel.loc[sel.qty < min_q, "item_norm"])
        sel = sel[sel.qty >= min_q]
        # Denominator = total pieces of that component type (ENV or TAG) per
        # case of the FG, i.e. bags per case. For a single-flavour FG this is
        # the item's own BOM qty (12/20 -> 240); for a variety pack it is the
        # sum over its flavours (4 x 60 = 240), as in the manual report.
        typ = sel[sel.item_type.isin(["ENV", "TAG"])]
        per_case = typ.groupby(["fg_norm", "item_type"]).qty.sum()
        n_same = typ.groupby(["fg_norm", "item_type"]).item_norm.nunique()
        sel = sel.assign(
            case_qty=[float(per_case.get((f, t), q)) for f, t, q in zip(sel.fg_norm, sel.item_type, sel.qty)],
            # single-flavour FG: the item is the only ENV (or only TAG) line
            single=[int(n_same.get((f, t), 1)) <= 1 for f, t in zip(sel.fg_norm, sel.item_type)],
        )
        self.by_item = {k: g for k, g in sel.groupby("item_norm")}
        self.pending_by_id = orders.groupby("prod_id").pending_qty.sum().to_dict()
        self.pending_by_name = orders.groupby("prod_norm").pending_qty.sum().to_dict()

    def pending_for(self, fg_id, fg_norm) -> float:
        if fg_id is not None and not pd.isna(fg_id) and int(fg_id) in self.pending_by_id:
            return float(self.pending_by_id[int(fg_id)])
        return float(self.pending_by_name.get(fg_norm, 0.0))

    def lookup(self, item_norm: str):
        """Return (fg_name, bom_qty, pending_qty, has_order) or None if not in BOM."""
        g = self.by_item.get(item_norm)
        if g is None or g.empty:
            return None
        g = g.drop_duplicates(["fg_norm"])
        pend = [self.pending_for(r.fg_id, r.fg_norm) for r in g.itertuples()]
        g = g.assign(pending=pend)
        ordered = g[g.pending > 0]
        if not ordered.empty:
            # FG with the largest pending order quantity (single-flavour FGs
            # first, then variety packs).
            best = ordered.sort_values(["single", "pending", "case_qty"],
                                       ascending=[False, False, True]).iloc[0]
            return best.fg_name, float(best.case_qty), float(best.pending), True
        # No pending order: the most common pack size among its single-flavour
        # FGs (all FGs if there is none), ties -> the smaller pack.
        pool = g[g.single] if g.single.any() else g
        counts = Counter(pool.case_qty)
        top = max(counts.values())
        qty = min(q for q, c in counts.items() if c == top)
        best = pool[pool.case_qty == qty].iloc[0]
        return best.fg_name, float(qty), 0.0, False


# --------------------------------------------------------------------------- #
# Main build
# --------------------------------------------------------------------------- #
def _shared_base(key: str, cfg: dict) -> str:
    """'GV FLAVOUR' -> 'GV', 'ICA COMMON' -> 'ICA', 'BOM DIA (MATADI)' -> 'BOM DIA'."""
    words = cfg["grouping"].get("shared_tag_words", [])
    base = re.sub(r"\([^)]*\)", " ", key)
    for w in words:
        m = re.search(rf"\b{re.escape(w.upper())}\b", base)
        if m:
            base = base[: m.start()]
    return re.sub(r"\s+", " ", base).strip(" -")


def build_report(stock: StockData, orders: pd.DataFrame, bom: pd.DataFrame, cfg: dict,
                 master: list[MasterRow] | None = None) -> ReportResult:
    factors = {k: float(v) for k, v in cfg["conversion_factors"].items()}
    lookup = BomLookup(bom, orders, cfg)
    prefixes = {p.upper() for p in cfg["grouping"]["item_prefixes"]}
    overrides = {norm(k): [norm(x) for x in v] for k, v in cfg["grouping"].get("shared_tag_rows", {}).items()}

    regions_cfg = cfg.get("regions", [])
    region_rows = {r["title"]: [] for r in regions_cfg}

    # ---- master rows from the reference report (kept in their order) ----- #
    master_rows: list[ReportRow] = []
    by_key: dict[str, ReportRow] = {}
    for m in master or []:
        _, key = row_key(m.label, cfg)
        if key in by_key:
            continue
        row = ReportRow(key=key, label=m.label, in_template=True, template_dens=dict(m.dens))
        master_rows.append(row)
        by_key[key] = row
    master_index = {r.key: i for i, r in enumerate(master_rows)}
    new_rows: dict[str, ReportRow] = {}
    new_has_env: dict[str, bool] = {}

    def targets(prefix: str, key: str) -> list[ReportRow]:
        if key in by_key:
            return [by_key[key]]
        if key in new_rows:
            return [new_rows[key]]
        if prefix == "TAG" and master_rows:
            # explicitly configured shared tag -> range of reference rows
            if key in overrides:
                ks = [k for k in overrides[key] if k in master_index]
                if ks:
                    i0, i1 = master_index[ks[0]], master_index[ks[-1]]
                    return master_rows[min(i0, i1): max(i0, i1) + 1]
            # brand-level tag (e.g. TAG BARISSIMO, TAG GV FLAVOUR, TAG ICA COMMON):
            # the contiguous reference rows of that brand, as in the manual report
            base = _shared_base(key, cfg)
            if base:
                group = []
                for r in master_rows:
                    if r.key == base or r.key.startswith(base + " "):
                        group.append(r)
                    elif group:
                        break
                if group:
                    return group
        row = ReportRow(key=key, label="")
        new_rows[key] = row
        new_has_env[key] = False
        return [row]

    details, exceptions = [], []
    kg_main = kg_regions = 0.0

    for rec in stock.items.itertuples():
        item, ml, kg = rec.item, rec.machine_line, rec.qty
        prefix = item.split(" ", 1)[0].upper()

        # ---- region (Ghana / Abidjan / Angola) items -------------------- #
        region = None
        for r in regions_cfg:
            if norm(r["marker"]) in rec.item_norm or rec.item_norm in {norm(x) for x in r.get("extra_items", [])}:
                region = r
                break
        if region is not None:
            env_kg = kg if prefix == "ENV" else 0.0
            tag_kg = kg if prefix == "TAG" else 0.0
            region_rows[region["title"]].append((item, env_kg, tag_kg, prefix))
            kg_regions += kg
            continue

        block, inp = _column_for(item, ml, cfg)
        if block is None:
            if prefix in prefixes:
                exceptions.append({"Item": item, "Machine Line": ml, "KG": kg,
                                   "Issue": "Machine Line not mapped to any report column (see config/settings.json)"})
            continue

        kg_main += kg
        _, key = row_key(item, cfg)
        group = targets(prefix, key)
        row = group[0]
        if row.key in new_has_env:
            new_has_env[row.key] |= prefix == "ENV"
        if len(group) > 1:
            row.merge_spans[inp["kg_col"]] = max(row.merge_spans.get(inp["kg_col"], 1), len(group))
            row._groups = {**getattr(row, "_groups", {}), inp["kg_col"]: group}  # type: ignore[attr-defined]

        fkey = inp["factor_key"]
        calc = ItemCalc(item=item, machine_line=ml, kg=kg, kg_col=inp["kg_col"], cases_col=inp["cases_col"],
                        factor_key=fkey, factor=factors.get(fkey))
        hit = lookup.lookup(rec.item_norm)
        if hit:
            calc.fg_name, calc.bom_qty, calc.pending_qty, calc.has_order = hit
            calc.bom_found = True
        cell = row.cells.setdefault(inp["kg_col"], Cell())
        cell.kg += kg
        cell.items.append(calc)

        if calc.factor is None:
            exceptions.append({"Item": item, "Machine Line": ml, "KG": kg,
                               "Issue": f"Conversion factor '{fkey}' missing in config"})
        if not calc.bom_found and kg:
            why = ("Exploded BOM qty per case is below 1 PCS (BOM data error)"
                   if rec.item_norm in lookup.invalid_items else
                   "Item not found in Exploded BOM (as a CFC component)")
            fallback = row.template_dens.get(inp["kg_col"])
            action = (f"divisor {fallback:g} taken from the reference report row '{row.label.strip()}'"
                      if fallback else "Make Cases left blank")
            exceptions.append({"Item": item, "Machine Line": ml, "KG": kg, "Issue": f"{why} - {action}"})

    # ---- labels for new rows -------------------------------------------- #
    for key, row in new_rows.items():
        row.label = f"{'ENV' if new_has_env[key] else 'TAG'} {key} "

    # ---- order: reference rows as-is; new rows slotted in alphabetically -- #
    ordered_rows = list(master_rows)
    for row in sorted(new_rows.values(), key=lambda r: norm(r.label)):
        pref = norm(row.label).split(" ", 1)[0]
        same = [i for i, r in enumerate(ordered_rows) if norm(r.label).split(" ", 1)[0] == pref]
        pos = None
        for i in same:
            if norm(ordered_rows[i].label) > norm(row.label):
                pos = i
                break
        if pos is None:
            pos = (same[-1] + 1) if same else (len(ordered_rows) if pref == "TAG" else
                                                next((i for i, r in enumerate(ordered_rows)
                                                      if norm(r.label).startswith("TAG")), len(ordered_rows)))
        ordered_rows.insert(pos, row)

    # ---- merges: only keep a span if the rows below have nothing of their own
    pos_of = {id(r): i for i, r in enumerate(ordered_rows)}
    for row in ordered_rows:
        for col in list(row.merge_spans):
            group = getattr(row, "_groups", {}).get(col)
            if not group:
                del row.merge_spans[col]
                continue
            # span from the brand's first to last reference row; new rows that
            # were slotted in between are covered too
            i0, i1 = pos_of[id(group[0])], pos_of[id(group[-1])]
            span = ordered_rows[i0: i1 + 1]
            if i0 != pos_of[id(row)] or any(col in r.cells and r.cells[col].items for r in span[1:]):
                del row.merge_spans[col]
            else:
                row.merge_spans[col] = len(span)

    # ---- divisors and remarks ------------------------------------------- #
    rc = cfg["remarks"]
    for row in ordered_rows:
        any_order = any_bom = any_item = False
        rep = None
        for col, cell in row.cells.items():
            if not cell.items:
                continue
            any_item = True
            with_bom = [c for c in cell.items if c.bom_found]
            if with_bom:
                # the item with a pending order (largest pending), else the most KG
                best = sorted(with_bom, key=lambda c: (c.has_order, c.pending_qty, c.kg), reverse=True)[0]
                cell.denominator, cell.den_source = best.bom_qty, "Exploded BOM"
                any_bom = True
                any_order |= any(c.has_order for c in with_bom)
                if rep is None or (best.has_order and not rep[1]):
                    rep = (best.bom_qty, best.has_order)
            elif row.template_dens.get(col):
                cell.denominator, cell.den_source = row.template_dens[col], "Reference report"
        row.row_denominator = rep[0] if rep else None
        if not any_item:
            row.remark = rc.get("no_stock", "")
        else:
            row.remark = rc["with_order"] if any_order else (rc["no_order"] if any_bom else rc["no_bom"])

    for r in ordered_rows:
        for cell in r.cells.values():
            for c in cell.items:
                cases = (c.kg * c.factor / cell.denominator) if (c.factor and cell.denominator) else None
                details.append({
                    "Report Row": r.label.strip(), "Column": c.kg_col, "Item": c.item,
                    "Machine Line": c.machine_line, "KG": c.kg, "Conversion Factor": c.factor,
                    "Factor Source": c.factor_key, "FG Used": c.fg_name, "BOM Qty / Case": c.bom_qty,
                    "Cell Denominator": cell.denominator, "Divisor Source": cell.den_source,
                    "Pending Order Qty": c.pending_qty, "Make Cases": cases,
                    "Row in reference report": "Yes" if r.in_template else "New row",
                })

    regions = []
    for r in regions_cfg:
        items = region_rows[r["title"]]
        extras = {norm(x) for x in r.get("extra_items", [])}
        items.sort(key=lambda t: (t[3] != "ENV", norm(t[0]) in extras, norm(t[0])))
        regions.append(RegionSection(title=r["title"], has_tag_col=r.get("has_tag_col", False),
                                     rows=[(i, e, t) for i, e, t, _ in items]))

    return ReportResult(rows=ordered_rows, regions=regions,
                        details=pd.DataFrame(details),
                        exceptions=pd.DataFrame(exceptions, columns=["Item", "Machine Line", "KG", "Issue"]),
                        report_date=stock.report_date, factors=factors,
                        stock_kg_main=kg_main, stock_kg_regions=kg_regions)


def cell_denominator(row: ReportRow, kg_col: str) -> float | None:
    """Divisor for a Make-Cases cell.

    A cell fed by stock items uses the Exploded-BOM qty (or, if the item is not
    in the BOM, the divisor the reference report used for that row; None if
    neither exists, so no wrong number is produced). An empty cell (0 KG) keeps
    the reference report's divisor, else the row's BOM divisor.
    """
    cell = row.cells.get(kg_col)
    if cell is not None and cell.items:
        return cell.denominator
    return row.template_dens.get(kg_col) or row.row_denominator


def row_values(row: ReportRow, cfg: dict) -> dict[str, float | None]:
    """Compute the numbers the Excel formulas will produce (used for preview & checks)."""
    factors = {k: float(v) for k, v in cfg["conversion_factors"].items()}
    out: dict[str, float | None] = {}
    for block in cfg["blocks"]:
        for inp in block["inputs"]:
            cell = row.cells.get(inp["kg_col"], Cell())
            out[inp["kg_col"]] = cell.kg
            den = cell_denominator(row, inp["kg_col"])
            f = factors.get(inp["factor_key"])
            out[inp["cases_col"]] = (cell.kg * f / den) if (den and f is not None) else None
        expr = block["max_formula"]
        vals = [out.get(c) or 0.0 for c in re.findall(r"\{(\w+)\}", expr)]
        out[block["max_col"]] = max(vals) if expr.startswith("MAX") else vals[0]
    out[cfg["row_total"]["col"]] = sum((out.get(c) or 0.0) for c in re.findall(r"([A-Z]+)\{r\}", cfg["row_total"]["formula"]))
    return out
