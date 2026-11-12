# Tag & Env Stock Report Generator

Streamlit app that replaces the manual **Tag & Env Stock** Excel report.
Upload the three ERP exports and download the finished workbook in the approved format.

| Upload | Used for |
|---|---|
| Stock Summary Report (`StockSummaryReport-*.xls`) | Available KG per ENV / TAG item and its **Machine Line Name** |
| Order Status Report (`OrderStatusReport-*.xlsx`) | Pending orders per FG (sheet *Order Status-By Buyer*) – decides which FG's BOM qty is used |
| Exploded BOM (`Exploded_bom-*.xlsx`) | Pieces of each ENV / TAG item per case (CFC) of every FG |

The output uses `templates/Tag_Env_Stock_Template.xlsx` (the approved report) as the
template. Sheet names, order and visibility, headers, merged header cells, fonts, fills,
borders, column widths, number formats, page setup and freeze panes all come from that
file. Only the data rows are rewritten.

## Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

Command line (same output, prints the format-validation table):

```bash
python generate_report.py STOCK.xls ORDERS.xlsx BOM.xlsx [OUT.xlsx]
```

## Business logic

### Make Cases

```
Make Cases = (Available KG × Conversion Factor) / Exploded-BOM qty per case
```

Each cell is written as a live formula, as in the manual report, e.g. `=(D5*$E$4)/240`.
Row 4 holds the conversion factors, and the denominator is looked up from the Exploded BOM
for each item. Nothing is hardcoded.

* **Available KG**: the Stock Summary `Qty` of the item.
* **Conversion Factor**: pieces per KG of the item's machine line (`config/settings.json`,
  from the *Conversion Round Off* table):

  | Machine line | Factor | Report column |
  |---|---|---|
  | Constanta Tag | 10780 | C250 Tag Kgs (B) |
  | CONSTANTA-250 | 1740 | C250 Envelope Kgs (D) |
  | MD-20 Tag | 12600 | MD20 Tag Kgs (H) – items with `HS` go to MD20 HS (V) |
  | MD-20 Crimp Envelope | 1700 | MD20 Envelope Kgs (J) |
  | Constanta Envelope | 1525 | Constanta Envelope Kgs (N) |
  | Perfecta Envelope (PENV) | 1520 | Perfecta Envelope Kgs (R) |
  | MD HS Envelope | 1425 (Paper) / 1190 (Foil) | MD20 HS Envelope Kgs (X) |
  | Maisa Tag / Maisa Envelope | 9080 / 1410 | MAISA (AB / AD) |
  | PK-45, PERFACTA TAG | 10780 | Perfecta Tag / PK45 (AH) |

* **Exploded-BOM qty per case**: for the item, take every FG in the Exploded BOM whose
  `FG Uom` is **CFC**. The qty is the FG's total ENV (or TAG) pieces per case, i.e. bags per
  case: 12/20 → 240, 24/25 → 600. For a variety pack it is the sum of all its flavours
  (4 × 60 = 240), which is how the manual report treats variety packs.
  * If the item is used in several FGs, the FG with the **largest pending order qty** in
    the Order Status Report is used. Single-flavour FGs are preferred over variety packs.
  * If none of its FGs has a pending order, the most common pack size among its
    single-flavour FGs is used (ties → the smaller pack).
* **Missing data**: if the item is not in the Exploded BOM, or its BOM qty is below
  1 PCS per case (a BOM data error), the Make Cases cell is **left blank**. The item is then
  listed on the *Exceptions* tab. Items whose machine line has no report column are listed
  there too.

`Max Can Make` and the `AL` column keep the reference formulas (`MAX(C,E)`, `=S`,
`F+L+P+T`). **Remarks**: `Order Received` if any FG of the row's items has a pending order,
`No Order` if not, and `BOM not found` if none of the row's items is in the BOM.

### Rows

One row per brand/variant. The row key is the item name without `ENV`/`TAG`, cut at the
first machine token (`DC`, `C250`, `PK45`, `PTAG`, `PENV`). Any qualifier after `TBGS` is
kept, e.g. `(MATADI)`, `-FT` or `(v2)`. Paper-weight qualifiers such as `(70 GSM)` are
ignored.

The ENV and TAG items of the same variant share a row. A shared tag (for example
`TAG BARISSIMO DC PTAG TBGS`) gets its own `TAG …` row, so every stock KG appears exactly
once. ENV rows are listed alphabetically, followed by the TAG-only rows, and a SUM total
row closes the sheet.

### Ghana & AB & AG Stock

Items marked `(GHANA)`, `(ABIDJAN)` or `(ANGOLA)` go to their own sections, plus
`ENV INDUS DC TBGS (70GSM)`, which the reference lists under Ghana. Each section has a
total, and a grand total closes the sheet.

## Validation

Every run compares the output with the template, shown on the *Format validation* tab.
It checks:

* sheet names, order and visibility
* column widths
* header text, styles and merges
* the row-4 factors
* that every data and total cell uses a style the template uses in that column
* that every Make Cases cell has the `=(KG*$F$4)/qty` shape
* that the SUM ranges cover all rows
* that there are no external links
* freeze panes and page setup
* KG reconciliation against the Stock Summary

The *Calculation audit* download lists, for every item, the FG chosen, its BOM qty, the
pending qty and the factor used.

## Configuration

Everything that might change lives in `config/settings.json`:

* conversion factors
* machine line → column mapping
* case UOM
* remark texts
* regional sections
* grouping tokens

The sidebar can also override the factors for one run, and its **Download settings.json**
button saves them.

## Differences from the manually prepared reference (28-10-26)

* The manual sheet's MD20 HS, MAISA and Perfecta-Tag blocks referenced the wrong factor
  cells (`$K$4` = 1700 instead of 1200 / 1410, `$S$4` = 1520 instead of 10780). Each block
  now uses its own factor.
* The HS envelope factor follows the conversion table (Paper 1425 by default; Foil 1190 can
  be picked in the sidebar). The manual sheet used 1200.
* Data-area merged cells (manual groupings such as one Barissimo tag spread over nine
  rows) are not recreated. Shared tags get their own row instead, so totals are not counted
  twice.
* External links to the ERP workbook (`[1]Stock Summary…`) are replaced by values, so the
  file opens without an *update links* prompt. Excel recalculates all formulas on open.
