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

Upload the three files and the report is generated automatically.

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
  1 PCS per case (a BOM data error), the divisor that the reference report used for that
  row and column is used instead. If there is none (a new row), the Make Cases cell is
  **left blank**. Either way the item is listed on the *Exceptions* tab. Items whose machine line has no report column are listed
  there too.

`Max Can Make` and the `AL` column keep the reference formulas (`MAX(C,E)`, `=S`,
`F+L+P+T`).

**Remarks (column AM)** are left **empty** by default (`"fill_remarks": false`); the
`Remarks` header and formatting stay. With `"fill_remarks": true` they use the approved
report's vocabulary, and the first rule that applies wins:

| Remark | Rule |
|---|---|
| **Running** | The brand is marked running and still has a pending order. Running can't be read from the three ERP files, so it's carried forward from the previous report's remarks, and the app's *Running brands* list lets you tick or untick brands. |
| **Multiline Stock** | Envelope stock on 2 or more machine lines (C250 / MD20 / Constanta / Perfecta / HS / Maisa envelope columns). |
| **Order Received** | Any FG of the row's items has a pending order in the Order Status Report. |
| **Leftover Stock** | No pending order. |

Rows with no stock get no remark. The wording and the `running_rows` list are in
`config/settings.json`. Upload the latest approved report as the template to carry its
Running marks forward.

### Rows

The row list, labels and order come from the reference report (column A of
`templates/Tag_Env_Stock_Template.xlsx`, or of the template uploaded in the app). Rows
with no stock stay in place with zeros, as in the manual report.

* **Matching items to rows.** Each stock item is matched by its row key: the name
  without `ENV`/`TAG`, cut at the first machine token (`DC`, `C250`, `PK45`, `PTAG`,
  `PENV`), keeping any qualifier after `TBGS` such as `(MATADI)` or `-FT`, and ignoring
  `(70 GSM)` / `(WO GLUE)`. So `ENV CHAYA DC SF C250 TBGS`, `TAG CHAYA DC SF TBGS` and
  `ENV CHAYA DC PENV TBGS (WO GLUE)` all land in the reference row `ENV CHAYA DC PENV TBGS`.
* **Brand-level tags** such as `TAG BARISSIMO`, `TAG GV FLAVOUR` or `TAG ICA COMMON` go
  into the first row of that brand. Their KG and Can Make cells are merged across the
  brand's rows, as in the manual report. `shared_tag_rows` in `config/settings.json`
  pins the Jumbo tags to their row ranges.
* **New items** with no row in the reference are inserted in alphabetical position. The
  Format validation tab reports how many.

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

The sidebar can override the factors for one session (defaults need no change), and its
**Download settings.json** button saves them.

## Matching the approved report exactly

The generated workbook is a copy of the approved report with only the data replaced:

* **Rows:** the same rows, labels and order, each with that row's own formatting.
* **Merged cells:** the same merged ranges, such as the brand tag cells.
* **Row 4 and formulas:** the same row-4 values and the same formula shapes. In the
  default `"factor_mode": "reference"`, the MD20 HS / MAISA / Perfecta Tag formulas read
  `$I$4`, `$K$4` and `$S$4`, and Y4 shows 1200, as in the approved report.
* **AL and AM:** the same AL cells, and the same total row.

The *Format validation* tab checks each reference row's formatting and formula shape
against the same row in the approved report.

Deliberate differences:

* **Remarks (AM)** are left empty (switch on with `"fill_remarks": true`).
* **MD20 HS columns V and X** hold the stock KG as values. In the approved report they are
  lookups into another workbook (`[1]MD-20 HS`), which would break and prompt to
  *update links*.
* **New stock items** with no row in the approved report get a new row in alphabetical
  position.
* **Factor mode:** to have each block use its own factor instead, set
  `"factor_mode": "own_block"` (or switch off *Formulas exactly as in the approved
  report* in the sidebar). The MD HS envelope then uses the Paper factor 1425.

## Automatic generation

As soon as all three files are uploaded, the report is generated. The preview, checks and
download appear with no further steps, and changing a file or a setting regenerates it.
