# PM Weekly Requirement Report Generator

A Streamlit app that reproduces the full PM report pipeline: upload the 4
files that change every run (Export orders, Domestic orders, Pending PO,
Stock Report) and download the finished, formatted `.xlsx`.

## What's included right now

- FG & PM STOCK (Can Pack, with correct CFC/CTN case-conversion)
- PM Week Wise Summary (8-week rolling window, no-floor cascade)
- PM Monthly Summary (3-month view, same logic)
- Laminates (fixed 38-item list, 6-month view)
- Shortfall (one row per item, incremental weekly amounts, Suggested MOQ,
  PO Issue-by/Arrive-by)
- Orders, Export orders in M4, Domestic orders in M4

All formatting (colors, freeze panes, column widths, conditional formatting)
comes from `data/Report_Template.xlsx` - edit that file and redeploy if the
approved format ever changes; no code changes needed for pure formatting
tweaks.

## Local run

```bash
pip install -r requirements.txt
streamlit run app.py
```

Opens at `http://localhost:8501`.

## Deploying for a public shareable link (Streamlit Community Cloud - free)

1. **Create a GitHub repo** and push this whole folder to it (including the
   `data/` folder - it has to be in the repo since Streamlit Cloud pulls
   straight from GitHub, there's no separate file upload step for the
   reference data).
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   git remote add origin https://github.com/<your-username>/<repo-name>.git
   git push -u origin main
   ```
   Note: `Exploded_BOM.xlsx` is ~12 MB. GitHub's normal size limit is 100 MB
   per file, so this is fine as a regular commit - no Git LFS needed.

2. **Go to** [share.streamlit.io](https://share.streamlit.io) and sign in
   with your GitHub account.

3. Click **"New app"**, pick your repo, branch (`main`), and set the main
   file path to `app.py`.

4. Click **Deploy**. After a minute or two you'll get a public URL like
   `https://<your-app-name>.streamlit.app` - that's the shareable link.

5. **Whenever the reference data changes** (Active FG list, BOM, Can Pack
   master, Nav mapping, or the approved report format): replace the file in
   `data/`, commit, and push - Streamlit Cloud auto-redeploys.

## Updating the calculation logic

Everything data-related lives in `pm_pipeline/`:
- `weeks.py` - date/week/month bucketing (stateless, derives everything from
  the report date)
- `reference_data.py` - MOQ tiers, the fixed Laminated item list, loaders for
  the bundled files
- `stock.py`, `pending_po.py`, `orders.py` - the 3 uploaded-file parsers
- `pm_universe.py` - the Active-FG ∪ order-active-FG universe rule
- `cascade.py` - the no-floor cascade and the incremental shortfall fix
- `canpack.py` - Can Pack sheet logic
- `report_writer.py` - all the Excel-writing/formatting code
- `pipeline.py` - orchestrates all of the above

## Known limitations of this first version

- The **Stock Report** parser expects either a raw dump sheet
  (`Location_Name`/`Warehouse_Name`/`Item_Name`/`Qty Total`) or a pivot sheet
  starting with `Row Labels` in A1 with a `Grand Total` column - matching
  every format seen in practice so far. If the source system changes its
  export layout again, `pm_pipeline/stock.py` is the place to adjust.
- The **Pending PO** parser looks for a sheet whose header row has "Item
  Description" in column F and "Delivery Date" in column G - if that ever
  moves, update `pm_pipeline/pending_po.py`.
- This app is stateless between runs by design (each run picks today's date
  fresh) - if you ever need it to remember "we already treated Week 38 as
  starting on the 14th, don't restart the truncation," that would need a
  small persistence layer (e.g. writing the last-used date to a file in the
  repo, or a tiny database) - not included here since the pure date-driven
  approach is simpler and self-correcting.
