import streamlit as st
from datetime import date, datetime

from pm_pipeline.pipeline import generate_report

st.set_page_config(page_title="PM Weekly Requirement Report", page_icon="📦", layout="centered")

st.title("📦 PM Weekly Requirement Report Generator")
st.caption(
    "Upload the 4 files that change every run. Everything else (Active FG list, "
    "BOM, Can Pack master, Nav Item Code mapping, MOQ tiers, formatting template) "
    "is bundled with this app."
)

with st.expander("What this does", expanded=False):
    st.markdown("""
    - Combines **Export + Domestic** order demand for every FG that either is on
      the Active FG list *or* has a real order (even if it isn't on the list).
    - Builds an **8-week rolling window** starting today (Week numbers stay
      consistent run to run - no manual tracking needed) and a **3-month**
      PM Monthly Summary, plus the 6-month **Laminates** sheet.
    - Runs the requirement cascade with the *no-floor* rule: a shortfall
      carries forward and compounds until something arrives to offset it.
    - Produces the **Shortfall** sheet with one row per item, incremental
      per-week shortfall amounts (not raw running balance), Suggested MOQ,
      and PO Issue-by / Arrive-by columns (-4 weeks / +2 weeks).
    - Rebuilds **Can Pack** (CFC/CTN converted to case-equivalents using each
      FG's own ratio), and the raw **Export/Domestic orders in M4** sheets.
    """)

st.subheader("1. Upload today's files")

col1, col2 = st.columns(2)
with col1:
    export_file = st.file_uploader("Export Order Status", type=["xlsx"], key="export")
    pending_po_file = st.file_uploader("Pending PO", type=["xlsx"], key="po")
with col2:
    domestic_file = st.file_uploader("Domestic Order Status", type=["xlsx"], key="domestic")
    stock_file = st.file_uploader("Stock Report Summary", type=["xlsx"], key="stock")

st.subheader("2. Report date")
report_date = st.date_input("Treat this as 'today' for the report", value=date.today())

st.subheader("3. Generate")
all_uploaded = all([export_file, domestic_file, pending_po_file, stock_file])

if not all_uploaded:
    st.info("Upload all 4 files to enable report generation.")

generate = st.button("Generate Report", type="primary", disabled=not all_uploaded)

if generate:
    progress_bar = st.progress(0.0)
    status = st.empty()

    def progress_cb(pct, msg):
        progress_bar.progress(pct)
        status.write(msg)

    try:
        xlsx_bytes, stats = generate_report(
            export_file, domestic_file, pending_po_file, stock_file,
            today=report_date, progress_cb=progress_cb,
        )
    except Exception as e:
        st.error(f"Something went wrong: {e}")
        st.exception(e)
    else:
        w1, w2, s, e = stats["week_range"]
        st.success("Report generated.")
        st.markdown(f"""
        **Week range:** Week {w1}-{w2} ({s.strftime('%d-%b-%y')} to {e.strftime('%d-%b-%y')})
        **PM items in universe:** {stats['pm_item_count']:,}
        **Items with a weekly shortfall:** {stats['weekly_shortfall_items']:,}
        **Shortfall sheet rows:** {stats['consolidated_shortfall_rows']:,}
        **Order-active FGs not on the Active FG list:** {stats['order_active_extra_fgs']:,}
        """)
        filename = f"Packing_Materials_-_Weekly_Requirements_{report_date.strftime('%d%b%Y')}.xlsx"
        st.download_button(
            "⬇️ Download report",
            data=xlsx_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

st.divider()
with st.expander("Reference data currently bundled with this app"):
    st.markdown("""
    Replace these files in the `data/` folder and redeploy if they change:
    - `Active_FGs.xlsx`
    - `Exploded_BOM.xlsx`
    - `Can_Pack_Master.xlsx`
    - `Nav_Mapping.xlsx`
    - `Report_Template.xlsx` (controls all formatting: colors, freeze panes, column widths)
    """)
