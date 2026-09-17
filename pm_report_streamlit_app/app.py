import streamlit as st
from datetime import date, datetime

from pm_pipeline.pipeline import generate_report

st.set_page_config(page_title="PM Weekly Requirement Report", page_icon="📦", layout="wide")

st.markdown(
    """
    <style>
    :root {
        --jay-gold-light: #FDE68A;
        --jay-gold: #F2B705;
        --jay-gold-dark: #B8860B;
        --jay-black: #111111;
        --jay-black-soft: #1C1C1C;
    }

    .block-container {
        padding-top: 2rem;
        max-width: 1100px;
    }

    .jay-hero {
        background: radial-gradient(ellipse at center, #2a2a2a 0%, var(--jay-black) 75%);
        border: 1px solid var(--jay-gold-dark);
        border-radius: 18px;
        padding: 2rem 2.25rem;
        margin-bottom: 1.75rem;
        box-shadow: 0 4px 24px rgba(0, 0, 0, 0.45);
    }

    .jay-hero h1 {
        margin: 0;
        font-size: 2rem;
        font-weight: 800;
        letter-spacing: 0.02em;
        background: linear-gradient(180deg, var(--jay-gold-light) 0%, var(--jay-gold) 45%, var(--jay-gold-dark) 100%);
        -webkit-background-clip: text;
        background-clip: text;
        color: transparent;
    }

    .jay-hero p {
        margin: 0.5rem 0 0 0;
        color: #C9C9C9;
        font-size: 0.95rem;
        line-height: 1.5;
    }

    .jay-section-label {
        display: inline-block;
        font-size: 0.8rem;
        font-weight: 700;
        letter-spacing: 0.08em;
        text-transform: uppercase;
        color: var(--jay-black);
        background: linear-gradient(135deg, var(--jay-gold-light), var(--jay-gold));
        padding: 0.2rem 0.7rem;
        border-radius: 999px;
        margin-bottom: 0.6rem;
    }

    div[data-testid="stExpander"] {
        border: 1px solid #333333;
        border-radius: 12px;
        background-color: var(--jay-black-soft);
    }

    div[data-testid="stFileUploaderDropzone"] {
        border: 1.5px dashed var(--jay-gold-dark);
        border-radius: 10px;
        background-color: var(--jay-black-soft);
    }

    .stButton > button, .stDownloadButton > button {
        border-radius: 10px;
        font-weight: 700;
        border: 1px solid var(--jay-gold-dark);
    }

    div[data-testid="stProgress"] div[role="progressbar"] > div {
        background: linear-gradient(90deg, var(--jay-gold-dark), var(--jay-gold));
    }

    hr {
        border-color: #333333 !important;
    }
    </style>

    <div class="jay-hero">
        <h1>📦 PM Weekly Requirement Report Generator</h1>
        <p>
            Upload the 4 files that change every run. Everything else (Active FG list,
            BOM, Can Pack master, Nav Item Code mapping, MOQ tiers, formatting template)
            is bundled with this app.
        </p>
    </div>
    """,
    unsafe_allow_html=True,
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

st.markdown('<span class="jay-section-label">Step 1</span>', unsafe_allow_html=True)
st.subheader("Upload today's files")

with st.container(border=True):
    col1, col2 = st.columns(2)
    with col1:
        export_file = st.file_uploader("Export Order Status", type=["xlsx"], key="export")
        pending_po_file = st.file_uploader("Pending PO", type=["xlsx"], key="po")
    with col2:
        domestic_file = st.file_uploader("Domestic Order Status", type=["xlsx"], key="domestic")
        stock_file = st.file_uploader("Stock Report Summary", type=["xlsx"], key="stock")

st.markdown('<span class="jay-section-label">Step 2</span>', unsafe_allow_html=True)
st.subheader("Report date")
report_date = st.date_input("Treat this as 'today' for the report", value=date.today())

st.markdown('<span class="jay-section-label">Step 3</span>', unsafe_allow_html=True)
st.subheader("Generate")
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
