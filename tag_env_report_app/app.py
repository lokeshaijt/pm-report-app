import copy
import io
import json

import pandas as pd
import streamlit as st

from tag_env.config import CONFIG_PATH, TEMPLATE_PATH, load_config
from tag_env.engine import build_report
from tag_env.template import read_master_rows
from tag_env.pipeline import preview_frame
from tag_env.readers import read_exploded_bom, read_order_status, read_stock_summary
from tag_env.validate import validate_report
from tag_env.writer import write_report

st.set_page_config(page_title="Tag & Env Stock Report", page_icon="🏷️", layout="wide")

st.markdown(
    """
    <style>
    .block-container { padding-top: 2rem; max-width: 1300px; }
    .hero { border: 1px solid #B8860B; border-radius: 16px; padding: 1.4rem 1.8rem; margin-bottom: 1.4rem;
            background: radial-gradient(ellipse at center, #2a2a2a 0%, #111111 75%); }
    .hero h1 { margin: 0; font-size: 1.8rem; font-weight: 800;
               background: linear-gradient(180deg, #FDE68A 0%, #F2B705 45%, #B8860B 100%);
               -webkit-background-clip: text; background-clip: text; color: transparent; }
    .hero p { margin: .45rem 0 0 0; color: #C9C9C9; font-size: .93rem; line-height: 1.5; }
    </style>
    <div class="hero">
      <h1>Tag &amp; Env Stock Report</h1>
      <p>Upload the Stock Summary, Order Status and Exploded BOM reports. The app calculates
      Make Cases = (Available KG &times; Conversion Factor) / Exploded-BOM qty per case and fills the
      approved Tag &amp; Env Stock workbook, so it is ready to download in the same format.</p>
    </div>
    """,
    unsafe_allow_html=True,
)


# --------------------------------------------------------------------------- #
# Cached parsing (the Exploded BOM is ~12 MB and takes a few seconds)
# --------------------------------------------------------------------------- #
@st.cache_data(show_spinner=False)
def _stock(data: bytes):
    return read_stock_summary(data)


@st.cache_data(show_spinner=False)
def _orders(data: bytes):
    return read_order_status(data)


@st.cache_data(show_spinner=False)
def _bom(data: bytes):
    return read_exploded_bom(data)


base_cfg = load_config()

# --------------------------------------------------------------------------- #
# Sidebar: conversion factors (editable for this run)
# --------------------------------------------------------------------------- #
with st.sidebar:
    st.header("Conversion factors")
    st.caption("Pieces per KG by Machine Line. Defaults come from `config/settings.json`; "
               "changes here apply to this run only.")
    cfg = copy.deepcopy(base_cfg)
    for name, val in base_cfg["conversion_factors"].items():
        cfg["conversion_factors"][name] = st.number_input(name, value=float(val), min_value=0.0, step=10.0,
                                                          format="%.0f", key=f"f_{name}")
    hs = next(i for b in cfg["blocks"] for i in b["inputs"] if i["kg_col"] == "X")
    hs_choice = st.radio("MD HS Envelope stock is", ["MD HS Envelope Paper", "MD HS Envelope Foil"],
                         index=0 if hs["factor_key"].endswith("Paper") else 1,
                         help="The Stock Summary reports a single 'MD HS Envelope' machine line; "
                              "pick which factor applies to it.")
    hs["factor_key"] = hs_choice
    st.download_button("Download settings.json", json.dumps(cfg, indent=2), "settings.json",
                       help="Save these factors and replace config/settings.json to make them permanent.")

# --------------------------------------------------------------------------- #
# Uploads
# --------------------------------------------------------------------------- #
c1, c2, c3 = st.columns(3)
with c1:
    up_stock = st.file_uploader("1 · Stock Summary Report", type=["xls", "xlsx"])
with c2:
    up_orders = st.file_uploader("2 · Order Status Report", type=["xlsx", "xls"])
with c3:
    up_bom = st.file_uploader("3 · Exploded BOM Report", type=["xlsx", "xls"])

with st.expander("Report template (optional)"):
    st.write("The approved Tag & Env Stock workbook bundled with the app is used as the output format. "
             "Upload a newer approved copy here only if the format has changed.")
    up_tpl = st.file_uploader("Reference Tag & Env Stock workbook", type=["xlsx"], key="tpl")

ready = all([up_stock, up_orders, up_bom])
if st.button("Generate report", type="primary", disabled=not ready, width="stretch"):
    try:
        with st.spinner("Reading input files..."):
            stock = _stock(up_stock.getvalue())
            orders = _orders(up_orders.getvalue())
            bom = _bom(up_bom.getvalue())
        with st.spinner("Calculating and writing the report..."):
            template = io.BytesIO(up_tpl.getvalue()) if up_tpl else TEMPLATE_PATH
            result = build_report(stock, orders, bom, cfg, master=read_master_rows(template, cfg))
            xlsx = write_report(result, cfg, template)
            if up_tpl:
                template.seek(0)
            validation = validate_report(xlsx, template, result, cfg)
        st.session_state["out"] = dict(result=result, xlsx=xlsx, validation=validation, cfg=cfg,
                                       counts=(len(stock.items), len(orders), len(bom)))
    except Exception as exc:  # show a readable message instead of a stack trace
        st.session_state.pop("out", None)
        st.error(f"Could not generate the report: {exc}")

if not ready and "out" not in st.session_state:
    st.info("Upload all three files to enable **Generate report**.")

out = st.session_state.get("out")
if out:
    result, cfg_used = out["result"], out["cfg"]
    preview = preview_frame(result, cfg_used)
    date_txt = result.report_date.strftime("%d-%m-%y") if result.report_date else "latest"
    fname = f"Tag_Env_Stock_as_on_{date_txt}.xlsx"

    n_fail = int((out["validation"].Result == "FAIL").sum())
    m = st.columns(5)
    m[0].metric("Report date", result.report_date.strftime("%d-%m-%Y") if result.report_date else "-")
    m[1].metric("Brand rows", len(result.rows))
    m[2].metric("Main-sheet stock (kg)", f"{result.stock_kg_main:,.0f}")
    m[3].metric("Exceptions", len(result.exceptions))
    m[4].metric("Format checks", "All passed" if n_fail == 0 else f"{n_fail} failed")

    d1, d2 = st.columns(2)
    d1.download_button("⬇️ Download Tag & Env Stock report", out["xlsx"], fname,
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary", width="stretch")
    audit = io.BytesIO()
    with pd.ExcelWriter(audit, engine="openpyxl") as xw:
        result.details.to_excel(xw, sheet_name="Calculation details", index=False)
        result.exceptions.to_excel(xw, sheet_name="Exceptions", index=False)
        out["validation"].to_excel(xw, sheet_name="Format validation", index=False)
    d2.download_button("⬇️ Download calculation audit (FG, BOM qty, factor per item)", audit.getvalue(),
                       f"Tag_Env_Calculation_Audit_{date_txt}.xlsx", width="stretch")

    t1, t2, t3, t4, t5 = st.tabs(["Env & Tag Stock", "Ghana & AB & AG Stock", "Calculation details",
                                  f"Exceptions ({len(result.exceptions)})", "Format validation"])
    with t1:
        st.caption("Values shown are what the workbook formulas evaluate to. Blank Make Cases = no valid "
                   "Exploded-BOM qty for that item (see Exceptions).")
        show = preview.copy()
        for c in show.columns:
            if c in ("Brand", "Remarks"):
                continue
            fmt = "{:,.2f}" if "Kgs" in c else "{:,.0f}"
            show[c] = [("" if pd.isna(v) else ("-" if v == 0 else fmt.format(v)))
                       for v in pd.to_numeric(show[c], errors="coerce")]
        st.dataframe(show, width="stretch", height=560, hide_index=True)
    with t2:
        for sec in result.regions:
            st.subheader(sec.title)
            cols = ["Brand", "Envelope in Kgs"] + (["Tag in Kgs"] if sec.has_tag_col else [])
            df = pd.DataFrame([(i, e, t)[:len(cols)] for i, e, t in sec.rows], columns=cols)
            st.dataframe(df, width="stretch", hide_index=True)
    with t3:
        st.caption("One line per stock item: the FG chosen from the Exploded BOM, its pieces per case, "
                   "the pending order qty that drove the choice, and the conversion factor applied.")
        st.dataframe(result.details, width="stretch", height=560, hide_index=True)
    with t4:
        if result.exceptions.empty:
            st.success("No exceptions.")
        else:
            st.warning("These items could not be calculated reliably and were left blank in Make Cases "
                       "(or were not placed on the report). Update the Exploded BOM export or "
                       "config/settings.json, then regenerate.")
            st.dataframe(result.exceptions, width="stretch", hide_index=True)
    with t5:
        def _color(v):
            return {"PASS": "color: #2e7d32", "FAIL": "color: #c62828"}.get(v, "color: #9e9e9e")
        st.dataframe(out["validation"].style.map(_color, subset=["Result"]),
                     width="stretch", hide_index=True, height=640)
