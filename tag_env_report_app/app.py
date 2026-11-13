import copy
import io
import json
import sys

# Streamlit Cloud re-runs app.py after a git update but keeps already-imported
# modules in memory. Drop our own package so every run uses the deployed code
# (otherwise a new app.py can meet an old tag_env and fail).
for _name in [m for m in sys.modules if m == "tag_env" or m.startswith("tag_env.")]:
    del sys.modules[_name]

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
      approved Tag &amp; Env Stock workbook automatically, in exactly the same rows, order and format.</p>
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


@st.cache_data(show_spinner=False, max_entries=5)
def _generate(stock_b: bytes, orders_b: bytes, bom_b: bytes, tpl_b: bytes | None, cfg_json: str):
    cfg = json.loads(cfg_json)
    stock, orders, bom = _stock(stock_b), _orders(orders_b), _bom(bom_b)
    tpl = (lambda: io.BytesIO(tpl_b)) if tpl_b else (lambda: TEMPLATE_PATH)
    result = build_report(stock, orders, bom, cfg, master=read_master_rows(tpl(), cfg))
    xlsx = write_report(result, cfg, tpl())
    validation = validate_report(xlsx, tpl(), result, cfg)
    return dict(result=result, xlsx=xlsx, validation=validation, cfg=cfg)


base_cfg = load_config()

# --------------------------------------------------------------------------- #
# Sidebar: settings (defaults reproduce the approved report; nothing to change)
# --------------------------------------------------------------------------- #
with st.sidebar:
    st.header("Settings")
    st.caption("Defaults reproduce the approved report exactly - no changes are needed. "
               "Changes here apply to this session only; use Download settings.json to make them permanent.")
    cfg = copy.deepcopy(base_cfg)
    exact = st.toggle("Formulas exactly as in the approved report", value=base_cfg.get("factor_mode") != "own_block",
                      help="On: the MD20 HS / MAISA / Perfecta Tag formulas read $I$4, $K$4 and $S$4 and Y4 "
                           "shows 1200, exactly like the approved report. Off: every block reads its own "
                           "factor (MD HS envelope = Paper 1425).")
    cfg["factor_mode"] = "reference" if exact else "own_block"
    with st.expander("Conversion factors (pieces per KG)"):
        for name, val in base_cfg["conversion_factors"].items():
            cfg["conversion_factors"][name] = st.number_input(name, value=float(val), min_value=0.0, step=10.0,
                                                              format="%.0f", key=f"f_{name}")
    st.download_button("Download settings.json", json.dumps(cfg, indent=2), "settings.json")

# --------------------------------------------------------------------------- #
# Uploads -> the report is generated automatically once all three are in
# --------------------------------------------------------------------------- #
c1, c2, c3 = st.columns(3)
with c1:
    up_stock = st.file_uploader("1 · Stock Summary Report", type=["xls", "xlsx"])
with c2:
    up_orders = st.file_uploader("2 · Order Status Report", type=["xlsx", "xls"])
with c3:
    up_bom = st.file_uploader("3 · Exploded BOM Report", type=["xlsx", "xls"])

with st.expander("Report template (optional)"):
    st.write("The approved Tag & Env Stock workbook bundled with the app sets the rows, order and format. "
             "Upload a newer approved copy here only if the approved report itself has changed.")
    up_tpl = st.file_uploader("Reference Tag & Env Stock workbook", type=["xlsx"], key="tpl")

# Running brands: carried forward from the reference report's remarks; editable here.
_tpl_src = io.BytesIO(up_tpl.getvalue()) if up_tpl else TEMPLATE_PATH
_master = read_master_rows(_tpl_src, cfg)
_labels = [m.label for m in _master]
_run_word = cfg["remarks"].get("running", "Running").strip().upper()
_default_run = [m.label for m in _master if any(r.strip().upper() == _run_word for r in m.prev_remarks)]
running = _default_run
if cfg.get("fill_remarks", False):   # Remarks column is left empty by default
    with st.expander(f"Running brands ({len(_default_run)} carried forward from the previous report)"):
        st.caption("Remarks rule (first match): **Running** = ticked here and has a pending order -> "
                   "**Multiline Stock** = envelope stock on 2+ machine lines -> **Order Received** = pending "
                   "order in the Order Status Report -> **Leftover Stock** = no pending order. "
                   "Running cannot be read from the three ERP files, so it is taken from the previous "
                   "report; tick or untick brands that are on a machine now.")
        running = st.multiselect("Brands running on a machine", _labels, default=_default_run,
                                 format_func=lambda x: x.strip())
cfg["running_rows"] = [r for r in running]
cfg["running_carry_forward"] = False

if all([up_stock, up_orders, up_bom]):
    try:
        with st.spinner("Generating the Tag & Env Stock report..."):
            st.session_state["out"] = _generate(up_stock.getvalue(), up_orders.getvalue(), up_bom.getvalue(),
                                                up_tpl.getvalue() if up_tpl else None,
                                                json.dumps(cfg, sort_keys=True))
    except Exception as exc:  # show a readable message instead of a stack trace
        st.session_state.pop("out", None)
        st.error(f"Could not generate the report: {exc}")
else:
    st.session_state.pop("out", None)
    missing = [n for n, f in [("Stock Summary", up_stock), ("Order Status", up_orders),
                              ("Exploded BOM", up_bom)] if not f]
    st.info("Upload " + ", ".join(missing) + " - the report is generated automatically once all three are in.")

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
            st.warning("These stock items have no usable Exploded-BOM qty. Their Make Cases use the divisor of "
                       "the same row in the approved report where one exists, otherwise they are left blank. "
                       "A newer Exploded BOM export removes most of them.")
            st.dataframe(result.exceptions, width="stretch", hide_index=True)
    with t5:
        def _color(v):
            return {"PASS": "color: #2e7d32", "FAIL": "color: #c62828"}.get(v, "color: #9e9e9e")
        st.dataframe(out["validation"].style.map(_color, subset=["Result"]),
                     width="stretch", hide_index=True, height=640)
