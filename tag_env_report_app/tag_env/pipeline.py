"""One-call entry point used by the Streamlit app and the CLI."""
from __future__ import annotations

from dataclasses import dataclass

import pandas as pd

from .config import TEMPLATE_PATH, load_config
from .engine import ReportResult, build_report, row_values
from .template import read_master_rows
from .readers import read_exploded_bom, read_order_status, read_stock_summary
from .validate import validate_report
from .writer import write_report


@dataclass
class PipelineOutput:
    xlsx: bytes
    result: ReportResult
    preview: pd.DataFrame
    validation: pd.DataFrame
    cfg: dict


def preview_frame(result: ReportResult, cfg: dict) -> pd.DataFrame:
    """The main sheet as a DataFrame with the values the formulas evaluate to."""
    cols = []
    for block in cfg["blocks"]:
        for inp in block["inputs"]:
            cols += [(block["name"], inp["kg_col"], f"{inp['kind'].title()} Kgs"),
                     (block["name"], inp["cases_col"], "Can Make Cases")]
        cols.append((block["name"], block["max_col"], "Max Can Make"))
    recs = []
    for row in result.rows:
        v = row_values(row, cfg)
        rec = {"Brand": row.label.strip()}
        for blk, col, name in cols:
            rec[f"{blk} | {name} ({col})"] = v.get(col)
        rec["Total (AL)"] = v.get(cfg["row_total"]["col"])
        if cfg.get("fill_remarks", False):
            rec["Remarks"] = row.remark
        recs.append(rec)
    return pd.DataFrame(recs)


def run(stock_bytes: bytes, order_bytes: bytes, bom_bytes: bytes, cfg: dict | None = None,
        template_path=TEMPLATE_PATH) -> PipelineOutput:
    cfg = cfg or load_config()
    stock = read_stock_summary(stock_bytes)
    orders = read_order_status(order_bytes)
    bom = read_exploded_bom(bom_bytes)
    result = build_report(stock, orders, bom, cfg, master=read_master_rows(template_path, cfg))
    xlsx = write_report(result, cfg, template_path)
    validation = validate_report(xlsx, template_path, result, cfg)
    return PipelineOutput(xlsx=xlsx, result=result, preview=preview_frame(result, cfg),
                          validation=validation, cfg=cfg)
