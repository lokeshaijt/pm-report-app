"""Command-line use: python generate_report.py STOCK ORDERS BOM [OUTPUT.xlsx]"""
import sys
from pathlib import Path

from tag_env.pipeline import run


def main():
    if len(sys.argv) < 4:
        sys.exit(__doc__)
    stock, orders, bom = (Path(p).read_bytes() for p in sys.argv[1:4])
    out = run(stock, orders, bom)
    date_txt = out.result.report_date.strftime("%d-%m-%y") if out.result.report_date else "latest"
    target = Path(sys.argv[4]) if len(sys.argv) > 4 else Path(f"Tag_Env_Stock_as_on_{date_txt}.xlsx")
    target.write_bytes(out.xlsx)
    print(out.validation.to_string(index=False))
    print(f"\n{len(out.result.rows)} rows, {len(out.result.exceptions)} exceptions -> {target}")


if __name__ == "__main__":
    main()
