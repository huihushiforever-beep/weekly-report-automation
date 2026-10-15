#!/usr/bin/env python3
"""
Weekly report automation - raw exports in, formatted spreadsheet out.

It takes the two files somebody used to copy-paste from every Monday
(orders from the billing system, accounts from the CRM), matches them on the
shared account id, works out the week-over-week numbers, flags the rows that
don't line up, and writes one formatted Excel report.

Usage
-----
    python build_report.py
    python build_report.py --orders sample_data/sales_orders.csv \
                           --accounts sample_data/crm_accounts.csv --out output
"""

from __future__ import annotations

import argparse
import logging
import sys
from datetime import timedelta
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, size=14, color="1F4E79")
LABEL_FONT = Font(bold=True)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY = '#,##0.00'
COUNT = '#,##0'


def build_logger() -> logging.Logger:
    logger = logging.getLogger("report")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    handler = logging.StreamHandler(sys.stdout)
    handler.setFormatter(logging.Formatter("%(asctime)s  %(levelname)-7s %(message)s",
                                           "%Y-%m-%d %H:%M:%S"))
    logger.addHandler(handler)
    return logger


def parse_date(value) -> pd.Timestamp | pd.NaT:
    """The billing system exports ISO dates, the finance team's export uses DD/MM/YYYY."""
    text = str(value).strip()
    for pattern in ("%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d"):
        try:
            return pd.to_datetime(text, format=pattern)
        except (ValueError, TypeError):
            continue
    return pd.NaT


def style_header(sheet, row: int, columns: int) -> None:
    for col in range(1, columns + 1):
        cell = sheet.cell(row=row, column=col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER


def autosize(sheet, widths: dict[int, int]) -> None:
    for index, width in widths.items():
        sheet.column_dimensions[get_column_letter(index)].width = width


def write_table(sheet, start_row: int, headers: list[str], rows: list[list],
                formats: dict[int, str] | None = None) -> int:
    for col, header in enumerate(headers, start=1):
        sheet.cell(row=start_row, column=col, value=header)
    style_header(sheet, start_row, len(headers))
    formats = formats or {}
    for offset, row in enumerate(rows, start=1):
        for col, value in enumerate(row, start=1):
            cell = sheet.cell(row=start_row + offset, column=col, value=value)
            cell.border = BORDER
            if col in formats:
                cell.number_format = formats[col]
    return start_row + len(rows)


def main() -> int:
    parser = argparse.ArgumentParser(description="Build the weekly summary report.")
    parser.add_argument("--orders", default="sample_data/sales_orders.csv")
    parser.add_argument("--accounts", default="sample_data/crm_accounts.csv")
    parser.add_argument("--out", default="output")
    args = parser.parse_args()

    logger = build_logger()
    logger.info("reading %s and %s", args.orders, args.accounts)

    orders = pd.read_csv(args.orders, dtype={"account_id": str, "qty": str})
    accounts = pd.read_csv(args.accounts, dtype={"account_id": str})
    issues: list[tuple[str, str, str]] = []

    orders["order_date_parsed"] = orders["order_date"].apply(parse_date)
    bad_dates = orders[orders["order_date_parsed"].isna()]
    for row in bad_dates.itertuples():
        issues.append((row.order_id, "unreadable order_date", str(row.order_date)))
    orders = orders[orders["order_date_parsed"].notna()].copy()

    orders["qty_num"] = pd.to_numeric(orders["qty"], errors="coerce")
    blank_qty = orders[orders["qty_num"].isna()]
    for row in blank_qty.itertuples():
        issues.append((row.order_id, "missing quantity", row.product))
    orders = orders[orders["qty_num"].notna()].copy()
    orders["unit_price"] = pd.to_numeric(orders["unit_price"], errors="coerce").fillna(0.0)
    orders["revenue"] = orders["qty_num"] * orders["unit_price"]

    merged = orders.merge(accounts, on="account_id", how="left", indicator=True)
    orphans = merged[merged["_merge"] == "left_only"]
    for row in orphans.itertuples():
        issues.append((row.order_id, "account missing from CRM", row.account_id))
    merged = merged[merged["_merge"] == "both"].copy()

    idle_accounts = accounts[~accounts["account_id"].isin(merged["account_id"])]
    for row in idle_accounts.itertuples():
        issues.append((row.account_id, "no orders this period", row.company))

    if merged.empty:
        logger.error("nothing left to report after cleaning - check the input files")
        return 1

    last_day = merged["order_date_parsed"].max()
    this_start = last_day - timedelta(days=6)
    prev_start = this_start - timedelta(days=7)
    prev_end = this_start - timedelta(days=1)

    this_week = merged[(merged["order_date_parsed"] >= this_start) &
                       (merged["order_date_parsed"] <= last_day)]
    last_week = merged[(merged["order_date_parsed"] >= prev_start) &
                       (merged["order_date_parsed"] <= prev_end)]

    def headline(frame: pd.DataFrame) -> dict:
        return {
            "revenue": float(frame["revenue"].sum()),
            "orders": int(len(frame)),
            "units": int(frame["qty_num"].sum()),
            "accounts": int(frame["account_id"].nunique()),
            "aov": float(frame["revenue"].sum() / len(frame)) if len(frame) else 0.0,
        }

    now, before = headline(this_week), headline(last_week)

    def delta(key: str) -> float:
        return (now[key] - before[key]) / before[key] if before[key] else 0.0

    by_product = (this_week.groupby("product")
                  .agg(units=("qty_num", "sum"), orders=("order_id", "count"),
                       revenue=("revenue", "sum"))
                  .sort_values("revenue", ascending=False).reset_index())
    total_revenue = float(this_week["revenue"].sum()) or 1.0
    by_product["share"] = by_product["revenue"] / total_revenue

    by_region = (this_week.groupby("region")
                 .agg(revenue=("revenue", "sum"), orders=("order_id", "count"),
                      accounts=("account_id", "nunique"))
                 .sort_values("revenue", ascending=False).reset_index())

    workbook = Workbook()
    summary = workbook.active
    summary.title = "Summary"

    summary["A1"] = "Weekly Revenue Report"
    summary["A1"].font = TITLE_FONT
    summary["A2"] = f"Period: {this_start:%d %b %Y} - {last_day:%d %b %Y}"
    summary["A2"].font = Font(italic=True, color="595959")

    summary_rows = [
        ("Revenue", now["revenue"], before["revenue"], delta("revenue"), MONEY),
        ("Orders", now["orders"], before["orders"], delta("orders"), COUNT),
        ("Units sold", now["units"], before["units"], delta("units"), COUNT),
        ("Active accounts", now["accounts"], before["accounts"], delta("accounts"), COUNT),
        ("Average order value", now["aov"], before["aov"], delta("aov"), MONEY),
    ]
    for column, header in enumerate(["Metric", "This week", "Last week", "Change"], start=1):
        summary.cell(row=4, column=column, value=header)
    style_header(summary, 4, 4)
    for offset, (label, current, previous, change, number_format) in enumerate(summary_rows, start=1):
        row = 4 + offset
        summary.cell(row=row, column=1, value=label).font = LABEL_FONT
        summary.cell(row=row, column=2, value=current).number_format = number_format
        summary.cell(row=row, column=3, value=previous).number_format = number_format
        summary.cell(row=row, column=4, value=change).number_format = '0.0%'
        for column in range(1, 5):
            summary.cell(row=row, column=column).border = BORDER
    autosize(summary, {1: 24, 2: 16, 3: 16, 4: 12})
    summary.freeze_panes = "A5"

    product_sheet = workbook.create_sheet("By product")
    product_sheet["A1"] = "Revenue by product"
    product_sheet["A1"].font = TITLE_FONT
    last_row = write_table(product_sheet, 3,
                           ["Product", "Units", "Orders", "Revenue", "Share of revenue"],
                           [[r.product, int(r.units), int(r.orders), float(r.revenue), float(r.share)]
                            for r in by_product.itertuples()],
                           formats={2: COUNT, 3: COUNT, 4: MONEY, 5: '0.0%'})
    autosize(product_sheet, {1: 26, 2: 10, 3: 10, 4: 14, 5: 18})
    product_sheet.freeze_panes = "A4"

    region_sheet = workbook.create_sheet("By region")
    region_sheet["A1"] = "Revenue by region"
    region_sheet["A1"].font = TITLE_FONT
    write_table(region_sheet, 3,
                ["Region", "Revenue", "Orders", "Accounts"],
                [[r.region, float(r.revenue), int(r.orders), int(r.accounts)]
                 for r in by_region.itertuples()],
                formats={2: MONEY, 3: COUNT, 4: COUNT})
    autosize(region_sheet, {1: 22, 2: 14, 3: 10, 4: 12})
    chart = BarChart()
    chart.title = "Revenue by region"
    chart.y_axis.title = "Revenue"
    chart.add_data(Reference(region_sheet, min_col=2, min_row=3,
                             max_row=3 + len(by_region)), titles_from_data=True)
    chart.set_categories(Reference(region_sheet, min_col=1, min_row=4,
                                   max_row=3 + len(by_region)))
    chart.height, chart.width = 8, 16
    region_sheet.add_chart(chart, f"F3")

    issues_sheet = workbook.create_sheet("Data issues")
    issues_sheet["A1"] = "Rows that need a human look"
    issues_sheet["A1"].font = TITLE_FONT
    issues_sheet["A2"] = "Nothing was silently dropped - everything skipped or flagged is listed here."
    issues_sheet["A2"].font = Font(italic=True, color="595959")
    write_table(issues_sheet, 4, ["Reference", "Issue", "Detail"], [list(i) for i in issues] or
                [["-", "no issues found", "-"]])
    autosize(issues_sheet, {1: 20, 2: 30, 3: 40})
    issues_sheet.freeze_panes = "A5"

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / f"weekly_report_{last_day:%Y-%m-%d}.xlsx"
    workbook.save(path)

    logger.info("period %s -> %s", this_start.date(), last_day.date())
    logger.info("revenue %.2f (%+.1f%% vs last week) | orders %s | accounts %s",
                now["revenue"], delta("revenue") * 100, now["orders"], now["accounts"])
    logger.info("data issues flagged: %s", len(issues))
    logger.info("report written to %s", path.resolve())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
