# Weekly Report Automation — two raw exports in, one formatted report out

Replaces the Monday-morning copy-paste: it takes the orders export and the CRM export, matches them
on the shared account id, works out the week-over-week numbers, and writes a formatted Excel report —
including a sheet listing the rows that need a human look.

## What it handles

- **Messy real-world input**: two different date formats in the same file, blank quantities, orders
  whose account no longer exists in the CRM
- **Nothing is dropped silently** — every row that was skipped or flagged is listed on the
  `Data issues` sheet with the reason
- **Formatted output**: styled headers, real number formats, frozen panes, a region chart
- **Self-check**: `verify_output.py` re-reads the finished workbook and compares the numbers against
  a fresh calculation from the raw files

## Run it

```bash
pip install -r requirements.txt
python make_sample_exports.py     # writes the two raw exports (seeded, deterministic)
python build_report.py            # -> output/weekly_report_2026-10-11.xlsx
python verify_output.py           # re-reads the file and checks the numbers
```

## Output

One workbook with four sheets:

| Sheet | Contents |
|---|---|
| `Summary` | revenue / orders / units / active accounts / AOV — this week vs last week and % change |
| `By product` | units, orders, revenue and share of revenue per product |
| `By region` | revenue, orders and account count per region, with a bar chart |
| `Data issues` | every skipped or flagged row, with the reason |

Verified run (2026-10-15): summary figures **21,519.69 revenue / 17 orders / 89 units / 9 accounts**;
`verify_output.py` → *all checks passed*.

## Time saved

The manual version of this — pull two exports, match by hand, rebuild the pivot, redo the formatting —
runs about **40–60 minutes** every week. This runs in **under 3 seconds**.

## Using it on real data

Point the column names at the top of `build_report.py` at your own export headers; the rest of the
logic does not change.
