#!/usr/bin/env python3
"""
Creates the two raw exports the report job normally receives every week:

  sales_orders.csv   - pulled from the e-commerce / billing system
  crm_accounts.csv   - pulled from the CRM

They are synthetic (seeded, so they are identical on every run) but they keep the
mess real data has: a couple of orders whose account is missing from the CRM,
one account with no orders, mixed date formats and a few blank quantities.

    python make_sample_exports.py
"""

import csv
import random
from datetime import date, timedelta
from pathlib import Path

HERE = Path(__file__).resolve().parent
random.seed(20261011)

PRODUCTS = [
    ("SKU-101", "Starter Plan"),
    ("SKU-102", "Team Plan"),
    ("SKU-201", "Onboarding Package"),
    ("SKU-202", "Priority Support"),
    ("SKU-301", "Data Migration"),
    ("SKU-302", "Custom Integration"),
]
ACCOUNTS = [
    ("ACC-001", "Northwind Trading", "Ana Ruiz", "Spain", "Team"),
    ("ACC-002", "Blue Harbor Logistics", "Tom Becker", "Germany", "Team"),
    ("ACC-003", "Kestrel Analytics", "Lian Chen", "Singapore", "Starter"),
    ("ACC-004", "Riverbend Foods", "Marta Silva", "Portugal", "Team"),
    ("ACC-005", "Pinewood Studio", "Jonas Meyer", "Germany", "Starter"),
    ("ACC-006", "Orchard Health", "Priya Nair", "India", "Enterprise"),
    ("ACC-007", "Atlas Freight", "Sofia Rossi", "Italy", "Starter"),
    ("ACC-008", "Lumen Retail", "David Kim", "South Korea", "Team"),
    ("ACC-009", "Harbor & Vale", "Grace Okafor", "United Kingdom", "Enterprise"),
    ("ACC-010", "Sequoia Labs", "Ivan Petrov", "Netherlands", "Starter"),
]


def main() -> None:
    start = date(2026, 9, 14)
    orders = []
    for index in range(72):
        order_day = start + timedelta(days=random.randint(0, 27))
        account = random.choice(ACCOUNTS)
        sku, product = random.choice(PRODUCTS)
        # a handful of malformed rows, on purpose
        account_id = account[0]
        if index in (17, 41):
            account_id = "ACC-999"          # account deleted from the CRM
        qty = "" if index == 33 else str(random.randint(1, 9))
        unit_price = f"{random.uniform(45, 480):.2f}"
        stamp = (order_day.strftime("%Y-%m-%d") if index % 7 else
                 order_day.strftime("%d/%m/%Y"))   # one system uses another format
        orders.append({
            "order_id": f"SO-{10000 + index}",
            "order_date": stamp,
            "account_id": account_id,
            "sku": sku,
            "product": product,
            "qty": qty,
            "unit_price": unit_price,
            "region": account[3],
        })

    with (HERE / "sales_orders.csv").open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(orders[0].keys()))
        writer.writeheader()
        writer.writerows(orders)

    accounts = [{"account_id": a[0], "company": a[1], "owner": a[2],
                 "country": a[3], "plan": a[4]} for a in ACCOUNTS]
    with (HERE / "crm_accounts.csv").open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(accounts[0].keys()))
        writer.writeheader()
        writer.writerows(accounts)

    print(f"wrote {len(orders)} orders and {len(accounts)} accounts to {HERE}")


if __name__ == "__main__":
    main()
