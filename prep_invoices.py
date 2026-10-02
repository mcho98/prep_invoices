#!/usr/bin/env python3
"""Turn the timesheet workbook into an invoice-import CSV (same columns as
sample_invoice_import.csv).

One invoice per client. Each day and service with hours becomes one line:
  Customer             = client name (first line of the invoice only)
  Item(Product/Service)= Homemaking / Personal Care / Respite
  ItemDescription      = the date (YYYY-MM-DD)
  ItemQuantity         = hours
  ItemRate             = the client's $/hr
  ItemAmount           = hours x rate, rounded to cents

Workbook layout: one date row near the top (first date cell starts the month,
one column per day), then one section per client starting at a "Client Name"
label with the name to its right, holding a "rate $/hr" cell and the rows
"Homemaking hours", "Personal Care Hours", "Respite hours".

Usage:
  python prep_invoices.py mock_data.xlsx
  python prep_invoices.py mock_data.xlsx -o invoices.csv --invoice-date 2026-08-31 --terms-days 30 --first-invoice-no 1001
"""
import argparse
import csv
import datetime as dt
import sys
from decimal import Decimal, ROUND_HALF_UP

import openpyxl

SERVICES = {
    "homemaking hours": "Homemaking",
    "personal care hours": "Personal Care",
    "respite hours": "Respite",
}
CLIENT_LABEL = "client name"
RATE_LABEL = "rate $/hr"
CENT = Decimal("0.01")

HEADER = ["*InvoiceNo", "*Customer", "*InvoiceDate", "*DueDate", "Terms", "Location", "Memo",
          "Item(Product/Service)", "ItemDescription", "ItemQuantity", "ItemRate", "*ItemAmount",
          "ItemTaxAmount"]


def norm(value):
    return value.strip().lower() if isinstance(value, str) else None


def as_date(value):
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    return None


def find_date_columns(ws):
    """[(column, date)] for each day of the month, starting from the first date cell."""
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 20)):
        for cell in row:
            start = as_date(cell.value)
            if start is None:
                continue
            columns, day, col = [], start, cell.column
            while day.month == start.month:
                columns.append((col, day))
                col += 1
                day += dt.timedelta(days=1)
            return columns
    sys.exit("Could not find the date row (no date cell in the first 20 rows).")


def find_clients(ws):
    starts = []
    for row in ws.iter_rows():
        for cell in row:
            if norm(cell.value) == CLIENT_LABEL:
                name = ws.cell(cell.row, cell.column + 1).value
                starts.append((cell.row, str(name).strip() if name else f"(unnamed, row {cell.row})"))
    return [(name, row, starts[i + 1][0] - 1 if i + 1 < len(starts) else ws.max_row)
            for i, (row, name) in enumerate(starts)]


def read_client(ws, name, first_row, last_row, date_columns):
    """Return (rate, [(date, service, hours)]). Exits on a missing rate or bad hours."""
    rate, service_rows = None, {}
    for row in ws.iter_rows(min_row=first_row, max_row=last_row):
        for cell in row:
            label = norm(cell.value)
            if label == RATE_LABEL and rate is None:
                rate = ws.cell(cell.row, cell.column + 1).value
            elif label in SERVICES:
                service_rows.setdefault(cell.row, SERVICES[label])
    lines = []
    for row, service in sorted(service_rows.items()):
        for col, day in date_columns:
            cell = ws.cell(row, col)
            if cell.value in (None, "", 0):
                continue
            if not isinstance(cell.value, (int, float)) or not 0 < cell.value <= 24:
                sys.exit(f"{name}: {cell.coordinate} has invalid hours ({cell.value!r}).")
            lines.append((day, service, cell.value))
    if lines and not (isinstance(rate, (int, float)) and rate > 0):
        sys.exit(f"{name}: no valid hourly rate next to '{RATE_LABEL}' (found {rate!r}).")
    lines.sort(key=lambda l: (l[0], l[1]))
    return rate, lines


def fmt(n):
    """Number without trailing zeros (3.5 -> '3.5', 34.03 -> '34.03', 4.0 -> '4')."""
    return f"{n:f}".rstrip("0").rstrip(".") if isinstance(n, Decimal) else f"{n:g}"


def build_rows(ws, invoice_date, due_date, terms, first_no):
    date_columns = find_date_columns(ws)
    rows, number = [], first_no
    for name, first, last in find_clients(ws):
        rate, lines = read_client(ws, name, first, last, date_columns)
        if not lines:
            print(f"Skipping {name}: no hours.", file=sys.stderr)
            continue
        total = Decimal(0)
        for i, (day, service, hours) in enumerate(lines):
            amount = (Decimal(str(hours)) * Decimal(str(rate))).quantize(CENT, rounding=ROUND_HALF_UP)
            total += amount
            head = [number, name, invoice_date.strftime("%d/%m/%Y"), due_date.strftime("%d/%m/%Y"), terms,
                    "", ""] if i == 0 else [number, "", "", "", "", "", ""]
            rows.append(head + [service, day.strftime("%Y-%m-%d"), fmt(hours), fmt(rate), f"{amount:.2f}", ""])
        print(f"Invoice {number}: {name}, {len(lines)} lines, ${total:,.2f}", file=sys.stderr)
        number += 1
    return rows


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("workbook")
    p.add_argument("-o", "--output", default="invoice_import.csv")
    p.add_argument("--invoice-date", type=dt.date.fromisoformat, default=dt.date.today(),
                   help="YYYY-MM-DD (default: today)")
    p.add_argument("--terms-days", type=int, default=30, help="days until due (default 30; 0 = Due on receipt)")
    p.add_argument("--first-invoice-no", type=int, default=1001)
    args = p.parse_args()

    ws = openpyxl.load_workbook(args.workbook).active
    due = args.invoice_date + dt.timedelta(days=args.terms_days)
    terms = f"Net {args.terms_days}" if args.terms_days else "Due on receipt"
    rows = build_rows(ws, args.invoice_date, due, terms, args.first_invoice_no)
    with open(args.output, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(HEADER)
        w.writerows(rows)
    print(f"Wrote {len(rows)} lines to {args.output}", file=sys.stderr)


if __name__ == "__main__":
    main()
