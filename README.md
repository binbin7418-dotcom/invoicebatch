# InvoiceBatch

Batch-generate invoice PDFs from one spreadsheet of invoice lines, with a
full amount reconciliation and an independent re-check of every total.

## What it does

- Input: `.csv` / `.xlsx` with one row per invoice line —
  `invoice_no, date, client, description, qty, unit_price, tax_rate` and an
  optional `stated_total` (the total your sheet *claims* per invoice).
- Recomputes every line amount, subtotal, tax and total with Decimal
  arithmetic. If a stated total disagrees with its own lines, the
  recomputed value is used and the correction is logged.
- Normalizes dates to ISO (`DD/MM/YYYY` and `5 Mar 2026` supported).
- Outputs one PDF per invoice, a combined `all_invoices.pdf`, and
  `reconciliation_report.md`.
- Reports — never silently changes — duplicate invoice numbers used for
  different clients, missing client names, negative quantities, and gaps
  in the invoice number sequence.
- Re-check: totals are recomputed a second time from the source rows and
  compared against the generated invoices to the cent.

100% local. No account, no upload, no subscription.

## Run from source

```bash
python invoice_batch.py sample_lines.csv out/
python gui.py
```

## Acceptance

`test/acceptance_test.py` runs the seeded sample (9 lines, 6 invoices,
5 planted defects) and asserts the exact fix/issue counts, the grand
total and the PDF output. See `test/ACCEPTANCE.md`.
