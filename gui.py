#!/usr/bin/env python3
"""InvoiceBatch desktop shell (tkinter).
Pick a spreadsheet of invoice lines -> generate invoice PDFs + reconciliation report."""
import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from invoice_batch import run_invoicebatch


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("InvoiceBatch — Invoice PDFs & Reconciliation")
        self.geometry("680x540")
        self.src = tk.StringVar()
        self.seller = tk.StringVar()
        tk.Label(self, text="Invoice lines file (.csv / .xlsx) — one row per invoice line:").pack(anchor="w", padx=12, pady=(12, 2))
        row = tk.Frame(self)
        row.pack(fill="x", padx=12)
        tk.Entry(row, textvariable=self.src).pack(side="left", fill="x", expand=True)
        tk.Button(row, text="Browse…", command=self.pick).pack(side="left", padx=6)
        tk.Label(self, text="Your business name (printed on each invoice, optional):").pack(anchor="w", padx=12, pady=(10, 2))
        tk.Entry(self, textvariable=self.seller).pack(fill="x", padx=12)
        self.btn = tk.Button(self, text="Generate Invoices & Reconcile", command=self.run, height=2)
        self.btn.pack(pady=12)
        self.log = scrolledtext.ScrolledText(self, wrap="word")
        self.log.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        self.write("InvoiceBatch turns a spreadsheet of invoice lines into one PDF per invoice,\n"
                   "recomputes every total from the lines, and re-checks the arithmetic a second\n"
                   "time before you send anything. Everything runs locally.\n"
                   "Columns: invoice_no, date, client, description, qty, unit_price, tax_rate, stated_total (optional).\n")

    def write(self, s):
        self.log.insert("end", s)
        self.log.see("end")

    def pick(self):
        p = filedialog.askopenfilename(filetypes=[("Spreadsheet files", "*.csv *.xlsx *.xlsm"), ("All files", "*.*")])
        if p:
            self.src.set(p)

    def run(self):
        src = self.src.get().strip()
        if not src or not os.path.isfile(src):
            messagebox.showerror("InvoiceBatch", "Choose an invoice lines file first.")
            return
        self.btn.config(state="disabled")
        threading.Thread(target=self._work, args=(src,), daemon=True).start()

    def _work(self, src):
        try:
            outdir = os.path.join(os.path.dirname(os.path.abspath(src)), "invoicebatch_output")
            s = run_invoicebatch(src, outdir, seller=self.seller.get().strip())
            self.write(f"\nDone. {s['invoices']} invoices from {s['rows']} lines.\n"
                       f"Corrected automatically: {s['fixed']}  ·  Needs your decision: {s['issues']} "
                       f"({s['errors']} blocking)  ·  Re-check mismatches: {s['recheck_problems']}\n"
                       f"Grand total across all invoices: {s['grand_total']}\n"
                       f"PDFs: {os.path.join(outdir, 'pdfs')}\nReport: {s['report_path']}\n")
        except Exception as e:
            self.write(f"\nError: {e}\n")
        finally:
            self.btn.config(state="normal")


if __name__ == "__main__":
    App().mainloop()
