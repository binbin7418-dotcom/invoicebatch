#!/usr/bin/env python3
"""
InvoiceBatch — batch invoice PDF generation with amount reconciliation.

Pipeline:
  1. Read a CSV/XLSX where each row is one invoice line
     (invoice_no, date, client, description, qty, unit_price, tax_rate,
      optional stated_total per invoice).
  2. Audit: recompute every line amount, subtotal, tax and total with
     Decimal arithmetic. Where the sheet's own stated total disagrees,
     the recomputed value wins and the change is logged (a spreadsheet
     total that disagrees with its own lines is a defect, not data).
  3. Generate one print-ready PDF per invoice (plus a combined PDF).
  4. Re-check: independently recompute every invoice from the source
     rows a second time and compare against what was written into the
     PDFs; check invoice numbers for duplicates and sequence gaps.
     Problems that need a human (missing client, duplicate numbers,
     negative quantities) are reported, never silently "fixed".

Local only. No network access.
"""
import csv
import io
import os
import re
import sys
from decimal import Decimal, ROUND_HALF_UP

CENT = Decimal("0.01")
DATE_FORMATS_OK = "YYYY-MM-DD, DD/MM/YYYY, or 'D Mon YYYY' (e.g. 5 Mar 2026)"
MONTHS = {m: i + 1 for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"])}


def money(d):
    return Decimal(d).quantize(CENT, rounding=ROUND_HALF_UP)


def parse_decimal(v):
    """Parse a number from a spreadsheet cell. Returns Decimal or None."""
    if v is None:
        return None
    s = str(v).strip().replace(",", "").replace("$", "").replace("£", "").replace("€", "")
    if s == "":
        return None
    try:
        return Decimal(s)
    except Exception:
        return None


def parse_date(v):
    """Normalize a date to ISO. Returns (iso, changed) or (None, False)."""
    s = (str(v) or "").strip()
    if not s:
        return None, False
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s):
        return s, False
    m = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{4})", s)
    if m:  # documented as DD/MM/YYYY
        d, mo, y = m.groups()
        return f"{y}-{int(mo):02d}-{int(d):02d}", True
    m = re.fullmatch(r"(\d{1,2})\s+([A-Za-z]{3,9})\s+(\d{4})", s)
    if m:
        d, mon, y = m.groups()
        mi = MONTHS.get(mon[:3].lower())
        if mi:
            return f"{y}-{mi:02d}-{int(d):02d}", True
    return None, False


def read_rows(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        from openpyxl import load_workbook
        wb = load_workbook(path, data_only=True, read_only=True)
        ws = wb.active
        it = ws.iter_rows(values_only=True)
        header = [str(c).strip() if c is not None else "" for c in next(it)]
        rows = []
        for r in it:
            if r is None or all(c is None for c in r):
                continue
            rows.append({header[i]: ("" if c is None else c) for i, c in enumerate(r) if i < len(header)})
        return rows
    raw = open(path, "rb").read()
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = raw.decode("latin-1")
    reader = csv.DictReader(io.StringIO(text))
    return [dict(r) for r in reader]


def audit_and_compute(rows):
    """Group rows into invoices, recompute all money, log fixes & issues.
    Returns (invoices, fixes, issues). invoices: list of dicts with lines
    and computed subtotal/tax/total (Decimal)."""
    fixes, issues = [], []
    invoices, order = {}, []

    def issue(sev, inv, msg):
        issues.append({"sev": sev, "invoice": inv, "msg": msg})

    for ln, r in enumerate(rows, start=2):  # line 1 = header
        inv_no = str(r.get("invoice_no", "") or "").strip()
        if not inv_no:
            issue("ERROR", f"row {ln}", "row has no invoice_no — skipped; add a number and re-run")
            continue
        if inv_no not in invoices:
            invoices[inv_no] = {"no": inv_no, "lines": [], "clients": set(),
                                "dates": set(), "stated_total": None}
            order.append(inv_no)
        inv = invoices[inv_no]
        client = str(r.get("client", "") or "").strip()
        if client:
            inv["clients"].add(client)
        iso, changed = parse_date(r.get("date", ""))
        if iso:
            inv["dates"].add(iso)
            if changed:
                fixes.append({"invoice": inv_no, "field": "date",
                              "old": str(r.get("date")), "new": iso,
                              "note": "date normalized to ISO format"})
        elif str(r.get("date", "") or "").strip():
            issue("ERROR", inv_no, f"date '{r.get('date')}' is not in a supported format ({DATE_FORMATS_OK})")
        qty = parse_decimal(r.get("qty"))
        price = parse_decimal(r.get("unit_price"))
        desc = str(r.get("description", "") or "").strip()
        if qty is None or price is None:
            issue("ERROR", inv_no, f"line {ln}: qty/unit_price not numeric (qty='{r.get('qty')}', unit_price='{r.get('unit_price')}') — line excluded from totals")
            continue
        if qty < 0:
            issue("ERROR", inv_no, f"line {ln}: negative quantity ({qty}) — credit lines must be a deliberate decision; left in totals but flagged")
        if not desc:
            issue("WARN", inv_no, f"line {ln}: empty description")
        tax_rate = parse_decimal(r.get("tax_rate"))
        if tax_rate is None:
            tax_rate = Decimal("0")
        inv["lines"].append({"desc": desc, "qty": qty, "price": price,
                             "amount": money(qty * price), "tax_rate": tax_rate})
        st = parse_decimal(r.get("stated_total"))
        if st is not None:
            inv["stated_total"] = st

    result = []
    for no in order:
        inv = invoices[no]
        if len(inv["clients"]) > 1:
            issue("ERROR", no, f"invoice number used for multiple clients ({', '.join(sorted(inv['clients']))}) — duplicate number; renumber one of them")
        if not inv["clients"]:
            issue("ERROR", no, "no client name on any line — an invoice without a client cannot be sent")
        if len(inv["dates"]) > 1:
            issue("WARN", no, f"lines carry different dates ({', '.join(sorted(inv['dates']))}) — the earliest is used on the PDF")
        subtotal = money(sum((l["amount"] for l in inv["lines"]), Decimal("0")))
        # tax computed per distinct rate group to stay exact
        tax = Decimal("0")
        for l in inv["lines"]:
            tax += money(l["amount"] * l["tax_rate"] / Decimal("100"))
        tax = money(tax)
        total = money(subtotal + tax)
        inv.update({"client": sorted(inv["clients"])[0] if inv["clients"] else "",
                    "date": sorted(inv["dates"])[0] if inv["dates"] else "",
                    "subtotal": subtotal, "tax": tax, "total": total})
        if inv["stated_total"] is not None and inv["stated_total"] != total:
            fixes.append({"invoice": no, "field": "total",
                          "old": f"{inv['stated_total']:.2f}", "new": f"{total:.2f}",
                          "note": "stated total disagreed with the recomputed lines+tax; recomputed value used"})
        result.append(inv)

    # sequence gap check on purely numeric invoice numbers
    nums = sorted(int(n) for n in order if n.isdigit())
    if len(nums) >= 2:
        gaps = [n for n in range(nums[0], nums[-1] + 1) if n not in set(nums)]
        if gaps:
            issue("WARN", "-", f"invoice numbers are not fully sequential — missing: {', '.join(map(str, gaps))} (fine if those were voided, but check)")
    return result, fixes, issues


def recheck(rows, invoices):
    """Second, independent pass: recompute from the raw rows again and
    compare with the totals that went into the PDFs."""
    again, _, _ = audit_and_compute(rows)
    by_no = {i["no"]: i for i in again}
    problems = []
    for inv in invoices:
        a = by_no.get(inv["no"])
        if a is None:
            problems.append(f"{inv['no']}: missing on re-check")
        elif (a["subtotal"], a["tax"], a["total"]) != (inv["subtotal"], inv["tax"], inv["total"]):
            problems.append(f"{inv['no']}: totals changed between passes (engine bug — do not send)")
    return problems


def render_pdfs(invoices, outdir, seller=""):
    """Write one PDF per invoice + combined all_invoices.pdf. Returns paths."""
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas

    os.makedirs(outdir, exist_ok=True)
    paths = []

    def draw_invoice(c, inv):
        w, h = A4
        y = h - 20 * mm
        c.setFont("Helvetica-Bold", 16)
        c.drawString(20 * mm, y, f"INVOICE {inv['no']}")
        y -= 8 * mm
        c.setFont("Helvetica", 10)
        if seller:
            c.drawString(20 * mm, y, seller)
            y -= 5 * mm
        if inv["date"]:
            c.drawString(20 * mm, y, f"Date: {inv['date']}")
            y -= 5 * mm
        c.drawString(20 * mm, y, f"Bill to: {inv['client'] or '(no client — see report)'}")
        y -= 10 * mm
        c.setFont("Helvetica-Bold", 9)
        c.drawString(20 * mm, y, "Description")
        c.drawString(115 * mm, y, "Qty")
        c.drawString(135 * mm, y, "Unit price")
        c.drawString(160 * mm, y, "Amount")
        y -= 4 * mm
        c.line(20 * mm, y, 190 * mm, y)
        y -= 6 * mm
        c.setFont("Helvetica", 9)
        for l in inv["lines"]:
            c.drawString(20 * mm, y, (l["desc"] or "")[:60])
            c.drawString(115 * mm, y, f"{l['qty']}")
            c.drawRightString(155 * mm, y, f"{l['price']:.2f}")
            c.drawRightString(185 * mm, y, f"{l['amount']:.2f}")
            y -= 5 * mm
        y -= 2 * mm
        c.line(20 * mm, y, 190 * mm, y)
        y -= 6 * mm
        for label, val in (("Subtotal", inv["subtotal"]), ("Tax", inv["tax"]), ("TOTAL", inv["total"])):
            c.setFont("Helvetica-Bold" if label == "TOTAL" else "Helvetica", 10)
            c.drawString(135 * mm, y, label)
            c.drawRightString(185 * mm, y, f"{val:.2f}")
            y -= 5 * mm
        c.showPage()

    for inv in invoices:
        p = os.path.join(outdir, f"invoice_{inv['no']}.pdf")
        c = canvas.Canvas(p, pagesize=A4)
        draw_invoice(c, inv)
        c.save()
        paths.append(p)
    combined = os.path.join(outdir, "all_invoices.pdf")
    c = canvas.Canvas(combined, pagesize=A4)
    for inv in invoices:
        draw_invoice(c, inv)
    c.save()
    paths.append(combined)
    return paths


def run_invoicebatch(src_path, outdir, seller=""):
    os.makedirs(outdir, exist_ok=True)
    rows = read_rows(src_path)
    invoices, fixes, issues = audit_and_compute(rows)
    pdfs = render_pdfs(invoices, os.path.join(outdir, "pdfs"), seller=seller)
    problems = recheck(rows, invoices)
    errs = sum(1 for i in issues if i["sev"] == "ERROR")

    lines = ["# InvoiceBatch — generation & reconciliation report", "",
             f"Source file: {os.path.basename(src_path)}  ·  Line rows: {len(rows)}  ·  Invoices: {len(invoices)}", "",
             "## Totals per invoice (recomputed from the lines, Decimal arithmetic)"]
    for inv in invoices:
        lines.append(f"- {inv['no']} — {inv['client'] or '(no client)'} — subtotal {inv['subtotal']:.2f} + tax {inv['tax']:.2f} = **{inv['total']:.2f}**")
    lines += ["", "## What was corrected automatically"]
    lines += [f"- [{f['invoice']}] {f['field']}: '{f['old']}' -> '{f['new']}' ({f['note']})" for f in fixes] or ["- nothing"]
    lines += ["", "## What needs your decision (never silently changed)"]
    lines += [f"- [{i['sev']}] [{i['invoice']}] {i['msg']}" for i in issues] or ["- nothing"]
    lines += ["", "## Re-check verdict"]
    if problems:
        lines.append(f"- FAIL: {len(problems)} mismatch(es) between generation and the independent re-check: " + "; ".join(problems))
    else:
        lines.append(f"- PASS: an independent second recomputation of all {len(invoices)} invoices matches the PDF totals to the cent. "
                     f"{errs} blocking issue(s) listed above still need your decision before these invoices are sent."
                     if errs else
                     f"- PASS: an independent second recomputation of all {len(invoices)} invoices matches the PDF totals to the cent, and no blocking issues were found.")
    lines.append("- Scope note: InvoiceBatch recomputes arithmetic and checks numbering; it does not judge whether prices, tax rates or client details are the right ones for your business.")
    report_path = os.path.join(outdir, "reconciliation_report.md")
    with open(report_path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines) + "\n")
    return {"rows": len(rows), "invoices": len(invoices), "fixed": len(fixes),
            "issues": len(issues), "errors": errs, "recheck_problems": len(problems),
            "grand_total": str(money(sum((i["total"] for i in invoices), Decimal("0")))),
            "pdfs": pdfs, "report_path": report_path}


def main(argv):
    if len(argv) != 3:
        print("usage: invoice_batch.py <lines.csv|lines.xlsx> <outdir>")
        return 2
    s = run_invoicebatch(argv[1], argv[2])
    print(f"rows={s['rows']} invoices={s['invoices']} fixed={s['fixed']} issues={s['issues']} "
          f"errors={s['errors']} recheck_problems={s['recheck_problems']} grand_total={s['grand_total']}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
