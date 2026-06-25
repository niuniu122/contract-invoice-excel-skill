---
name: contract-invoice-excel
description: Generate Hong Kong Wangjing contract and stamped invoice deliverables from a sales-detail Excel workbook. Use when the user provides or mentions an .xlsx file containing sales order rows and asks to make, regenerate, batch produce, or fix contract/invoice outputs, including one-sheet workbooks, manual page breaks, A4 print-check PDFs, invoice seals/stamps, totals, and generation manifests.
---

# Contract Invoice Excel

## Purpose

Turn one sales-detail `.xlsx` into the complete deliverable set:

- one contract workbook with all contracts in a single worksheet
- one invoice workbook with all invoices in a single worksheet and one seal per invoice
- print-check PDFs for both workbooks
- `生成清单.txt` with counts, sums, formulas, stamps, and PDF page verification

Use the bundled script whenever possible:

```powershell
python "C:\Users\Administrator\.codex\skills\contract-invoice-excel\scripts\generate_contract_invoice_outputs.py" "<source.xlsx>"
```

The script defaults to the current workspace pattern: source files under `主体表格`, outputs under `产物`, and templates/seal copied from the prior corrected output. If defaults are unavailable, pass explicit paths:

```powershell
python "C:\Users\Administrator\.codex\skills\contract-invoice-excel\scripts\generate_contract_invoice_outputs.py" "<source.xlsx>" `
  --contract-template "<corrected_contract.xlsx>" `
  --invoice-template "<corrected_invoice.xlsx>" `
  --seal "<transparent_seal.png>"
```

## Input Contract

The source workbook should use the first non-empty worksheet and these columns by position:

`A 单据编号`, `B Customer`, `C Order Date`, `D Product Name`, `E Monetary Unit`, `F Quantity Unit`, `G Quantity`, `H Unit Price`, `I Amount`, `J Salesperson`, `K Delivery Address`, `L Total行数`.

Group rows by consecutive `单据编号`. `Total行数` must equal the actual consecutive row count for that order. Stop and report the script error if validation fails; do not silently deduplicate rows.

## Generation Rules

- Produce one document per sales order, not one document per row.
- For multi-line orders, merge customer/date fields vertically and write a `SUM` formula in the total row.
- Keep all contracts in one worksheet named `合同`; keep all invoices in one worksheet named `发票`.
- Insert a manual page break after every document so print preview never mixes two orders on one page.
- Use fixed A4 print scale `70%`, not automatic fit-to-page.
- Set print title rows to `1:5`.
- Use the transparent seal image for invoices. The script normalizes every Excel shape with a default `--stamp-offset 45`, placing the stamp low enough to overlap both the footer line and the `HONG KONG...` company text.
- The output directory is timestamped when an existing output folder already exists, preventing WPS/Excel lock-file problems and stale-view overwrites.

## Verification Checklist

After running the script, verify the newest output directory:

```powershell
Get-ChildItem -LiteralPath "<output-dir>" -Force | Select-Object Name,Length,LastWriteTime
Get-Content -LiteralPath "<output-dir>\生成清单.txt" -Encoding UTF8 -Raw
```

The required checks are:

- source row count equals generated data row count
- source `Amount` sum equals generated contract and invoice amount sums
- document count equals PDF page count for both contract and invoice
- invoice image count equals document count
- manual page breaks equal document count minus one
- total `SUM` formula count equals number of multi-line orders

If the user asks about seal placement, do not rely only on workbook XML. Open/export the PDF and inspect a rendered page, or use Excel/WPS visual coordinates. The desired placement has the stamp overlapping the footer horizontal line and the `HONG KONG WIN KING PAPER CO.，LIMITED` footer text.

## Common Fixes

- If WPS has an old workbook open, close it or open the newly timestamped output folder. Do not claim an old file changed while WPS is still displaying it.
- If the stamp is too high, rerun with a larger offset, for example `--stamp-offset 55`.
- If Excel COM is unavailable, rerun with `--no-pdf` only as a fallback and tell the user PDFs/stamp normalization were skipped.
- If templates are missing, locate the previous corrected contract workbook, invoice workbook, and transparent seal PNG, then pass them explicitly.
