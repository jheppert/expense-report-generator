# Expense Report Builder

A single-file, browser-based tool that turns your monthly SaaS invoices (Adobe Photoshop and/or Sketch) into a finished consultant expense report — no manual data entry.

**→ Use it: <https://jheppert.github.io/expense-report-generator/>**

Drop the PDFs in — one or both — and it reads the purchase dates and amounts, fills them into the report template, embeds the receipts, and downloads a completed `.xlsx`.

## What it does

- Reads the **purchase date** and **amount** from each invoice (identified by content, so drop order doesn't matter)
- Works with **either invoice alone or both** — line items fill from the top of the detail table, so a single-invoice month has no blank row
- Fills the detail table, lands the combined total on the correct week's Friday, and lets Excel recalculate every downstream total on open
- Embeds the receipt images **side by side** on the Receipts sheet (a lone receipt sits on the left)
- Keeps the original template's formatting, logos, and banner intact
- Names the file for the month automatically

Everything is editable before you generate — detected values show in two slots (either one is optional), and anything it couldn't read is highlighted for you to fill in.

## Use it

**Hosted:** <https://jheppert.github.io/expense-report-generator/> — nothing to install, just drop in the PDFs.

**Locally:** open `index.html` in any modern browser and drop in the PDFs.

Needs an internet connection on load — it pulls three small libraries (pdf.js, ExcelJS, and web fonts) from a CDN. Your invoices never leave your machine; all parsing and file generation happen in the browser.

## Hosting on GitHub Pages

This repo is already deployed at <https://jheppert.github.io/expense-report-generator/>, serving `index.html` from `main` at the root. To set it up again, or in a fork:

1. Keep the tool named `index.html` so it serves at the site root.
2. Push to a GitHub repo.
3. In the repo, go to **Settings → Pages**.
4. Under **Source**, choose **Deploy from a branch**, pick your `main` branch and the `/ (root)` folder, and save.
5. After a minute the site is live at `https://<your-username>.github.io/<repo-name>/`.

CDN libraries and fonts load over HTTPS, so there are no mixed-content issues on Pages.

## Built with

- [pdf.js](https://mozilla.github.io/pdf.js/) — invoice text extraction and receipt rendering
- [ExcelJS](https://github.com/exceljs/exceljs) — reading and writing the `.xlsx` template
- Vanilla HTML/CSS/JS — no build step, no framework

## Notes

- Two full-width receipts sit side by side (~8.5in total), so the Receipts sheet is sized for on-screen viewing. If you print it, the pair fills the page width.
- Built for the recurring Adobe + Sketch pair, with either one optional. Extending it to other vendors is a small change — see `context.md`.
