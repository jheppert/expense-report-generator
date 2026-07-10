# Expense Report Builder

A single-file, browser-based tool that turns two monthly invoices (Adobe Photoshop + Sketch) into a finished consultant expense report — no manual data entry.

Drop the two PDFs in, and it reads the purchase dates and amounts, fills them into the report template, embeds both receipts, and downloads a completed `.xlsx`.

## What it does

- Reads the **purchase date** and **amount** from each invoice (identified by content, so drop order doesn't matter)
- Fills the detail table, lands the combined total on the correct week's Friday, and lets Excel recalculate every downstream total on open
- Embeds both receipt images **side by side** on the Receipts sheet
- Keeps the original template's formatting, logos, and banner intact
- Names the file for the month automatically

Everything is editable before you generate — detected values show in two slots, and anything it couldn't read is highlighted for you to fill in.

## Use it

**Locally:** open `expense-report-tool.html` in any modern browser and drop in the two PDFs.

**Hosted (GitHub Pages):** see below.

Needs an internet connection on load — it pulls three small libraries (pdf.js, ExcelJS, and web fonts) from a CDN. Your invoices never leave your machine; all parsing and file generation happen in the browser.

## Hosting on GitHub Pages

1. For the tool to serve at the site root, rename the file to `index.html` (otherwise it's reachable at `/expense-report-tool.html`).
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
- Built for the recurring Adobe + Sketch pair. Extending it to other vendors is a small change — see `context.md`.
