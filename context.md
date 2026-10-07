# Project Context

Working notes for picking this project back up. Written July 2026.

## What this is

A single self-contained `index.html` (~256KB) that automates Jeff's monthly consultant expense report for Apple (via Red Oak Technologies). Two SaaS invoices used to come in every month — **Adobe Photoshop** and **Sketch** — and the report is otherwise identical month to month. Either one is now optional — Adobe was paused as of October 2026 and may come back — so the tool accepts one or both PDFs, extracts the date + amount from each, fills an embedded `.xlsx` template, embeds the receipt images, and downloads the finished file.

Everything runs in the browser. No backend, no build step. The report template is embedded in the HTML as a base64 string (`TEMPLATE_B64`).

## Where it lives

Hosted on GitHub Pages at <https://jheppert.github.io/expense-report-generator/> (served from `index.html` on `main`).

## Current status

Complete and validated. The full spreadsheet-generation path was tested against the real invoices and rendered to confirm it matches a correctly-filled report (all totals correct, zero formula errors). The only pieces not exercisable outside a real browser are the DOM/drag-drop and pdf.js canvas rendering, which are standard.

## The report template

Two sheets: **"Expense Report"** and **"Receipts"**.

Cells the tool writes on "Expense Report":

| Cell | Meaning |
|------|---------|
| `A10` | Week ending date (the Sunday ending a Mon–Sun week) |
| `F12` | Date under the FRI column of the weekly grid |
| `F27` | "Other Expenses" combined amount, sits in the FRI column — **this is the value that flows to all totals** |
| `A43` / `B43` / `F43` / `I43` | First line item — date / item name / purpose / amount |
| `A44` / `B44` / `F44` / `I44` | Second line item, or cleared when only one invoice is used |

The template bakes item names and purpose text into these rows (`B43` = "Adobe Photoshop", `F43` = "Photo editing for design work"; `B44`/`F44` the Sketch equivalents), **but the tool overwrites them per row** — see the single-invoice note under Fill logic. Rows 45–48 are spare and identically formatted, which is the room a third vendor would use.

The detail table does **not** auto-link to the weekly grid — `F27` is what drives `I27 → I29 → K12 → K14 → K18` ("Balance Due Employee"). That's why the tool sets `F27` directly as a value rather than relying on the detail rows. Verified: no formula anywhere on the sheet references rows 43–48, so moving an item between those rows can't disturb a total.

## Key decisions & gotchas

These are the non-obvious things that cost time to work out. Don't relearn them the hard way.

- **Engine is ExcelJS, not raw XML surgery.** ExcelJS preserves merged cells, shared formulas, number formats, fonts, borders, and pictures on round-trip. The tradeoff below made it the right call once the banner was handled.

- **ExcelJS drops vector shapes on save.** The template originally had the "EXPENSE REPORT" title as a vector textbox shape; it vanished on re-save. Fix: that banner was rasterized to a PNG (`#70AD47` green, bold-italic serif) and baked into the template as a *picture*, which ExcelJS preserves. If the template is ever rebuilt from a fresh Excel file, this conversion has to be redone.

- **`fullCalcOnLoad` must be set at generate time by mutation.** ExcelJS's parser *discards* `fullCalcOnLoad` on load, so baking it into the template doesn't survive. It also won't take if you replace the whole object. The working form is:
  ```js
  wb.calcProperties = wb.calcProperties || {};
  wb.calcProperties.fullCalcOnLoad = true;
  ```
  This is required: `F27` is written as a value, so the downstream formula caches are stale until Excel recalculates on open.

- **Dates are set at local noon** (`new Date(y, mo, d, 12, 0, 0)`) to avoid a timezone off-by-one when ExcelJS serializes to a date serial (Jeff is on Pacific). The date format shows only the date, so the noon time is invisible.

- **Side-by-side receipts use integer start columns.** ExcelJS's fractional-column → pixel-offset math uses a different column width than Excel actually renders with, so a fractional `col` (e.g. `6.35`) drifts and the images overlap in Excel even though the file looks fine to ExcelJS. Anchoring at an **integer** column (`col: 0` for Adobe, `col: 7` for Sketch) sidesteps this — integer column indices are interpreted identically by both. Col 7 clears the 408px-wide left image with a ~56px gap on the Receipts sheet's ~62px columns. Multi-page invoices stack downward within their own column.

- **Orphaned media bloats the file.** Clearing image *anchors* from a worksheet (`worksheet._media.length = 0`) leaves the actual image *files* in `xl/media/`. The old receipt JPGs had to be stripped from the template zip separately (`zip -d`). If the template is rebuilt, check `xl/media/` for orphans.

## PDF parsing

Vendor is detected by keyword in the extracted text (`"adobe"` / `"sketch"`), **not** the filename — so drop order and file names don't matter. If neither keyword is found, the file routes to the first open slot and the UI asks the user to confirm the vendor.

Dates handle both invoice formats:
- Adobe: `Invoice Date DD-MON-YYYY` (e.g. `09-JUL-2026`)
- Sketch: `Date of issue Month D, YYYY` (e.g. `July 3, 2026`)
- Plus generic fallbacks for each shape.

Amounts try, in order: `GRAND TOTAL (USD) N.NN` (Adobe) → `Amount due $N.NN` (Sketch) → `Total $N.NN`.

All of this lives in `parseInvoice()` in the HTML. If a future invoice uses a new layout, that's the function to extend.

## Fill logic

- **One or both invoices.** `presentSlots()` returns the filled slots in canonical Adobe-then-Sketch order, and everything downstream iterates that list — the generate gate (`readyToGenerate()`: at least one slot, every present slot complete), the detail rows, the total, and the receipt columns. A missing vendor is simply absent, not zero.
- **Detail rows fill from the top.** Items go into `DETAIL_ROWS` (43, 44) by position, so a Sketch-only month lands on **row 43**, not row 44 with a blank line above it. Because the template hard-codes Adobe's name and purpose into row 43, the tool writes `B`/`F` per row from the `VENDORS` table instead of trusting the template text, and clears `A`/`B`/`F`/`I` on any unused row. Clearing with `null` preserves the cell styles (verified against the template's style ids).
- **Week ending** defaults to the Sunday ending the Mon–Sun week that contains the *latest* of the two invoice dates. **Friday** = week ending − 2 days. Both are editable; the combined total lands on that Friday (`F27` + `F12`). This mirrors the pattern in the original July report exactly.
- **Combined total** = Adobe amount + Sketch amount, to 2 decimals, written to `F27`.
- **Receipts** render at 408px wide (≈4.25in, matching the template), height by aspect ratio. The first present invoice goes in the left column (`col: 0`), the second in the right (`col: 7`) — so a lone receipt sits on the left rather than stranded on the right.
- **Filename**: `Consultant_Expense_Report_-_Jeff_Heppert_-_<Month>_<Year>.xlsx`, month/year taken from the week-ending date.

## Dependencies (all CDN, loaded at runtime)

- pdf.js `4.0.379` (ESM build + worker) — text extraction + page-to-canvas rendering
- ExcelJS `4.4.0` (UMD global `ExcelJS`)
- Google Fonts: Sora, Inter, IBM Plex Mono

No internet = no libraries, so the tool needs a connection on load. Everything else is local.

## Design

Grounded in a ledger/receipt theme (the subject's own world): warm paper background, ink-navy text, mono figures for money and dates, a deep ledger-green for the action and totals, amber for fields that need attention. The signature element is the two invoice "slots" styled as receipt cards with a torn/perforated top edge. Single centered column, reduced-motion respected, keyboard-accessible.

## Possible future work

- **More vendors / variable line items.** Still hard-wired to the Adobe + Sketch pair, though the detail-row writer is already positional (`DETAIL_ROWS` + the `VENDORS` table), so generalizing is mostly a dynamic list of slots plus extending `DETAIL_ROWS` into the spare rows 45–48.
- **Split charges across different weeks.** Right now both charges lump onto one week's Friday. If invoices ever need to land on different Fridays, the grid logic would need per-invoice week assignment.
- **Print-friendly receipts.** Side-by-side receipts span ~8.5in; a print-optimized layout could scale them down or offer a stacked toggle.
- **Rebuilding the template.** If Red Oak changes the template, re-embed it: blank the month-specific cells (`A10`, `F12`, `F27`, `A43`, `A44`, `I43`, `I44`), rasterize any vector shapes to pictures, strip old receipt media, base64-encode, and replace `TEMPLATE_B64` in the HTML.

## Verifying changes without a browser

The generate path can be exercised headlessly, which is worth doing for anything touching cell writes or receipt anchoring. The approach that worked: extract the page's module `<script>` from `index.html`, strip the pdf.js `import`, run it via `new Function` with small stubs for `document`/`window`/`URL` and `globalThis.ExcelJS = require('exceljs')`, and capture the bytes by stubbing `Blob` to keep `parts[0]`. Append a line exposing the internals (`globalThis.__T = { state, generate, ... }`) so the test can drive them. That runs the real code rather than a copy. Then unzip the output and assert on the XML directly: `fullCalcOnLoad` present in `xl/workbook.xml`, receipt anchors with `colOff=0` in the Receipts drawing, the 40 formula cells unchanged against the template (compare maps — shared-formula followers have no `<f>` text, so don't filter them out), and rows 43/44 style ids unchanged.

## Resuming with Claude

If you hand this back to Claude later, the fastest path is: share `index.html` plus this file. The tool is self-contained — the template is embedded, so no other files are needed to regenerate or modify it. The trickiest areas to touch are the receipt anchoring (integer-column rule above) and the `fullCalcOnLoad` mutation; both are easy to break in ways that only show up when the file is opened in real Excel.
