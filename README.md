# Fieldbook

Simple books for a one-person practice: write down money in and money out, and get a
profit & loss statement. It's meant to be better than a spreadsheet without turning into
accounting software.

- **Ledger**: add an entry in one line. Familiar clients and descriptions fill in their
  category. Search and filter by year, kind or category.
- **Reports**: profit & loss for a year, year to date, a quarter, a month or any dates you
  choose. It prints cleanly and exports to Excel. You also get month-by-month totals and
  income by client.
- **Settings** (the gear at the top right): backups, categories (which become the P&L lines) and your practice name.
- **Undo** after every change, sample entries to try it out, light and dark themes, and it
  works offline.

## Where the data lives

Everything is stored in the browser's `localStorage` under the key `fieldbook`. There are
no accounts and no server. Clearing browser data wipes it, and Safari may clear site data
for sites you haven't opened in a while. So backups are the safety net:

- **Download backup** saves an ordinary `.xlsx` with these sheets: *Entries*,
  *P&L by year*, *Categories* and *About*. The brass tag in the header counts changes
  since the last backup and turns red after a week.
- **Restore or import…** reads a Fieldbook backup back in, replacing what's there. It
  also reads any other spreadsheet (.xlsx, .xls, .csv, .ods) with a *Date* column and
  either an *Amount* column or separate *Income* and *Expense* columns, and adds those
  rows. Before changing anything it shows a preview and skips rows that are already in
  the books.
- **Bank exports** (a CSV with *Posted Date*, *Description*, *Amount*, *Balance*,
  *Transaction Type* and *Memo*, plus *From/To Account Name* for transfers) import the
  same way. *Transaction Type* (credit/debit) decides money in or out, the memo is added to
  the description, and transfers between your own accounts (owner draws, moves to savings)
  are left out because they aren't income or expenses. A bank row counts as already in the
  books when an entry has the same amount and direction within 4 days, so overlapping
  downloads and entries you typed in yourself aren't doubled up.

Amounts are stored as whole cents and dates as `YYYY-MM-DD` strings. Keep changes to the
stored shape backward compatible.

## How it's built

A static site with no build step, served from GitHub Pages:

| File | What it is |
|---|---|
| `index.html` | The screens |
| `styles.css` | Styles (light, dark and print) |
| `app.js` | Everything else |
| `vendor/xlsx.full.min.js` | [SheetJS](https://sheetjs.com) 0.20.3, for reading and writing Excel |
| `sw.js` | Service worker for offline use (bump `CACHE` when the file list changes) |

To run it locally, serve the folder (for example `python3 -m http.server 8766`) and open
http://localhost:8766.
