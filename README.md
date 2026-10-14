# Saint Eyewear month-end report

Plain HTML + JavaScript. No server, no Python, no password, nothing hardcoded per month. Everything runs in the browser.

Follows the SOP in five steps on one page: 1 Inputs, 2 Tie-out gate, 3 Build, 4 Review, 5 Deliver.

1. Open the page.
2. Drop in the Xero Profit and Loss, Xero Balance Sheet and Daily Reporting Sheet (any order).
   The first time, also drop last month's Month-End report (.docx).
3. Press **Build report**. Download the Word report and tie-out workbook; copy the call points.

Each finished report is remembered by this browser and used as next month's template, so after the first run you only drop the three files.
If you use another browser or computer, drop last month's report again.

    index.html   the page
    src/         the tool source (readers, tie-out checks, charts, Word editing, read-back check)
    app.js       the same code bundled into one file - this is what the page loads (readers, tie-out checks, charts, Word editing, read-back check)
    lib/         ExcelJS and JSZip (bundled)

Deploy: push to GitHub, import in Vercel, no settings (any static host works).
No PDF is made. Word asks to update fields on opening - click Yes so the contents page numbers refresh.
New Xero account headings stop the run with a message; add them to `src/config.js`.

## Changing the code
`app.js` is built from `src/`. After editing anything in `src/`, run `npx esbuild src/app.js --bundle --format=iife --target=es2020 --outfile=app.js`. It works without a server (open index.html directly) and on any static host.
