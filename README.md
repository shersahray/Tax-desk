# Tax Desk

An internal T1 tax-return workflow tracker for preparation and review.

## First version

- Create and search T1 return records
- Assign a preparer and reviewer
- Move a return through preparation, review, approval, and filing
- Record review notes
- Import an existing Excel or CSV list

This initial version stores sample and imported data in the browser. Before office rollout, connect it to the shared office database so every computer sees the same list.

## Run it

Open `index.html` in a modern browser. The Excel import uses the SheetJS CDN, so the browser needs internet access when importing `.xlsx` files.
