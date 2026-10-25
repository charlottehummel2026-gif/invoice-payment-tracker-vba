# Validation and test coverage

## Browser financial rules

`npm test` runs 17 regression tests using Node's built-in test runner. They cover:

- Reconciliation to the published workbook's totals and aging balance.
- Unique IDs, new invoices, partial/full payments, and overpayments.
- Future invoices/receipts and due-date boundaries.
- Every aging-bucket boundary.
- Invalid amounts, invalid dates, missing fields, and unmatched invoice IDs.
- Integer-cent arithmetic and reset isolation.

## Embedded VBA

`tests/workbook_smoke.py` reads the actual download, checks its VBA source against the exported modules, and runs the cover event, navigation, rejected empty form, invoice creation, partial/overpayment handling, and formula repair in pyOpenVBA's offline interpreter. Tests are in memory and do not save over the download.

**This is not native Excel execution.** Successful interpretation does not prove every Excel version's rendering, macro compilation, or button behavior. Earlier workbook builds also passed package/formula checks; those results are supplementary to the reproducible tests here.

## Portfolio release checks

The October 25, 2026 package passed **17 browser-rule tests**, **15 offline VBA smoke checks**, and **25 browser interaction/layout checks**. Browser checks covered the guided workflow, manual invoice/payment forms, rejected dates, text escaping, search and empty states, reporting-date changes, reset, the workbook download, and direct opening of `index.html` without a server. Layouts were inspected at 1440, 768, and 390 pixels wide; mobile forms remained within the viewport and Escape closed the dialog. No JavaScript runtime errors were reported.

The workbook previews are rendered from the downloadable workbook. The accompanying site uses the same fictional input records. The working workbook was not modified while creating this package.

## Native Excel checklist - still to verify

Use a fresh copy of the download in desktop Excel and record the Excel version and operating system when completing these checks:

- Open without a repair prompt; cover appears and **Open tracker** works.
- Compile the VBA project in the Visual Basic Editor.
- Run the sequence in `DEMO.md`, checking the resulting IDs and balances.
- Verify blank forms, due-before-issued dates, invalid amounts, and payments before invoice dates are rejected without adding records.
- Test a full payment and an overpayment; confirm outstanding does not become negative.
- Use overdue filtering, clear filtering, and formula refresh.
- Save while viewing another worksheet, close, and reopen with macros enabled; confirm Cover is activated.

## Verification limits

This demonstration has not been certified for production or concurrent multi-user use. The sample data is fictional. No measured business impact or time-saving claim is included.
