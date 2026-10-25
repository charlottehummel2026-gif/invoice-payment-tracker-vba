# A 90-second project walkthrough

Use the interactive browser demo for a quick review. Use the Excel sequence below to demonstrate the actual VBA. All examples use fictional data.

## Browser walkthrough

1. Open the portfolio and select **Try the workflow**.
2. Select **Start guided demo**. This resets the sample and fixes its reporting date at October 25, 2026.
3. Select **Add sample invoice**. A $1,200 invoice for Alder Office Co. is added with a unique ID. It is already overdue, so both outstanding and past due increase by $1,200.
4. Select **Record $400 payment**. The receipt references that invoice and reduces its remaining balance to $800.
5. Select **Show overdue invoices**. The remaining $800 is visible in the filtered register.
6. Select **Reset demo** to return to the original sample. Try the manual entry forms, search, payment register, and reporting date.

## Native Excel demonstration

Start from a fresh copy of the download. The original sample has 25 invoices and 15 payments. Keep the reporting date at October 25, 2026.

1. Open the workbook. The branded cover is first; **Open tracker** opens Overview.
2. Select **New invoice** and submit the empty form to show that missing input is rejected.
3. Enter: customer `Alder Office Co.`, description `Office support demo`, issued `October 1, 2026`, due `October 15, 2026`, amount `1200`, owner `Alex Morgan`. Select **Save invoice**.
4. Confirm invoice `INV-2026-026` appears, calculated columns populate, and outstanding increases from $39,450 to $40,650.
5. Select **Record payment**. Enter invoice `INV-2026-026`, received `October 20, 2026`, method `ACH`, amount `400`, reference `DEMO-RCPT-001`, recorded by `Alex Morgan`. Select **Save payment**.
6. Confirm `PAY-016` appears and invoice `INV-2026-026` has $400 received and $800 outstanding. Total outstanding is $40,250.
7. Return to Overview and select **Show overdue**. Explain that the invoice remains overdue because a partial payment does not settle its full balance.

## Suggested narration

"This project brings invoice entry, receipts, and receivables reporting into one Excel workbook. VBA checks the form before saving, assigns IDs, expands the tables, and restores calculated columns. Payments are separate records linked to invoice IDs, which supports partial payments and overpayments. The dashboard reports balances and aging as of a chosen date. The browser demo provides a quick way to explore the workflow; the download contains the actual VBA."

## Recording a video later

Record the native Excel sequence above after completing the checks in `TESTING.md`. Keep the workbook title and actions visible, use fictional entries, and pause briefly on each changed balance. A browser recording should be labeled **Browser workflow preview** rather than presented as Excel executing VBA.
