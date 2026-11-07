# Harborview Commercial Properties, LLC
## Simulated Commercial Property Accounting & Financial Reporting Workbook — FY2025

A complete accrual-basis accounting cycle for a small commercial real-estate owner, built as a
formula-linked Excel model. Nothing is typed in twice: the trial balance sums the ledger, the
ledger sums the journal, and the three financial statements read the trial balance.

**File:** `Commercial_Property_Accounting_Workbook_FY2025.xlsx`
**Entity:** Harborview Commercial Properties, LLC — two properties, fiscal year 1 Jan – 31 Dec 2025
**Basis:** Accrual, US GAAP · USD · pass-through LLC (no entity-level income tax)

---

## What is in the workbook (11 sheets)

| # | Sheet | What it demonstrates |
|---|-------|----------------------|
| 1 | Company Profile & COA | Entity profile, property portfolio, rent roll, accounting policies, 32-account chart of accounts, assumptions |
| 2 | Transaction Register | 50 transactions in date order with cash effect, counterparty and the journal each produced |
| 3 | Journal Entries | 209 double-entry lines; every entry carries a debit = credit check |
| 4 | General Ledger | Opening balance, debits, credits, closing balance per account, plus full detail with a running balance |
| 5 | Trial Balance | Opening + operating + adjusting = closing, proved to zero |
| 6 | Adjusting Entries | 8 year-end accruals, prepaid amortisation, depreciation and bank items, with the mortgage schedule |
| 7 | Bank Reconciliation | Bank statement vs cash book, deposits in transit and outstanding cheques, both sides proved |
| 8 | Profit & Loss Statement | Property revenue, property operating expenses, NOI, then interest, depreciation and tax |
| 9 | Balance Sheet | 2025 with a 31-Dec-2024 comparative column from opening balances |
| 10 | Cash Flow Statement | Operating (indirect), investing and financing, reconciled to the movement in cash |
| 11 | Management Financial Package | KPI panel, property-level P&L (Rivergate / Maple Street / Corporate / Total), debt position, commentary |

---

## The accounting story

**Two properties.** Rivergate Office Center (owned since 2023, 3 office tenants, 24,000/month) and
Maple Street Retail Shops (acquired 1 March 2025 for 2,140,000 funded by a 1,400,000 mortgage and a
700,000 member contribution, 3 retail tenants, 13,000/month).

**Accrual accounting, as required:**

- **December rent earned but not received.** December 2025 invoices of 42,100 are unpaid at
  31 December. The revenue is 2025 revenue and the balance sits in tenant receivables — it is never
  pushed into January's income. Collections run one month behind billing, so the receivable is one
  month of billings at every year end.
- **Prepaid insurance.** 54,000 was paid on 1 October for 12 months of cover. Only 13,500 expired in
  2025 (plus 9,000 carried in from the prior policy), so 22,500 is expensed and 40,500 remains an
  asset (AJ-03).
- **Accrued expenses.** Maple Street property tax for Mar–Dec (12,000, AJ-04), the Q4 operating
  settlement invoiced in January (29,160, AJ-05) and December mortgage interest (AJ-06) are all 2025
  expenses with 2026 payment dates.
- **Owner distributions.** 60,000 was paid to the members in two instalments. It is debited to a
  contra-equity account (3020), never to an expense, so it does not touch net income or NOI. The
  equity roll-forward on the balance sheet proves this.
- **Security deposits.** 36,000 collected and 6,000 refunded; they are a liability (2030), not income.
- **Bank items.** Service charges of 185 and an NSF cheque of 3,600 appear only on the bank statement;
  they are booked as AJ-07/AJ-08 and then reconcile.

---

## Headline results (all recomputed from the workbook)

| | |
|---|---|
| Property revenue | 479,250 |
| Property operating expenses | 209,900 |
| **Net operating income** | **269,350** (56.2% margin) |
| Depreciation | 100,952 |
| Mortgage interest | 138,749 |
| **Net income** | **30,654** |
| Total assets = liabilities + equity | 5,177,374 |
| Cash at 31 December 2025 | 89,126 |
| Distributions (equity) | 60,000 |
| Bank reconciliation | 77,026 statement + 42,100 in transit − 30,000 outstanding = 89,126 |

Cash flow: operating **+126,612**, investing **−2,140,000**, financing **+2,020,907**, net
**+7,520**, taking cash from 81,606 to 89,126 — proved back to the balance sheet.

---

## Controls built into the file

There are **249 check cells**; all of them evaluate to OK / BALANCED / RECONCILED:

- every journal entry and adjusting entry — debits equal credits
- general ledger totals — opening + movements = closing, proved to zero
- trial balance — total debits equal total credits
- balance sheet — total assets less (liabilities + equity) = 0, for both years
- equity roll-forward — agrees to the balance sheet
- cash flow — closing cash agrees to the balance sheet
- bank reconciliation — adjusted bank balance less adjusted book balance = 0
- management package — property totals agree to the consolidated income statement

The workbook is set to recalculate on open. All amounts are formulas over the journal, so if you
change a transaction the ledger, trial balance, statements and package all follow.

---

## Files

- `Commercial_Property_Accounting_Workbook_FY2025.xlsx` — the deliverable
- `build_workbook.py` — the generator script, kept so the model can be reproduced or re-timed
  (charts of accounts, opening balances, 50 transactions, 8 adjusting entries and both mortgage
  schedules are defined there and validated before the workbook is written)

---

## Suggested next steps

1. Take the workbook into Tally Prime or Zoho Books and re-key a sample month to compare the
   accrual postings — the journal here is the specification.
2. Add a 2026 column and roll the accruals and prepayments forward.
3. Build a simple budget-vs-actual variance on the package sheet.
