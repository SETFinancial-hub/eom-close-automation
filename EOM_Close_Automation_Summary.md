# SET Financial Corporation -- EOM Close Automation Summary

**Prepared for:** Robert & Elliot Davis Accounting Team
**Date:** March 2026
**Purpose:** Walk through exactly what the automation does, where every number comes from, and how it maps to QBO -- so we can validate side-by-side for the February 2026 close.

---

## What This Document Covers

1. How the current manual process maps to the automated process (step by step)
2. Which spreadsheet columns feed which GL accounts, for every journal entry
3. How the class (portfolio/state) tagging works
4. December 2025 as a worked example with real numbers
5. Known items to clean up (miscategorizations, reclassifications)

---

## The Big Picture: Current vs. Automated

| Step | Current (Manual) | Automated |
|------|-----------------|-----------|
| Pull registers from Nortridge | Download 4 Excel files manually | Same -- still manual download, placed in `data/2026-02/` folder |
| Summarize Collection Register | Open work paper, enter formulas, verify Cash Received = sum of components | Parser reads Sheet2 automatically, sums every column, breaks out by portfolio |
| Summarize Loan Register | Open work paper, enter formulas, verify Note Amount = components | Parser reads Sheet2, sums all fields, breaks out by portfolio |
| Summarize Charge-Off Report | Open work paper, enter totals | Parser reads Sheet2, validates Net + Unearned = Total automatically |
| Summarize Unearned Register | Open work paper, copy prior month formulas | Parser reads Sheet2, computes all unearned balances by type and portfolio |
| Build JEs in QBO | Open memorized transactions, enter amounts by class, reference prior month | System generates all 9 JEs with correct accounts, amounts, and class tags |
| Pier/DPV interest | Manually read PDF statements, enter JE | Parser extracts interest and balance from PDFs |
| Metacorp bad debt sale | Read closing statement, enter JE | Manual entry into `metacorp.json`, then auto-generates JE |
| Reconcile & tie out | Compare work papers to QBO, spot check | Automated balance checks, component validations, class total cross-checks |
| Record in QBO | Type each JE into QBO by class | CSV output ready for import, or use QBO Import sheet to copy/paste |

**What still requires human judgment:**
- JE-9 Allowance (management estimate -- flagged for review every month)
- JE-1 Finance Income (using interest collected as proxy until we have 2 months of unearned register data to do month-over-month delta)
- JE-2 Insurance Earnings (uses aggregate QBO prior balances -- no class-level priors yet)
- Any new vendor categorization or expense reclassification

---

## Source Files: What Goes In

Every month, 7 source files go into the `data/YYYY-MM/` folder:

| File | Where It Comes From | What It Contains |
|------|-------------------|-----------------|
| Collection Register (Excel) | Nortridge client reports | Every payment received: principal, interest, fees, renewals, recoveries, refunds |
| Loan Register (Excel) | Nortridge client reports | Every new loan originated: note amount, finance charge, insurance premiums, cash disbursed |
| Charge Off Report (Excel) | Nortridge client reports | Every charged-off account: net loss, unearned interest reversed, total removed |
| Unearned Register (Excel) | Nortridge client reports | Every active loan's unearned balances: interest, insurance by type (life, A&H, IUI, property, VSI) |
| Pier Statement (PDF) | PAT LLC Series 47 monthly statement | Ending balance, accrued interest due |
| DPV Statement (PDF) | DPV LLC / BofA invoice | Principal balance, interest amount |
| metacorp.json | Manual entry from Metacorp closing statement PNG | Number of accounts, total balance, transfer amount |

---

## Journal Entry Detail: Source Column --> GL Account

### JE-1: Finance Income (Earned Interest)

**What it does:** Recognizes the interest income earned during the month. The "unearned" balance on the books decreases as borrowers make payments -- that decrease is the earned income.

**Current method:** Accountant looks at prior month unearned vs. current month unearned, records the difference.

**Automated method:** Uses InterestCollected from the Collection Register by portfolio as the earned interest proxy (same net effect). Future months will use the unearned register month-over-month delta once we have consecutive months of data.

**Source:** Collection Register, column `InterestCollected`, grouped by PortfolioID

| Portfolio | Class | Amount (Dec 2025) |
|-----------|-------|-------------------|
| 0 (legacy) | SET | $324.48 |
| 1 | SC | $75,593.18 |
| 5 | UT | $12,451.95 |
| 6 | MO | $42,602.68 |
| 11 | ID | $6,267.04 |
| 12 | MS | $10,482.86 |
| **TOTAL** | | **$147,722.19** |

**GL mapping (per class):**

| Line | Account | Account Name | Debit | Credit |
|------|---------|-------------|-------|--------|
| DR | 110010 | Unearned Pre-Computed Interest | earned amount | |
| CR | 400000 | Finance Income | | earned amount |

**How it balances:** Total debit to 110010 = Total credit to 400000 = $147,722.19

---

### JE-2: Insurance Premium Earnings

**What it does:** Recognizes insurance commission income as premiums are earned over the loan life. Compares prior month unearned insurance balances (from QBO) to current month (from Unearned Register).

**Class:** SET (aggregate) -- we don't have class-level prior insurance balances in QBO yet. Amounts are very small ($1.99 total unearned insurance in Dec 2025).

**Source:** Unearned Register insurance columns vs. QBO prior month balances

| Insurance Type | Unearned Account | Earned Account | Prior Balance | Current Balance | Earned |
|---------------|-----------------|---------------|---------------|-----------------|--------|
| Credit Life | 110050 | 410100 | $1.28 | (from register) | difference |
| Disability (A&H) | 110060 | 410200 | $0.71 | (from register) | difference |
| IUI | 110070 | 410500 | $0.00 | (from register) | difference |
| Property | 110080 | 410300 | $0.00 | (from register) | difference |
| VSI (Auto) | 110090 | 410400 | $0.00 | (from register) | difference |

**December 2025:** $0 earned (balances were de minimis and unchanged).

---

### JE-3: Loan Originations

**What it does:** Books all new loans originated during the month. The gross note amount is the receivable; it gets split into unearned interest (finance charge), insurance premiums, cash sent to borrower, and any renewed balance from refinanced loans.

**Source:** Loan Register Sheet2, grouped by PortfolioID

| Column in Loan Register | GL Account | Account Name | DR/CR |
|------------------------|-----------|-------------|-------|
| NoteAmount | 110001 | Loans Receivable Gross | DR |
| FinanceCharge | 110010 | Unearned Pre-Computed Interest | CR |
| OriginalCreditLifePremium | 110050 | Unearned Life Ins Commission | CR |
| OriginalAndHPremium | 110060 | Unearned A&H Ins Commissions | CR |
| CashToBorrower | 100140 | Veritex Bank - Funding DDA | CR |
| BalanceRenewed | 110001 | Loans Receivable Gross | CR |

**How it balances:** NoteAmount = FinanceCharge + Insurance + CashToBorrower + BalanceRenewed

**December 2025 by class:**

| Class | Note Amount (DR 110001) | Cash to Borrower (CR 100140) |
|-------|------------------------|------------------------------|
| SC | $88,450.00 | $88,450.00 |
| UT | $23,300.00 | $23,300.00 |
| MO | $69,250.00 | $69,250.00 |
| ID | $7,550.00 | $7,550.00 |
| MS | $24,700.00 | $24,700.00 |
| **TOTAL** | **$213,250.00** | **$213,250.00** |

Note: December 2025 originations were all NLS simple-interest loans (NoteAmount = CashToBorrower, no finance charge or insurance). Pre-computed loans would also have CR lines to 110010, 110050, 110060.

---

### JE-4: Collections / Payments Received

**What it does:** Records all borrower payments received. Cash goes to the bank; the loan balance gets reduced; fees and recoveries go to their income accounts.

**Source:** Collection Register Sheet2, grouped by PortfolioID

| Column in Collection Register | GL Account | Account Name | DR/CR |
|------------------------------|-----------|-------------|-------|
| CashReceived | 100120 | Veritex Bank - Payments DDA | DR |
| BalanceRenewed | 110001 | Loans Receivable Gross | DR |
| *(computed residual)* | 110001 | Loans Receivable Gross | CR |
| LateFees | 440000 | Delinquent/NSF Fees | CR |
| NSFFees | 440000 | Delinquent/NSF Fees | CR |
| AmountToRefund | 460000 | Refunds | DR |
| InsuranceRebate | 410600 | Earned Insurance Prem Rebates | CR |
| Recovery | 420100 | Customer Recoveries | CR |

The **receivable credit** (CR to 110001) is the plug that makes the entry balance:
`Receivable CR = CashReceived + BalanceRenewed - LateFees - NSFFees - InsuranceRebate - Recovery + Refunds`

This receivable reduction includes both principal and interest portions of payments (the interest earned recognition is handled separately in JE-1).

**December 2025 by class:**

| Class | Cash In (DR Bank) | Receivable CR | Late Fees | NSF Fees | Recoveries |
|-------|-------------------|---------------|-----------|----------|------------|
| SET | $397.18 | $397.18 | -- | -- | -- |
| SC | $153,056.91 | $151,149.39 | $702.00 | $666.56 | $538.96 |
| UT | $23,588.09 | $23,190.25 | $320.00 | $40.00 | $37.84 |
| MO | $76,148.87 | $74,820.06 | $360.00 | $273.50 | $695.31 |
| AL | $5,626.89 | $5,121.79 | $198.00 | -- | $307.10 |
| ID | $12,633.59 | $12,478.59 | $87.50 | $67.50 | -- |
| MS | $18,107.16 | $18,031.70 | $75.46 | -- | -- |
| **TOTAL** | **$289,558.69** | **$285,188.96** | **$1,742.96** | **$1,047.56** | **$1,579.21** |

---

### JE-5: Charge-Offs

**What it does:** When a loan is written off, the full balance (TotalChargeOffAmt) is removed from the receivable via the contra account. The net loss (ChargeOffAmount) hits P&L expense. Any remaining unearned interest (PCInterestRebate) is reversed since it was never earned.

**Source:** Charge Off Report Sheet2, grouped by PortfolioID

| Column in Charge Off Report | GL Account | Account Name | DR/CR |
|-----------------------------|-----------|-------------|-------|
| ChargeOffAmount | 610500 | Bad Debt Writeoffs | DR |
| PCInterestRebate | 110010 | Unearned Pre-Computed Interest | DR |
| TotalChargeOffAmt | 110200 | Accumulated Charge Offs | CR |

**How it balances:** ChargeOffAmount + PCInterestRebate = TotalChargeOffAmt

**December 2025 by class:**

| Class | Net Writeoff (DR 610500) | Unearned Reversed (DR 110010) | Total Removed (CR 110200) |
|-------|--------------------------|-------------------------------|---------------------------|
| SC | $35,691.51 | $6,489.83 | $42,181.34 |
| UT | $5,309.54 | $531.53 | $5,841.07 |
| MO | $16,952.79 | $862.13 | $17,814.92 |
| AL | $5,793.51 | $0.00 | $5,793.51 |
| MS | $2,700.00 | $0.00 | $2,700.00 |
| **TOTAL** | **$66,447.35** | **$7,883.49** | **$74,330.84** |

---

### JE-6: Bad Debt Sale (Metacorp)

**What it does:** Records the proceeds from selling charged-off accounts to Metacorp.

**Class:** SET (company-level transaction, not portfolio-specific)

**Source:** Metacorp closing statement (entered manually into metacorp.json)

| Line | Account | Account Name | Amount (Dec 2025) |
|------|---------|-------------|-------------------|
| DR | 100150 | Veritex Bank - Operating DDA | $5,208.89 |
| CR | 420200 | Sale of Bad Debt | $5,208.89 |

December: 57 accounts sold, $5,208.89 proceeds.

---

### JE-7: Pier Facility Interest

**What it does:** Accrues monthly interest on the Pier (PAT LLC Series 47) credit facility.

**Class:** SET (facility-level)

**Source:** Pier monthly statement PDF

| Line | Account | Account Name | Amount (Dec 2025) |
|------|---------|-------------|-------------------|
| DR | 520100 | Interest Expense | $23,653.21 |
| CR | 290060 | Accrued Expenses | $23,653.21 |

Pier ending balance: $1,591,414.81

---

### JE-8: DPV LLC Interest

**What it does:** Records DPV LLC (BofA line of credit) interest, which is auto-debited from the bank account.

**Class:** SET (facility-level)

**Source:** DPV/BofA monthly invoice PDF

| Line | Account | Account Name | Amount (Dec 2025) |
|------|---------|-------------|-------------------|
| DR | 31200 | DPV Interest | $5,033.61 |
| CR | 100120 | Veritex Bank - Payments DDA | $5,033.61 |

DPV principal balance: $1,161,974.17

---

### JE-9: Allowance for Credit Losses

**What it does:** Adjusts the allowance reserve to the target percentage of the portfolio. This is a management judgment call -- always flagged for review.

**Class:** SET (aggregate management estimate)

**Source:** Calculated from portfolio balance and target percentage

| Input | Value (Dec 2025) |
|-------|-------------------|
| Portfolio balance (from Unearned Register) | $1,930,553.07 |
| Target allowance % | 18% |
| Target allowance $ | $347,499.55 |
| Current allowance (from QBO) | $328,194.02 |
| **Adjustment needed** | **$19,305.53** |

| Line | Account | Account Name | Amount |
|------|---------|-------------|--------|
| DR | 610502 | Allowance Adjustment | $19,305.53 |
| CR | 110002 | Allowance for Credit Losses | $19,305.53 |

---

## Class (Portfolio) Mapping

Every JE line is tagged with a QBO class that corresponds to the state where the portfolio operates:

| Portfolio ID in Nortridge | State | QBO Class |
|--------------------------|-------|-----------|
| 0 | Legacy/unassigned | SET |
| 1 | South Carolina | SC |
| 5 | Utah | UT |
| 6 | Missouri | MO |
| 8 | Alabama | AL |
| 11 | Idaho | ID |
| 12 | Mississippi | MS |

Facility-level entries (Pier, DPV, Metacorp, Allowance, Insurance) all use class **SET**.

**Which JEs are broken out by class vs. aggregate:**

| JE | By Class? | Why |
|----|-----------|-----|
| JE-1 Finance Income | Yes -- SC, UT, MO, AL, ID, MS, SET | Interest collected by portfolio from Collection Register |
| JE-2 Insurance Earnings | No -- SET only | No class-level prior insurance balances in QBO yet (de minimis) |
| JE-3 Originations | Yes | Loan Register has PortfolioID for each origination |
| JE-4 Collections | Yes | Collection Register has PortfolioID for each payment |
| JE-5 Charge-Offs | Yes | Charge Off Report has PortfolioID for each account |
| JE-6 Bad Debt Sale | No -- SET | Company-level Metacorp transaction |
| JE-7 Pier Interest | No -- SET | Facility-level |
| JE-8 DPV Interest | No -- SET | Facility-level |
| JE-9 Allowance | No -- SET | Aggregate management estimate |

---

## How Everything Ties Out

The system runs these automated checks every month:

1. **Every JE balances:** Total debits = Total credits (within $0.01)
2. **Charge-off components:** Net ($66,447.35) + Unearned ($7,883.49) = Total ($74,330.84)
3. **Collections components:** Cash received reconciles to principal + interest + fees + recoveries - refunds
4. **Class totals match aggregates:** The sum of all SC + UT + MO + AL + ID + MS + SET lines equals the aggregate register total
5. **No duplicate loan numbers** in the Loan Register
6. **All portfolio IDs are valid** (mapped in the system)
7. **Pier balance** ties to QBO (Statement $1,591,414.81 = QBO $1,591,414.81)
8. **Unearned interest** register total vs. QBO balance (may differ by timing)

---

## For February 2026 Testing

### What We Need

1. **Download the 4 Nortridge registers** for February 2026 and place in `data/2026-02/`:
   - Collection Register
   - Loan Register
   - Charge Off Report
   - Unearned Register

2. **Get the Pier and DPV statements** for February 2026 and place in the same folder

3. **Enter the Metacorp closing statement** values into `data/2026-02/metacorp.json`

4. **Add February prior balances** to the month config (January ending = February beginning)

5. **Run:** `PYTHONIOENCODING=utf-8 python src/run_close.py --month 2026-02`

### How to Validate

Compare the automation output side-by-side with the manual JEs:

| Check | Where to Look (Automated) | Where to Look (Manual) |
|-------|--------------------------|----------------------|
| JE totals match | Excel "Journal Entries" sheet, TOTALS row per JE | Manual work paper totals |
| Class breakdown matches | Each JE line has a Class column | QBO JE entry by class |
| Accounts are correct | Account # and Account Name columns | QBO chart of accounts |
| Everything balances | "Balanced" indicator on each JE | Manual footing |
| Component math works | Reconciliation sheet | Manual cross-check |

---

## Known Items to Clean Up (Beginning of Year Reclassifications)

These are categorization issues in QBO that the automation exposes but does not fix automatically. Beginning of the year is the ideal time to correct these because it keeps the annual financials clean from the start.

### 1. Expense vs. Income Misclassifications

Items currently coded as expenses that should be income (or vice versa). Common patterns to look for:
- **Fee income booked as expense offsets** -- e.g., late fees or NSF fees netted against an expense line instead of credited to 440000 (Delinquent/NSF Fees)
- **Recovery income in wrong category** -- recoveries on charged-off accounts should go to 420100 (Customer Recoveries), not netted against 610500 (Bad Debt Writeoffs)
- **Insurance rebates** -- should be 410600 (Earned Insurance Prem Rebates), not mixed into collections

### 2. Vendor Misclassifications

Vendors assigned to the wrong expense category:
- **Tech vendors in Marketing** -- software subscriptions, SaaS tools, data services that are coded to marketing but should be in 606300 (Technology Services) or similar
- **Marketing vendors in Tech** -- ad spend or lead generation services that landed in technology
- **Credit bureau fees** -- Republic Bank and similar should consistently go to 610800 (Credit Reporting Fees), per the Elliot Davis instructions

### 3. Bank Account Coding

Per the Elliot Davis instructions, specific bank feeds have known issues:
- **Veritex 6121 (Payments DDA):** Bank feed sometimes auto-categorizes deposits as "Unapplied Collections" -- should be corrected to "Unapplied Loan Disbursements" (100510, Class SET)
- **Veritex 6950 (Funding DDA):** "Electronic System Inc" should always go to 606300 (Technology Services), "Republic Bank" to 610800 (Credit Reporting Fees)
- **Large non-transfer deposits** on Payments DDA: Should go to 31000 Capital Contribution (Derek approval only)

### 4. Class Assignment Gaps

- Some older transactions may have no class assigned -- should be SET or the appropriate state
- Check that all register-based JEs now have classes assigned (the automation handles this going forward)

### 5. Suggested Review Process

1. Pull a QBO Profit & Loss by Class for January and February 2026
2. Look for any accounts where amounts appear in unexpected classes
3. Pull a vendor list and spot-check that each vendor's default account makes sense
4. Reclassify with journal entries at the beginning of 2026 so the full year is clean
5. Document any reclassifications so the automation can be updated if account mappings need to change

---

## Quick Reference: All GL Accounts Used by the Automation

### Balance Sheet Accounts

| Account | Name | Used In |
|---------|------|---------|
| 100120 | Veritex Bank - Payments DDA | JE-4 (cash in), JE-8 (DPV auto-debit) |
| 100140 | Veritex Bank - Funding DDA | JE-3 (cash out to borrowers) |
| 100150 | Veritex Bank - Operating DDA | JE-6 (Metacorp proceeds) |
| 110001 | Loans Receivable Gross | JE-3 (DR new loans), JE-4 (CR payments) |
| 110002 | Allowance for Credit Losses | JE-9 (allowance adjustment) |
| 110010 | Unearned Pre-Computed Interest | JE-1 (DR earned), JE-3 (CR new loans), JE-5 (DR reversed on CO) |
| 110050 | Unearned Life Ins Commission | JE-2 (DR earned), JE-3 (CR new loans) |
| 110060 | Unearned A&H Ins Commissions | JE-2 (DR earned), JE-3 (CR new loans) |
| 110070-90 | Other Unearned Insurance | JE-2 (DR earned) |
| 110200 | Accumulated Charge Offs | JE-5 (CR charge-offs) |
| 290060 | Accrued Expenses | JE-7 (CR Pier interest accrual) |
| 31200 | DPV Interest | JE-8 (DR DPV interest) |

### Income Accounts

| Account | Name | Used In |
|---------|------|---------|
| 400000 | Finance Income | JE-1 (CR earned interest) |
| 410100-500 | Earned Insurance Commissions | JE-2 (CR insurance earned) |
| 410600 | Earned Insurance Prem Rebates | JE-4 (CR rebates collected) |
| 420100 | Customer Recoveries | JE-4 (CR recoveries on charged-off accounts) |
| 420200 | Sale of Bad Debt | JE-6 (CR Metacorp sale proceeds) |
| 440000 | Delinquent/NSF Fees | JE-4 (CR late fees + NSF fees) |
| 460000 | Refunds | JE-4 (DR customer refunds) |

### Expense Accounts

| Account | Name | Used In |
|---------|------|---------|
| 520100 | Interest Expense | JE-7 (DR Pier interest) |
| 610500 | Bad Debt Writeoffs | JE-5 (DR net charge-off loss) |
| 610502 | Allowance Adjustment | JE-9 (DR/CR allowance provision or release) |

---

*This document can be updated as we validate the February close and identify additional items to address.*
