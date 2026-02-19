"""
Chart of Accounts mapping for SET Financial Corporation QBO.
All account codes and their descriptions for journal entry generation.
"""

# Revenue Accounts
FINANCE_INCOME = "400000"
EARNED_LIFE_INS = "410100"
EARNED_AH_INS = "410200"
EARNED_PROP_INS = "410300"
EARNED_AUTO_INS = "410400"
EARNED_IUI_INS = "410500"
EARNED_INS_REBATES = "410600"
CUSTOMER_RECOVERIES = "420100"
SALE_OF_BAD_DEBT = "420200"
NONFILING_FEES = "430000"
DELINQUENT_NSF_FEES = "440000"
REFUNDS = "460000"

# COGS / Finance Costs
CONTRACT_SERVICES = "510000"
INTEREST_EXPENSE = "520100"
COLLECTION_EXPENSES = "540400"
CREDIT_REPORTING_FEES = "540500"

# Operating Expenses
BAD_DEBT_WRITEOFFS = "610500"
ALLOWANCE_ADJUSTMENT = "610502"
DEPRECIATION = "610600"
AMORTIZATION_EXPENSE = "610700"
CC_PROCESSING_FEES = "610800"

# Balance Sheet - Loans Receivable
LOANS_RECEIVABLE_GROSS = "110001"
ALLOWANCE_CREDIT_LOSSES = "110002"
UNEARNED_PRECOMPUTED_INTEREST = "110010"
UNEARNED_LIFE_INS = "110050"
UNEARNED_AH_INS = "110060"
UNEARNED_IUI_INS = "110070"
UNEARNED_PROP_INS = "110080"
UNEARNED_AUTO_INS = "110090"
ACCUMULATED_CHARGE_OFFS = "110200"

# Balance Sheet - Liabilities
PIER_LOC = "291300"
ACCRUED_EXPENSES = "290060"

# Balance Sheet - Equity / Shareholder
DPV_LOC = "310000"
DPV_PRINCIPAL = "31100"
DPV_INTEREST = "31200"
EDWARDS_LOC = "320000"
EDWARDS_PRINCIPAL = "321000"
EDWARDS_INTEREST = "322000"

# Insurance unearned-to-earned mapping
INSURANCE_MAPPING = {
    "CreditLife": {"unearned": UNEARNED_LIFE_INS, "earned": EARNED_LIFE_INS},
    "Disability": {"unearned": UNEARNED_AH_INS, "earned": EARNED_AH_INS},
    "IUI": {"unearned": UNEARNED_IUI_INS, "earned": EARNED_IUI_INS},
    "Property": {"unearned": UNEARNED_PROP_INS, "earned": EARNED_PROP_INS},
    "VSI": {"unearned": UNEARNED_AUTO_INS, "earned": EARNED_AUTO_INS},
}

# Account name lookup
ACCOUNT_NAMES = {
    "400000": "Finance Income",
    "410100": "Earned Life Ins Commission",
    "410200": "Earned A&H Ins Commissions",
    "410300": "Earned Prop Ins Commissions",
    "410400": "Earned Auto Ins Commissions",
    "410500": "Earned IUI Commissions",
    "410600": "Earned Insurance Prem Rebates",
    "420100": "Customer Recoveries",
    "420200": "Sale of Bad Debt",
    "430000": "Nonfiling Fees/Personal Prop",
    "440000": "Delinquent/NSF Fees",
    "460000": "Refunds",
    "510000": "Contract Services",
    "520100": "Interest Expense",
    "540400": "Collection Expenses",
    "540500": "Credit Reporting Fees",
    "610500": "Bad Debt Writeoffs",
    "610502": "Allowance Adjustment",
    "610600": "Depreciation",
    "610700": "Amortization Expense",
    "610800": "Credit Card Processing Fees",
    "110001": "Loans Receivable Gross",
    "110002": "Allowance for Credit Losses",
    "110010": "Unearned Pre-Computed Interest",
    "110050": "Unearned Life Ins Commission",
    "110060": "Unearned A&H Ins Commissions",
    "110070": "Unearned IUI Commissions",
    "110080": "Unearned Prop Ins Commissions",
    "110090": "Unearned Auto Ins Commissions",
    "110200": "Accumulated Charge Offs",
    "291300": "Pier Active Transactions LOC",
    "290060": "Accrued Expenses",
    "310000": "DPV LLC Line of Credit",
    "31100": "DPV Principal",
    "31200": "DPV Interest",
}
