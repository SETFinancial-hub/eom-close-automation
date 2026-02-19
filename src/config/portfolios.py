"""
Portfolio and Branch mappings for SET Financial Corporation.
"""

PORTFOLIO_STATE_MAP = {
    0: "UNKNOWN",  # Unassigned/legacy
    1: "SC",    # South Carolina
    4: "BH",    # No longer active
    5: "UT",    # Utah
    6: "MO",    # Missouri
    7: "TN",    # Tennessee
    8: "AL",    # Alabama
    9: "MS",    # Mississippi (legacy EZLoan)
    10: "ID",   # Idaho
    11: "ID",   # Idaho (NLS alternate)
    12: "MS",   # Mississippi (NLS)
    13: "TN",   # Tennessee (NLS alternate)
}

BRANCH_MAP = {
    1: "Primary (SC-based)",
    2: "Secondary",
    3: "Legacy/EZLoan",
}

LOAN_TYPES = {
    "PC": "Pre-Computed (EZLoan legacy)",
    "NLS": "New Loan System (Nortridge)",
}

VALID_PORTFOLIO_IDS = set(PORTFOLIO_STATE_MAP.keys())
VALID_BRANCH_IDS = set(BRANCH_MAP.keys())
