"""
CSV report generator for QBO journal entry import.
"""
from pathlib import Path


def generate_csv_output(journal_entries, output_dir, period):
    """Generate a flat CSV for QBO journal entry import."""
    # Build dynamic filename from period
    period_parts = period.split()
    if len(period_parts) == 2:
        month_name, year = period_parts
        csv_filename = f"qbo_journal_entries_{month_name.lower()}{year}.csv"
    else:
        csv_filename = f"qbo_journal_entries_{period.replace(' ', '_').lower()}.csv"

    csv_path = Path(output_dir) / csv_filename
    with open(csv_path, 'w') as f:
        f.write("JE_Number,Date,Account_Number,Account_Name,Class,Debit,Credit,Memo\n")
        for je in journal_entries:
            for line in je.lines:
                debit_str = f"{line.debit}" if line.debit > 0 else ""
                credit_str = f"{line.credit}" if line.credit > 0 else ""
                memo = line.memo.replace(",", ";").replace('"', "'")
                f.write(f'{je.je_number},{je.je_date},{line.account_code},"{line.account_name}",{line.class_name},{debit_str},{credit_str},"{memo}"\n')
    print(f"  ✓ QBO import CSV saved: {csv_path}")
    return csv_path
