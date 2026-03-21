"""
create_sample_data.py
---------------------
Utility script to generate a sample Excel file (recruiters.xlsx)
with the expected column format so you can see the structure.

Run this once to create the sample, then replace with your real data.

Usage:
    python create_sample_data.py
"""

from openpyxl import Workbook


def create_sample_excel(filename: str = "recruiters.xlsx") -> None:
    """Create a sample Excel file with recruiter data columns."""

    wb = Workbook()
    ws = wb.active
    ws.title = "Recruiters"

    # --- Header row (must match exactly) ---
    headers = ["Name", "Email", "Company", "Role", "Status"]
    ws.append(headers)

    # --- Sample data rows ---
    sample_data = [
        ["Jane Smith", "jane.smith@example.com", "Google", "Software Engineer", "Not Contacted"],
        ["John Doe", "john.doe@example.com", "Meta", "Backend Developer", "Not Contacted"],
        ["Alice Johnson", "alice.j@example.com", "Amazon", "Data Engineer", "Not Contacted"],
    ]

    for row in sample_data:
        ws.append(row)

    wb.save(filename)
    print(f"✅ Sample Excel file created: {filename}")


if __name__ == "__main__":
    create_sample_excel()
