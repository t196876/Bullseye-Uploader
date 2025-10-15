# cleaner.py

import re

import openpyxl
FILE_PATH = "metadata\13190.xlsx"
import openpyxl

from datetime import datetime

def clean_excel_dates(file_path):
    """
    Cleans Excel file by:
    1. Removing extra spaces around text-based dates.
    2. Converting datetime-type cells to 'MM/DD/YY' format.
    Logs all changes.
    """
    try:
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active
        cleaned_count = 0

        for row in ws.iter_rows():
            for cell in row:
                val = cell.value

                # Handle datetime-type cells
                if isinstance(val, datetime):
                    formatted = val.strftime("%m/%d/%y")
                    if cell.value != formatted:
                        print(f"Normalized date: {val} → {formatted}")
                        cell.value = formatted
                        cleaned_count += 1

                # Handle string cells
                elif isinstance(val, str):
                    cleaned = val.strip()
                    if cleaned != val:
                        print(f"Removed extra spaces: '{val}' → '{cleaned}'")
                        cell.value = cleaned
                        cleaned_count += 1

        wb.save(file_path)
        wb.close()
        print(f"\nCleaning complete. Total cells updated: {cleaned_count}\n")

    except Exception as e:
        print(f"Error while cleaning Excel file: {e}")

# Example usage
clean_excel_dates("metadata/13190.xlsx")