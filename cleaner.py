# cleaner.py

import re

import openpyxl
FILE_PATH = "metadata\13190.xlsx"
def clean_excel_dates(EXCEL_PATH):

    """

    Removes leading/trailing spaces from date-like cells in Excel.

    Only modifies cells that look like dates (dd/mm/yy or mm/dd/yyyy).

    """

    try:

        wb = openpyxl.load_workbook(EXCEL_PATH)

        ws = wb.active

        date_pattern = re.compile(r"\b\d{1,2}/\d{1,2}/\d{2,4}\b")

        cleaned = 0

        for row in ws.iter_rows():

            for cell in row:

                if isinstance(cell.value, str):

                    original = cell.value

                    stripped = original.strip()

                    # Check if cell value looks like a date and has extra spaces

                    if stripped != original and date_pattern.search(stripped):

                        cell.value = stripped

                        cleaned += 1

                        print(f" Cleaned: '{original}' → '{stripped}'")

        if cleaned == 0:

            print(" No date cells had extra spaces.")

        else:
            print(f" {cleaned} date cells cleaned successfully.")

        wb.save(EXCEL_PATH)

        wb.close()

    except Exception as e:

        print(f" Error while cleaning Excel file: {e}")
 