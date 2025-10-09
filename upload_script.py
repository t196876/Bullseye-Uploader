import asyncio
import logging
import pandas as pd
from playwright.async_api import async_playwright
from dotenv import load_dotenv
import os
from datetime import datetime
from openpyxl import load_workbook

# CONFIGURATION


load_dotenv()

USERNAME = os.getenv("BULLSEYE_USERNAME")
PASSWORD = os.getenv("BULLSEYE_PASSWORD")
EXCEL_PATH = "data/metadata_mapping.xlsx"
LOG_FILE = "logs/upload_log.txt"
SUMMARY_FILE = "results/upload_summary.csv"
 #LOGIN INFO

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",



)


console = logging.StreamHandler()
console.setLevel(logging.INFO)
formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%H:%M:%S")
console.setFormatter(formatter)
logging.getLogger().addHandler(console)


#REMOVE WHITE SPACES
import re

from openpyxl import load_workbook

from datetime import datetime

def clean_excel_date_cells(excel_path):

    """Remove leading/trailing spaces from cells that are date-like (text or Excel date)."""

    wb = load_workbook(excel_path)

    sheet = wb.active

    print("Checking for extra spaces in date-like cells...")

    count = 0

    for row in sheet.iter_rows():

        for cell in row:

            val = cell.value

            # --- CASE 1: Text date (e.g., " 10/10/25 ")

            if isinstance(val, str) and any(x in val for x in ["/", "-"]):

                cleaned = val.strip()

                if cleaned != val:

                    print(f"🗓️ Cleaned text date: '{val}' → '{cleaned}'")

                    cell.value = cleaned

                    count += 1

            # --- CASE 2: Excel-stored date that was imported as string with padding

            elif isinstance(val, datetime):

                # Sometimes the cell has spaces due to formatting → convert to standard date format

                cell.number_format = "mm/dd/yy"

                count += 1

    wb.save(excel_path)

    print(f"Cleaning complete. Total date-like cells fixed: {count}")
 


async def run_upload():
    clean_excel_date_cells(EXCEL_PATH)
    df = pd.read_excel(EXCEL_PATH)

    summary = []

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False, slow_mo=200)
        page = await browser.new_page()
        try:

            logging.info("Logging into BE")
            await page.goto("https://bullseye.8451.com/ords/uskrgprh/f?p=138:LOGIN_DESKTOP::::::")
            await page.fill('input[name="P101_USERNAME"]', USERNAME)
            await page.fill('input[name="P101_PASSWORD"]', PASSWORD)
            await page.click("#P101_LOGIN")
            #await page.keyboard.press("Enter")
            await page.wait_for_timeout(4000)

            logging.info("Log-in successful")
            await page.wait_for_timeout(5000)

        except Exception as e:

            logging.error(f" Login Failed: {e}")
            return

        # Navigate to KPM Pipeline Tracker

        try:
            logging.info("opening KPM Pipeline Tracker")

            #await page.click('text="KPM Pipeline Tracker"')
            await page.click('a.dhtmlBottom:has-text("KPM Pipeline Tracker")')

            await page.wait_for_load_state('networkidle')

        except Exception as e:

            logging.error(f"Failed to Open KPM Pipeline Tracker: {e}")
            return

        # Loop through all rows

        for _, row in df.iterrows():

            kpm_id = str(row["KPM_ID"]).strip()
            offer_type = str(row["Offer_Type"]).strip().lower()
            file_path = str(row["File_Path"]).strip()
            logging.info(f"Processing KPM_ID: {kpm_id} | Offer Type: {offer_type}")

            try:

                # Search KPM ID

                await page.fill('input[id$="_search_field"]', str(kpm_id))
                await page.keyboard.press("Enter")
                await page.wait_for_timeout(2000)

                # Click the KPM ID link

                #await page.click(f'a.a-IRR-reportSummary-value:has-text("kpm_id")',timeout=30000)
               ##await page.click(f'td a:has-text("kpm_id")')
                await page.wait_for_selector(f'td a:has-text("{kpm_id}")', timeout=60000)
                element = await page.query_selector(f'td a:has-text("{kpm_id}")')
                await element.scroll_into_view_if_needed()
                await page.evaluate("(el) => el.click()", element)

                await page.wait_for_load_state('networkidle')

                # Navigate to Load Direct Mail Offers

                await page.click('text="Load Direct Mail Offers"')

                await page.wait_for_timeout(2000)

                # Choose Offer Type

                if "mixed" in offer_type:

                    await page.check('input[value="Mixed Offers"]')

                elif "spend" in offer_type:

                    await page.check('input[value="Spend X Get Y"]')

                else:

                    await page.wait_for_selector('label:has-text("Static Divisional Offer")', timeout=60000)
                    await page.click('label:has-text("Static Divisional Offer")')
                    logging.info("Offer Type 'Static Divisional Offer' selected.")

                # Upload File

                #await page.set_input_files('input[type="file"]', file_path)
                await page.set_input_files('#BOM_FILE_BROWSER', file_path)
                logging.info(f"Uploaded file for KPM_ID {kpm_id}: {file_path}")

                await page.wait_for_timeout(1000)

                # kip Submit for test phase

                logging.info(f"[..TEST MODE..] File uploaded but not submitted: {file_path}")

                summary.append({

                    "KPM_ID": kpm_id,

                    "File": file_path,

                    "Status": "Uploaded (Test Mode)",

                    "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                })

                # Navigate back

                await page.go_back()
                await page.wait_for_timeout(2000)

            except Exception as e:

                logging.error(f"Failed for KPM_ID {kpm_id}: {e}") #for recording
                summary.append({

                    "KPM_ID": kpm_id,
                    "File": file_path,
                    "Status": f"Failed - {e}",
                    "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                })

                continue

        # Save summary report

        pd.DataFrame(summary).to_csv(SUMMARY_FILE, index=False)
        logging.info(f"Summary report saved to {SUMMARY_FILE}")
        logging.info("Process completed successfully.")
        await browser.close()


if __name__ == "__main__":
    asyncio.run(run_upload())
