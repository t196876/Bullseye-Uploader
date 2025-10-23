import asyncio
import logging
import pandas as pd
from playwright.async_api import async_playwright
from dotenv import load_dotenv
import os
from datetime import datetime
from openpyxl import load_workbook
import openpyxl
from cleaner import clean_excel_dates
import re

# CONFIGURATION
load_dotenv()

USERNAME = os.getenv("BULLSEYE_USERNAME")
PASSWORD = os.getenv("BULLSEYE_PASSWORD")
EXCEL_PATH = r"C:\Users\t196876\Documents\bullseye_uploader\test"
LOG_FILE = "logs/upload_log.txt"
SUMMARY_FILE = "results/upload_summary.csv"

# LOGIN INFO
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


async def run_upload():
    excel_files = [f for f in os.listdir(EXCEL_PATH) if f.endswith(".xlsx")]
    if not excel_files:
        print("No files are present in the provided folder")
        return

    summary = []

    for file_name in excel_files:
        match = re.search(r"_(\d+)\.xlsx$", file_name)
        if not match:
            print(f"Skipping {file_name} (KPM_ID not found in filename).")
            continue

        kpm_id = match.group(1)
        file_path = os.path.join(EXCEL_PATH, file_name)
        print(f"\nProcessing file {file_name} | KPM_ID: {kpm_id}")

        # Clean Excel file if needed
        try:
            clean_excel_dates(file_path)
        except Exception as e:
            print(f"Error while cleaning Excel file: {e}")

        # Offer type logic (default or adjust if needed)
        offer_type = "static"  # Default since we're not reading Excel rows

        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=False, slow_mo=200)
            page = await browser.new_page()

            try:
                logging.info("Logging into BE")
                await page.goto("https://bullseye.8451.com/ords/uskrgprh/f?p=138:LOGIN_DESKTOP::::::")
                await page.fill('input[name="P101_USERNAME"]', USERNAME)
                await page.fill('input[name="P101_PASSWORD"]', PASSWORD)
                await page.click("#P101_LOGIN")
                await page.wait_for_timeout(4000)

                logging.info("Log-in successful")
                await page.wait_for_timeout(5000)

            except Exception as e:
                logging.error(f"Login Failed: {e}")
                return

            # Navigate to KPM Pipeline Tracker
            try:
                logging.info("Opening KPM Pipeline Tracker")
                await page.click('a.dhtmlBottom:has-text("KPM Pipeline Tracker")')
                await page.wait_for_load_state('networkidle')
            except Exception as e:
                logging.error(f"Failed to Open KPM Pipeline Tracker: {e}")
                return

            # Process single KPM_ID from filename
            try:
                logging.info(f"Processing KPM_ID: {kpm_id} | Offer Type: {offer_type}")

                # Search KPM ID
                await page.fill('input[id$="_search_field"]', str(kpm_id))
                await page.keyboard.press("Enter")
                await page.wait_for_timeout(2000)

                # Click the KPM ID link
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

                await page.set_input_files('input#P18_DM_FILE_BROWSE', file_path)
                await page.wait_for_timeout(2000)
                logging.info(f"Uploaded file for KPM_ID {kpm_id}: {file_path}")

                await page.wait_for_timeout(1000)
                logging.info(f"[..TEST MODE..] File uploaded but not submitted: {file_path}")
                

                #CURRENTLY UNABLED
                #await page.click('a.effo-uButtonSmall:has-text("Submit Offers")')
                #await page.wait_for_timeout(2000)

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
                logging.error(f"Failed for KPM_ID {kpm_id}: {e}")
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