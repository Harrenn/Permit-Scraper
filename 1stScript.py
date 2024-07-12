from playwright.sync_api import sync_playwright
import time
from openpyxl import Workbook
import multiprocessing
import os
from datetime import datetime, timedelta
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import json

# Constants
LOG_FILE = "last_run_log.json"
SCRAPE_LOG_FILE = "scrape_log.txt"
SHEET_ID = 'INSERT SHEET HERE' # example 'jdsjfdsfjds3345656j56rt'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def find_json_key_file():
    for file in os.listdir(os.path.dirname(__file__)):
        if file.endswith('.json'):
            return os.path.join(os.path.dirname(__file__), file)
    raise FileNotFoundError("Google Service Account JSON Key not found in the script directory")

def get_dates():
    """Calculate dates from 09/03/1999 to yesterday's date."""
    start_date = datetime.strptime('09/03/1999', '%m/%d/%Y').date()
    end_date = datetime.today().date() - timedelta(days=1)
    dates = [(start_date + timedelta(days=i)).strftime('%m/%d/%Y') for i in range((end_date - start_date).days + 1)]
    return dates

def scrape_application(app_number, date):
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()
        try:
            page.goto("https://permittingservices.montgomerycountymd.gov/DPS/online/eSearch.aspx?by=Address&SearchType=DataSearch#")
            print(f"Processing application: {app_number} for date: {date}")
            page.wait_for_selector("#tabsNamesSelect", state="visible")
            page.select_option("#tabsNamesSelect", "tabsDate")
            time.sleep(2)
            page.evaluate("""(date) => {
                document.getElementById('dpsTopSection_txtDate').value = date;
                var event = new Event('change', { bubbles: true });
                document.getElementById('dpsTopSection_txtDate').dispatchEvent(event);
            }""", date)
            search_button = page.wait_for_selector("#dpsTopSection_cmdDateSearch", state="visible")
            search_button.click()
            page.wait_for_selector("#listDateSummary", state="visible", timeout=60000)
            fence_permit_row = page.query_selector("#listDateSummary tr:has(td:text-is('Fence Permit'))")
            if fence_permit_row:
                applications_link = fence_permit_row.query_selector("td.clickLink")
                if applications_link:
                    applications_link.click()
                    page.wait_for_selector("#listDateDetail", state="visible", timeout=60000)
                    target_app_link = page.query_selector(f"#listDateDetail tbody tr td.clickLink:text-is('{app_number}')")
                    if target_app_link:
                        target_app_link.click()
                        page.wait_for_selector("#divApplicationDate", state="visible", timeout=60000)
                        application_date = page.query_selector("td.head:text-is('Application Date') + td").inner_text()
                        status = page.query_selector("td.head:text-is('Status') + td").inner_text()
                        site_address = page.query_selector("td.head:text-is('Site Address') + td").inner_text().replace('<br>', ' ')
                        square_footage = page.query_selector("td.head:text-is('Square Footage') + td").inner_text()
                        value = page.query_selector("td.head:text-is('Value') + td").inner_text()
                        subdivision = page.query_selector("td.head:text-is('Subdivision') + td").inner_text()
                        result = [app_number, application_date, status, site_address, square_footage, value, subdivision]
                        print(f"Scraped: {app_number}, {application_date}, {status}, {site_address}, {square_footage}, {value}, {subdivision}")
                        return result
                    else:
                        print(f"Application {app_number} not found in the list")
                else:
                    print("No applications link found for Fence Permit")
            else:
                print("Fence Permit row not found")
        except Exception as e:
            print(f"An error occurred while processing {app_number}: {str(e)}")
        finally:
            context.close()
            browser.close()
        return None

def upload_to_google_sheets(data):
    json_key_path = find_json_key_file()
    creds = Credentials.from_service_account_file(json_key_path, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)
    
    # Clear existing data (optional)
    clear_request = service.spreadsheets().values().clear(spreadsheetId=SHEET_ID, range='Sheet1')
    clear_request.execute()

    # Append headers and data
    headers = ["Application Number", "Application Date", "Status", "Site Address", "Square Footage", "Value", "Subdivision"]
    data_to_append = [headers] + data
    service.spreadsheets().values().append(
        spreadsheetId=SHEET_ID,
        range='Sheet1',
        body={'values': data_to_append},
        valueInputOption='USER_ENTERED'
    ).execute()
    print("Data uploaded to Google Sheets")

def run():
    dates = get_dates()
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()
        wb = Workbook()
        ws = wb.active
        headers = ["Application Number", "Application Date", "Status", "Site Address", "Square Footage", "Value", "Subdivision"]
        ws.append(headers)
        print("Excel workbook created and headers added")
        app_numbers = []
        try:
            for date in dates:
                page.goto("https://permittingservices.montgomerycountymd.gov/DPS/online/eSearch.aspx?by=Address&SearchType=DataSearch#")
                page.wait_for_selector("#tabsNamesSelect", state="visible")
                page.select_option("#tabsNamesSelect", "tabsDate")
                time.sleep(2)
                page.evaluate("""(date) => {
                    document.getElementById('dpsTopSection_txtDate').value = date;
                    var event = new Event('change', { bubbles: true });
                    document.getElementById('dpsTopSection_txtDate').dispatchEvent(event);
                }""", date)
                search_button = page.wait_for_selector("#dpsTopSection_cmdDateSearch", state="visible")
                search_button.click()
                page.wait_for_selector("#listDateSummary", state="visible", timeout=60000)
                fence_permit_row = page.query_selector("#listDateSummary tr:has(td:text-is('Fence Permit'))")
                if fence_permit_row:
                    applications_link = fence_permit_row.query_selector("td.clickLink")
                    if applications_link:
                        applications_link.click()
                        page.wait_for_selector("#listDateDetail", state="visible", timeout=60000)
                        app_rows = page.query_selector_all("#listDateDetail tbody tr")
                        for row in app_rows:
                            app_link = row.query_selector("td.clickLink")
                            if app_link:
                                app_numbers.append((app_link.inner_text(), date))
        finally:
            context.close()
            browser.close()
        print(f"Found {len(app_numbers)} applications to process")
        # Use multiprocessing to scrape the applications
        with multiprocessing.Pool(processes=os.cpu_count()) as pool:
            results = pool.starmap(scrape_application, app_numbers)
        data = []
        for result in results:
            if result:
                ws.append(result)
                data.append(result)
        wb.save("fence_permit_applications_details.xlsx")
        print("Data saved to fence_permit_applications_details.xlsx")
        # Upload data to Google Sheets
        upload_to_google_sheets(data)

if __name__ == "__main__":
    multiprocessing.freeze_support()  # Needed for Windows
    run()
