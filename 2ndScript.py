from playwright.sync_api import sync_playwright
import time
import multiprocessing
import os
from datetime import datetime, timedelta
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import json

# Constants
PERMIT_TYPES = ["Fence Permit", "Retaining Wall Permit", "Building Residential", "Electrical Permit"]
LOG_FILE = "last_run_log.json"
SCRAPE_LOG_FILE = "scrape_log.txt"

# Function to load the last execution date
def load_last_execution_date():
    if os.path.exists(LOG_FILE):
        with open(LOG_FILE, 'r') as file:
            data = json.load(file)
            return datetime.strptime(data['last_execution_date'], '%Y-%m-%d').date()
    return None

# Function to save the last execution date
def save_last_execution_date(date):
    with open(LOG_FILE, 'w') as file:
        json.dump({'last_execution_date': date.strftime('%Y-%m-%d')}, file)

# Function to log scrape messages
def log_scrape_message(message):
    with open(SCRAPE_LOG_FILE, 'a') as file:
        file.write(message + "\n")

# Function to find the Google Service Account JSON Key
def find_json_key_file():
    for file in os.listdir(os.path.dirname(__file__)):
        if file.endswith('.json'):
            return os.path.join(os.path.dirname(__file__), file)
    raise FileNotFoundError("Google Service Account JSON Key not found in the script directory")

# Function to scrape application details
def scrape_application(args):
    app_number, date, permit_type = args
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=True)
        context = browser.new_context()
        page = context.new_page()

        try:
            page.goto("https://permittingservices.montgomerycountymd.gov/DPS/online/eSearch.aspx?by=Address&SearchType=DataSearch#")
            log_scrape_message(f"Processing {permit_type} application: {app_number} for date: {date}")

            page.wait_for_selector("#tabsNamesSelect", state="visible")
            page.select_option("#tabsNamesSelect", "tabsDate")
            time.sleep(2)

            page.evaluate(f"""() => {{
                document.getElementById('dpsTopSection_txtDate').value = '{date}';
                var event = new Event('change', {{ bubbles: true }});
                document.getElementById('dpsTopSection_txtDate').dispatchEvent(event);
            }}""")

            search_button = page.wait_for_selector("#dpsTopSection_cmdDateSearch", state="visible")
            search_button.click()

            page.wait_for_selector("#listDateSummary", state="visible", timeout=60000)

            permit_row = page.query_selector(f"#listDateSummary tr:has(td:text-is('{permit_type}'))")
            if permit_row:
                applications_link = permit_row.query_selector("td.clickLink")
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

                        result = [permit_type, app_number, application_date, status, site_address, square_footage, value, subdivision]
                        log_scrape_message(f"Scraped: {permit_type}, {app_number}, {application_date}, {status}, {site_address}, {square_footage}, {value}, {subdivision}")
                        return result
                    else:
                        log_scrape_message(f"Application {app_number} not found in the list")
                else:
                    log_scrape_message(f"No applications link found for {permit_type}")
            else:
                log_scrape_message(f"{permit_type} row not found")

        except Exception as e:
            log_scrape_message(f"An error occurred while processing {permit_type} application {app_number}: {str(e)}")

        finally:
            context.close()
            browser.close()

    return None

# Function to get application numbers for a specific date and permit type
def get_app_numbers_for_date(date, permit_type):
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=True)
        context = browser.new_context()
        page = context.new_page()

        app_numbers = []

        try:
            page.goto("https://permittingservices.montgomerycountymd.gov/DPS/online/eSearch.aspx?by=Address&SearchType=DataSearch#")
            page.wait_for_selector("#tabsNamesSelect", state="visible")
            page.select_option("#tabsNamesSelect", "tabsDate")
            time.sleep(2)

            page.evaluate(f"""() => {{
                document.getElementById('dpsTopSection_txtDate').value = '{date}';
                var event = new Event('change', {{ bubbles: true }});
                document.getElementById('dpsTopSection_txtDate').dispatchEvent(event);
            }}""")

            search_button = page.wait_for_selector("#dpsTopSection_cmdDateSearch", state="visible")
            search_button.click()

            page.wait_for_selector("#listDateSummary", state="visible", timeout=60000)

            permit_row = page.query_selector(f"#listDateSummary tr:has(td:text-is('{permit_type}'))")
            if permit_row:
                applications_link = permit_row.query_selector("td.clickLink")
                if applications_link:
                    applications_link.click()

                    page.wait_for_selector("#listDateDetail", state="visible", timeout=60000)

                    app_rows = page.query_selector_all("#listDateDetail tbody tr")
                    for row in app_rows:
                        app_link = row.query_selector("td.clickLink")
                        if app_link:
                            app_numbers.append(app_link.inner_text())

        finally:
            context.close()
            browser.close()

        return app_numbers

# Main function to run the scraper
def run():
    # Find the Google Service Account JSON key file
    json_key_path = find_json_key_file()

    # Set up Google Sheets credentials
    creds = Credentials.from_service_account_file(json_key_path, scopes=['https://www.googleapis.com/auth/spreadsheets'])
    service = build('sheets', 'v4', credentials=creds)

    # ID of your Google Sheet
    SHEET_ID = 'INSERT SHEET HERE' # example 'harekkgaloud43er5jahreng34'

    # Clear existing data (optional)
    clear_request = service.spreadsheets().values().clear(spreadsheetId=SHEET_ID, range='Sheet1')
    clear_request.execute()

    headers = ["Permit Type", "Application Number", "Application Date", "Status", "Site Address", "Square Footage", "Value", "Subdivision"]

    # Append headers
    service.spreadsheets().values().append(
        spreadsheetId=SHEET_ID,
        range='Sheet1',
        body={'values': [headers]},
        valueInputOption='USER_ENTERED'
    ).execute()

    print("Google Sheet created and headers added")
    log_scrape_message("Google Sheet created and headers added")

    last_execution_date = load_last_execution_date()
    end_date = datetime.now().date()
    start_date = end_date - timedelta(days=30) if last_execution_date is None else last_execution_date

    all_app_numbers = []

    for permit_type in PERMIT_TYPES:
        print(f"Starting scraping for {permit_type}")
        log_scrape_message(f"Starting scraping for {permit_type}")
        for i in range((end_date - start_date).days):
            current_date = (start_date + timedelta(days=i)).strftime("%m/%d/%Y")
            app_numbers = get_app_numbers_for_date(current_date, permit_type)
            all_app_numbers.extend([(app_number, current_date, permit_type) for app_number in app_numbers])
            print(f"Found {len(app_numbers)} applications for {permit_type} on {current_date}")
            log_scrape_message(f"Found {len(app_numbers)} applications for {permit_type} on {current_date}")
        print(f"Finished scraping for {permit_type}")
        log_scrape_message(f"Finished scraping for {permit_type}")

    print(f"Total applications to process: {len(all_app_numbers)}")
    log_scrape_message(f"Total applications to process: {len(all_app_numbers)}")

    # Use multiprocessing to scrape the applications
    with multiprocessing.Pool(processes=os.cpu_count()) as pool:
        results = pool.map(scrape_application, all_app_numbers)

    # Add results to Google Sheet
    data_to_append = [result for result in results if result]
    if data_to_append:
        service.spreadsheets().values().append(
            spreadsheetId=SHEET_ID,
            range='Sheet1',
            body={'values': data_to_append},
            valueInputOption='USER_ENTERED'
        ).execute()

    print("Data saved to Google Sheet")
    log_scrape_message("Data saved to Google Sheet")

    # Update last execution date
    save_last_execution_date(end_date - timedelta(days=1))

if __name__ == "__main__":
    multiprocessing.freeze_support()  # Needed for Windows
    run()
