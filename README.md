# Permit Scraper

## Description

This script scrapes permit application data from the Montgomery County permitting services website and saves the data to a Google Sheet. It can scrape data for specified permit types and maintain a log of its last execution to only fetch new data on subsequent runs.

## Features

*   **Initial Run:** On the first run, the script will scrape permit data for the last 30 days (excluding the present day).
*   **Subsequent Runs:** On subsequent runs, the script will only scrape data from the last execution date to yesterday.
*   **Logging:** The script maintains two logs:
    *   `last_run_log.json`: Tracks the last execution date.
    *   `scrape_log.txt`: Logs important scraping messages and progress updates.

## Requirements

*   Python 3.x
*   Playwright
*   Google API Client
*   A Google Service Account with Sheets API enabled

## Installation

1.  **Clone the repository:**

```bash
git clone <repository_url>
cd <repository_directory>
```

2.  **Install required Python packages:**

```bash
pip install playwright google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client
```

3.  **Install Playwright browsers:**

```bash
playwright install
```

4.  **Set up Google Sheets API:**

    *   Create a Google Service Account and download the JSON key file.
    *   Share the target Google Sheet with the service account email.
    *   Save the JSON key file: Place your Google Service Account JSON key file in the appropriate directory and update the path in the script accordingly.

## Configuration

### Constants

*   `PERMIT_TYPES`: Edit this list to include the permit types you want to scrape.
*   `LOG_FILE`: Path to the last execution log file.
*   `SCRAPE_LOG_FILE`: Path to the scrape log file.
*   `SHEET_ID`: Update this with the ID of your Google Sheet.

### Google Service Account JSON Key

```python
creds = Credentials.from_service_account_file('/path/to/your/credentials.json', scopes=['https://www.googleapis.com/auth/spreadsheets'])
```

## Usage

Run the script using:

```bash
python your_script.py
```

Example:

```bash
python permit_scraper.py
```

## Script Overview

*   `load_last_execution_date()`: Reads the last execution date from the log file.
*   `save_last_execution_date(date)`: Saves the current execution date to the log file.
*   `log_scrape_message(message)`: Logs important scraping messages to the `scrape_log.txt` file.
*   `scrape_application(args)`: Handles the browser automation for scraping individual permit application details.
*   `get_app_numbers_for_date(date, permit_type)`: Fetches a list of application numbers for a specific date and permit type by querying the website.
*   `run()`: Main function that:
    *   Authenticates with Google Sheets API.
    *   Clears existing data in the Google Sheet.
    *   Sets headers for the sheet.
    *   Defines the scraping date range.
    *   Compiles a list of application numbers for each permit type within the date range.
    *   Uses multiprocessing to scrape permit details.
    *   Appends the scraped data to the Google Sheet.
    *   Logs the progress and updates the last execution date.

## License

This project is licensed under the MIT License. See the LICENSE file for details.

## Acknowledgements

*   Playwright
*   Google Sheets API

For any issues or contributions, feel free to open an issue or submit a pull request.
