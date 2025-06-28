# Screener.in Financial Data Scraper

This project contains a Python script to scrape financial data for a specific company from [screener.in](https://www.screener.in) and save it into a consolidated Excel file. Each financial statement (e.g., Quarterly Results, Profit & Loss) is saved to a separate sheet within the Excel workbook.

## Features

- Scrapes data from a specified company page on screener.in.
- Extracts multiple tables: Quarters, Profit & Loss, Balance Sheet, Cash Flow, and Ratios.
- Saves all extracted data into a single, well-organized Excel file.
- Robust error handling and logging.

## Prerequisites

- Python 3.8+

## How to Run

1.  **Clone the repository or create the files** as shown in the project structure.

2.  **Install the required dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

3.  **Run the scraper:**
    ```bash
    python scraper.py
    ```

4.  **Check the output:** A file named `HDFCBANK_consolidated_financials.xlsx` will be created in the same directory.