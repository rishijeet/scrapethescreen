## Author Rishijeet Mishra
## Change the company URL for now, will enhance to use to be used as param

import pandas as pd
import requests
from bs4 import BeautifulSoup
import logging
import sys



# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def get_soup(url: str):
    """
    Fetches the webpage content and returns a BeautifulSoup object.
    
    Args:
        url: The URL of the webpage to scrape.
        
    Returns:
        A BeautifulSoup object or None if the request fails.
    """
    try:
        # Using a User-Agent to mimic a browser
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()  # Raises an HTTPError for bad responses (4xx or 5xx)
        return BeautifulSoup(response.text, 'html.parser')
    except requests.exceptions.RequestException as e:
        logging.error(f"Error fetching URL {url}: {e}")
        return None

def scrape_table_to_dataframe(soup: BeautifulSoup, section_id: str) -> pd.DataFrame | None:
    """
    Finds a table within a specific section by its ID and scrapes it into a pandas DataFrame.

    Args:
        soup: The BeautifulSoup object of the page.
        section_id: The 'id' attribute of the <section> tag containing the table.

    Returns:
        A pandas DataFrame containing the scraped data, or None if the table is not found.
    """
    section = soup.find('section', id=section_id)
    if not section:
        logging.warning(f"Could not find section with id: {section_id}")
        return None

    table = section.find('table', class_='data-table')
    if not table:
        logging.warning(f"Could not find a data table in section: {section_id}")
        return None

    # Extract headers
    headers = [th.get_text(strip=True) for th in table.find('thead').find_all('th')]

    # Extract rows
    data_rows = []
    for row in table.find('tbody').find_all('tr'):
        cols = [td.get_text(strip=True) for td in row.find_all('td')]
        data_rows.append(cols)

    if not data_rows:
        logging.warning(f"No data rows found in table for section: {section_id}")
        return None

    # Create DataFrame
    df = pd.DataFrame(data_rows, columns=headers)
    return df

def main():
    """
    Main function to orchestrate the scraping and Excel file creation.
    """
    company_url = "https://www.screener.in/company/HDFCBANK/consolidated/"
    output_filename = "HDFCBANK_consolidated_financials.xlsx"

    logging.info(f"Starting scrape for {company_url}")
    soup = get_soup(company_url)
    if not soup:
        sys.exit(1) # Exit if page fetch failed

    # List of section IDs for the tables we want to scrape
    table_sections = [
        'quarters',
        'profit-loss',
        'balance-sheet',
        'cash-flow',
        'ratios'
    ]

    try:
        with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
            logging.info(f"Writing data to {output_filename}")
            for section_id in table_sections:
                logging.info(f"Scraping section: {section_id}...")
                df = scrape_table_to_dataframe(soup, section_id)
                if df is not None:
                    # Use a clean sheet name from the section ID
                    sheet_name = section_id.replace('-', ' ').title()
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    logging.info(f"Successfully wrote '{sheet_name}' sheet.")
                else:
                    logging.warning(f"Skipping section {section_id} as no data was found.")
        
        logging.info(f"Excel file '{output_filename}' created successfully.")
    except Exception as e:
        logging.error(f"An error occurred while writing the Excel file: {e}")

if __name__ == "__main__":
    main()