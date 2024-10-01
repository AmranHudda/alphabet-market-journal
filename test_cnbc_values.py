import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
import pandas as pd
from datetime import datetime
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


CNBC_URLS = {
    "BRENT CRUDE": "https://www.cnbc.com/quotes/%40LCO.1",
    "DAX": "https://www.cnbc.com/quotes/.GDAXI",
    "BITCOIN": "https://www.cnbc.com/quotes/BTC.CM%3D"
}

def get_cnbc_value(url):
    try:
        headers = {'User-Agent': 'Mozilla/5.0'}
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        value = soup.find('span', {'class': 'QuoteStrip-lastPrice'}).text.strip()
        value = float(value.replace(',', '').replace('%', '').strip())
        return value
    except Exception as e:
        logging.error(f"Error fetching data from CNBC for {url}: {str(e)}")
        return None

# Fetch values for Brent Crude, DAX, and Bitcoin
cnbc_data = {}
for name, url in CNBC_URLS.items():
    value = get_cnbc_value(url)
    if value is not None:
        cnbc_data[name] = value
        logging.info(f"{name}: {value}")
    else:
        logging.warning(f"Failed to fetch value for {name}")