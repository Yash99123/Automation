from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import pandas as pd
import time


def initialize_driver():
    """Initialize the Selenium WebDriver with options."""
    chrome_options = Options()
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920x1080")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")  # Bypass bot detection
    chrome_options.add_argument("start-maximized")  # Open in full-screen mode

    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=chrome_options)


def scroll_page(driver, times=10):
    """Scroll down the page to load dynamic content."""
    body = driver.find_element(By.TAG_NAME, "body")
    for _ in range(times):
        body.send_keys(Keys.PAGE_DOWN)
        time.sleep(1)
    time.sleep(5)  # Allow JavaScript to load


def extract_table_data(soup):
    """Extracts table-based data if present."""
    tables = soup.find_all("table")
    extracted_data = []

    for table in tables:
        headers = [th.text.strip() for th in table.find_all("th")]
        rows = []

        tbody = table.find("tbody")
        if tbody:
            for row in tbody.find_all("tr"):
                cells = [td.text.strip() for td in row.find_all("td")]
                if cells:
                    rows.append(cells)

        if rows:
            extracted_data.append(pd.DataFrame(rows, columns=headers if headers else None))

    return extracted_data


def extract_div_data(soup):
    """Extracts div-based data specifically for Central Auction House."""
    bid_listings = soup.find_all("div", class_="listGroupWrapper clearfix")  # Correct div for bids
    extracted_data = []

    for bid in bid_listings:
        title_tag = bid.find("a")
        title = title_tag.text.strip() if title_tag else "N/A"
        link = title_tag["href"] if title_tag else "N/A"

        status_tag = bid.find("span", class_="badge badge-info")  # Optional: status like "Open"
        status = status_tag.text.strip() if status_tag else "N/A"

        agency_tag = bid.find("div", class_="col-md-6 text-left")  # Where the agency name is
        agency = agency_tag.text.strip() if agency_tag else "N/A"

        bid_info = bid.find_all("li", class_="list-inline-item")
        bid_id, due_date, broadcast_date, planholders = "N/A", "N/A", "N/A", "N/A"

        for info in bid_info:
            text = info.text.strip()
            if "ID:" in text:
                bid_id = text.replace("ID:", "").strip()
            elif "Due:" in text:
                due_date = text.replace("Due:", "").strip()
            elif "Broadcast:" in text:
                broadcast_date = text.replace("Broadcast:", "").strip()
            elif "#Planholders:" in text:
                planholders = text.replace("#Planholders:", "").strip()

        extracted_data.append([title, link, status, agency, bid_id, due_date, broadcast_date, planholders])

    return pd.DataFrame(extracted_data,
                        columns=["Title", "Link", "Status", "Agency", "Bid ID", "Due Date", "Broadcast Date",
                                 "Planholders"]) if extracted_data else None

def extract_web_data(url):
    """Extracts structured data from a webpage and saves it to an Excel file."""
    driver = initialize_driver()

    try:
        driver.get(url)
        print(f"\n🔄 Loading: {url}")

        # Scroll down to ensure content is loaded
        scroll_page(driver)
        # Get the page source after JavaScript execution
        soup = BeautifulSoup(driver.page_source, "html.parser")

        # Try extracting data in a tabular format
        table_data = extract_table_data(soup)

        if table_data:
            print("✅ Extracted table-based data.")
            with pd.ExcelWriter("extracted_data.xlsx") as writer:
                for i, df in enumerate(table_data):
                    df.to_excel(writer, sheet_name=f"Table_{i + 1}", index=False)

        # If no table found, try extracting div-based data
        else:
            print("⚠️ No tables found. Trying div-based extraction...")
            div_data = extract_div_data(soup)
            if div_data is not None:
                div_data.to_excel("extracted_data.xlsx", index=False)
                print("✅ Extracted div-based data.")

        print("\n✅ Data successfully extracted and saved to 'extracted_data.xlsx'")

    finally:
        driver.quit()


# Get the URL from the user
url = input("Enter the URL of the webpage: ")
extract_web_data(url)
