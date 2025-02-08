import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import time


# Function to fetch HTML content using Selenium
def fetch_html_with_selenium(url):
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")  # Run in headless mode
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    # Set up WebDriver
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.get(url)

    # Wait for the table to load inside the scrollable div
    try:
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "DataTables_Table_0"))
        )
        time.sleep(5)  # Extra wait for JavaScript to load data

        # **Scroll inside the table container**
        scroll_div = driver.find_element(By.CLASS_NAME, "dataTables_scrollBody")
        driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight;", scroll_div)
        time.sleep(3)  # Wait for additional rows to load
    finally:
        html_content = driver.page_source
        driver.quit()

    return html_content


# Function to extract table data dynamically from any website
def extract_table_from_url(url):
    print(f"Extracting data from: {url}")

    html_content = fetch_html_with_selenium(url)

    if not html_content:
        print("Failed to retrieve content.")
        return

    # Parse the HTML content using BeautifulSoup
    soup = BeautifulSoup(html_content, 'html.parser')

    # Locate the specific table by ID
    table = soup.find("table", {"id": "DataTables_Table_0"})
    if not table:
        print(" Table with ID 'DataTables_Table_0' not found.")
        return

    print(" Table found. Extracting data...")

    # Extract table headers dynamically (if available)
    headers = [th.text.strip() for th in table.find_all("th")]

    # If headers are missing, create generic headers
    first_row = table.find("tbody").find("tr")
    if not headers and first_row:
        num_columns = len(first_row.find_all("td"))
        headers = [f"Column {i + 1}" for i in range(num_columns)]

    # Extract table rows
    data = []
    tbody = table.find("tbody")
    rows = tbody.find_all("tr") if tbody else []

    if not rows:
        print("No table rows found. The table might be empty or dynamically loaded.")
        return

    print(f"Found {len(rows)} rows in the table.")

    for row in rows:
        columns = row.find_all("td")
        if columns:
            data.append([col.text.strip() for col in columns])

    # Convert data into a pandas DataFrame
    df = pd.DataFrame(data, columns=headers)

    # Save the extracted data to an Excel file
    df.to_excel("output.xlsx", index=False)
    print("Data successfully saved to output.xlsx")


# Run the script dynamically for any URL
if __name__ == "__main__":
    url = input("Enter the URL to extract data from: ")
    extract_table_from_url(url)
