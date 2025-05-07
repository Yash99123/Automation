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
import re
import os


def initialize_driver():
    """Initialize the Selenium WebDriver with options."""
    chrome_options = Options()
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920x1080")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")  # Bypass bot detection
    chrome_options.add_argument("start-maximized")  # Open in full-screen mode
    chrome_options.add_argument("--headless")  # Run in headless mode to avoid showing browser

    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=chrome_options)


def scroll_page(driver, times=10):
    """Scroll down the page to load dynamic content."""
    body = driver.find_element(By.TAG_NAME, "body")
    for _ in range(times):
        body.send_keys(Keys.PAGE_DOWN)
        time.sleep(1)
    time.sleep(5)  # Allow JavaScript to load


def extract_table_data(soup, source_url):
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
                    # Add source URL to each row
                    cells.append(source_url)
                    rows.append(cells)

        if rows:
            if headers:
                headers.append("Source URL")
            extracted_data.append(pd.DataFrame(rows, columns=headers if headers else None))

    return extracted_data


def extract_div_data(soup, source_url):
    """Extracts div-based data specifically for Central Auction House."""
    bid_listings = soup.find_all("div", class_="listGroupWrapper clearfix")  # Correct div for bids
    extracted_data = []

    for bid in bid_listings:
        title_tag = bid.find("a")
        title = title_tag.text.strip() if title_tag else "N/A"
        link = title_tag["href"] if title_tag else "N/A"

        # Make link absolute if it's relative
        if link != "N/A" and not link.startswith(("http://", "https://")):
            base_url = "/".join(source_url.split("/")[:3])  # Gets domain like https://example.com
            link = f"{base_url}{link}" if link.startswith("/") else f"{base_url}/{link}"

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

        extracted_data.append([title, link, status, agency, bid_id, due_date, broadcast_date, planholders, source_url])

    return pd.DataFrame(extracted_data,
                        columns=["Title", "Link", "Status", "Agency", "Bid ID", "Due Date", "Broadcast Date",
                                 "Planholders", "Source URL"]) if extracted_data else None


def extract_generic_data(soup, source_url):
    """Extract data from generic RFP sites using various techniques."""
    all_data = []

    # Look for common RFP listing patterns
    # Pattern 1: Divs with listings
    rfp_divs = soup.find_all("div", class_=lambda c: c and any(
        term in c.lower() for term in ["listing", "item", "rfp", "bid", "opportunity"]))

    # Pattern 2: Lists with listings
    rfp_lists = soup.find_all(["ul", "ol"], class_=lambda c: c and any(
        term in c.lower() for term in ["listing", "list", "rfp", "bid", "opportunity"]))
    rfp_list_items = []
    for lst in rfp_lists:
        rfp_list_items.extend(lst.find_all("li"))

    # Process div-based listings
    for div in rfp_divs:
        title_elem = div.find(["h2", "h3", "h4", "a", "strong"])
        title = title_elem.text.strip() if title_elem else "N/A"

        link_elem = div.find("a")
        link = link_elem["href"] if link_elem and "href" in link_elem.attrs else "N/A"
        if link != "N/A" and not link.startswith(("http://", "https://")):
            base_url = "/".join(source_url.split("/")[:3])  # Gets domain
            link = f"{base_url}{link}" if link.startswith("/") else f"{base_url}/{link}"

        # Try to find status (open, closed, etc.)
        status_elem = div.find(["span", "div"], string=lambda s: s and any(
            word in s.lower() for word in ["open", "active", "closed", "pending"]))
        status = status_elem.text.strip() if status_elem else "N/A"

        # Look for date information
        date_patterns = ["due", "close", "deadline", "open until", "submit by"]
        date_elem = div.find(string=lambda s: s and any(pattern in s.lower() for pattern in date_patterns))
        due_date = date_elem if date_elem else "N/A"

        # Look for agency/organization info
        agency_elem = div.find(["span", "div"], string=lambda s: s and any(
            word in s.lower() for word in ["agency", "department", "organization", "posted by"]))
        agency = agency_elem.text.strip() if agency_elem else "N/A"

        description = div.get_text(strip=True)[:200] + "..." if len(div.get_text(strip=True)) > 200 else div.get_text(
            strip=True)

        all_data.append([title, link, status, agency, "N/A", due_date, "N/A", "N/A", source_url, description])

    # Process list-based listings
    for item in rfp_list_items:
        title_elem = item.find(["h2", "h3", "h4", "a", "strong"])
        title = title_elem.text.strip() if title_elem else item.text.strip()

        link_elem = item.find("a")
        link = link_elem["href"] if link_elem and "href" in link_elem.attrs else "N/A"
        if link != "N/A" and not link.startswith(("http://", "https://")):
            base_url = "/".join(source_url.split("/")[:3])  # Gets domain
            link = f"{base_url}{link}" if link.startswith("/") else f"{base_url}/{link}"

        description = item.get_text(strip=True)[:200] + "..." if len(
            item.get_text(strip=True)) > 200 else item.get_text(strip=True)

        all_data.append([title, link, "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", source_url, description])

    if not all_data:
        return None

    return pd.DataFrame(all_data, columns=["Title", "Link", "Status", "Agency", "Bid ID", "Due Date",
                                           "Broadcast Date", "Planholders", "Source URL", "Description"])


def is_it_cybersecurity_related(text):
    """Check if the text is related to IT or cybersecurity."""
    it_cyber_keywords = [
        # IT keywords
        "information technology", "it service", "software", "hardware", "network", "computer",
        "server", "database", "cloud", "infrastructure", "it support", "helpdesk", "help desk",
        "data center", "datacenter", "system", "application", "programming", "developer",
        "website", "web service", "api", "integration", "digital", "automation", "devops",

        # Cybersecurity keywords
        "cyber", "security", "firewall", "cyber security", "cybersecurity", "infosec",
        "information security", "data protection", "encryption",
        #staffing
        "Staffing", "onboarding"
        #machine learning
        "Machine learning", "ML Engineer", "ML ops"
        #AI
    ]

    if not isinstance(text, str):
        return False

    text_lower = text.lower()

    # Check if any keyword is in the text
    return any(keyword in text_lower for keyword in it_cyber_keywords)


def is_active_rfp(status_text):
    """Check if the RFP status is active/open."""
    if not isinstance(status_text, str):
        return False

    active_keywords = ["active", "open", "ongoing", "current", "live", "accepting", "available"]
    inactive_keywords = ["closed", "awarded", "complete", "expired", "ended", "canceled", "cancelled"]

    status_lower = status_text.lower()

    # If explicitly marked as inactive, return False
    if any(keyword in status_lower for keyword in inactive_keywords):
        return False

    # If explicitly marked as active or if no clear status, assume it's active
    return any(keyword in status_lower for keyword in active_keywords) or status_lower == "n/a"


def extract_web_data(url, driver):
    """Extracts structured data from a webpage."""
    try:
        driver.get(url)
        print(f"\n🔄 Loading: {url}")

        # Scroll down to ensure content is loaded
        scroll_page(driver)
        # Get the page source after JavaScript execution
        soup = BeautifulSoup(driver.page_source, "html.parser")

        # Try extracting data using different methods
        all_dataframes = []

        # 1. Try extracting table-based data
        table_data = extract_table_data(soup, url)
        if table_data:
            print(f"✅ Extracted {len(table_data)} table(s) from {url}")
            all_dataframes.extend(table_data)

        # 2. Try extracting div-based data for specific formats (Central Auction House)
        div_data = extract_div_data(soup, url)
        if div_data is not None and not div_data.empty:
            print(f"✅ Extracted div-based data ({len(div_data)} entries) from {url}")
            all_dataframes.append(div_data)

        # 3. Try generic data extraction if previous methods didn't yield results
        if not all_dataframes:
            generic_data = extract_generic_data(soup, url)
            if generic_data is not None and not generic_data.empty:
                print(f"✅ Extracted generic data ({len(generic_data)} entries) from {url}")
                all_dataframes.append(generic_data)
            else:
                print(f"⚠️ No data could be extracted from {url}")

        return all_dataframes

    except Exception as e:
        print(f"❌ Error processing {url}: {str(e)}")
        return []


def filter_it_cybersecurity_active(df):
    """Filter dataframe to include only IT/cybersecurity-related and active RFPs."""
    if "Description" not in df.columns:
        df["Description"] = ""  # Add description column if not present

    # Combine title and description for keyword search
    df["Combined_Text"] = df["Title"] + " " + df["Description"]

    # Filter by IT/Cybersecurity keywords
    it_cyber_mask = df["Combined_Text"].apply(is_it_cybersecurity_related)

    # Filter by active status
    active_mask = df["Status"].apply(is_active_rfp)

    # Apply both filters
    filtered_df = df[it_cyber_mask & active_mask]

    # Drop the temporary column
    filtered_df = filtered_df.drop(columns=["Combined_Text"])

    return filtered_df


def extract_and_filter_rfps(urls):
    """Process multiple URLs, extract RFP data, and filter for IT/cybersecurity active RFPs."""
    driver = initialize_driver()
    all_extracted_data = []

    try:
        for url in urls:
            dataframes = extract_web_data(url, driver)
            all_extracted_data.extend(dataframes)

        if not all_extracted_data:
            print("\n❌ No data was extracted from any of the provided URLs.")
            return

        # Combine all extracted dataframes
        combined_df = pd.concat(all_extracted_data, ignore_index=True)

        # Ensure consistent column names
        required_columns = ["Title", "Link", "Status", "Agency", "Bid ID", "Due Date",
                            "Broadcast Date", "Planholders", "Source URL"]
        for col in required_columns:
            if col not in combined_df.columns:
                combined_df[col] = "N/A"

        # Filter for IT/cybersecurity and active RFPs
        filtered_df = filter_it_cybersecurity_active(combined_df)

        # Check if we have any results after filtering
        if filtered_df.empty:
            print("\n⚠️ No IT/cybersecurity active RFPs found after filtering.")
            # Save original data for reference
            combined_df.to_excel("all_extracted_rfps.xlsx", index=False)
            print("✅ All extracted data (unfiltered) saved to 'all_extracted_rfps.xlsx'")
        else:
            # Save the filtered data
            filtered_df.to_excel("it_cybersecurity_active_rfps.xlsx", index=False)
            # Also save the original data for reference
            combined_df.to_excel("all_extracted_rfps.xlsx", index=False)

            print(f"\n✅ Found {len(filtered_df)} IT/cybersecurity active RFPs from {len(urls)} websites.")
            print("✅ Filtered data saved to 'it_cybersecurity_active_rfps.xlsx'")
            print("✅ All extracted data (unfiltered) saved to 'all_extracted_rfps.xlsx'")

    finally:
        driver.quit()


def main():
    print("===== RFP Scraper: IT & Cybersecurity Filter =====")
    print("This tool extracts RFPs from multiple websites and filters for IT/cybersecurity active RFPs.")

    urls = []
    print("\nEnter the URLs of the webpages (one per line, press Enter twice when done):")

    while True:
        url = input().strip()
        if not url:  # Empty line
            break
        if url.startswith(("http://", "https://")):
            urls.append(url)
        else:
            print(f"⚠️ Invalid URL format: {url} - must start with http:// or https://")

    if not urls:
        print("❌ No valid URLs provided. Exiting.")
        return

    print(f"\n📋 Processing {len(urls)} URLs...")
    extract_and_filter_rfps(urls)


if __name__ == "__main__":
    main()