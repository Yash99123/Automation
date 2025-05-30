import tkinter as tk
from tkinter import filedialog
import pandas as pd
import time
import re
import os
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup


class RFPExtractor:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()  # Hide the main window
        self.excel_data = None
        self.output_directory = os.path.join(os.path.expanduser("~"), "Desktop", "RFP_Extractions")
        self.rfp_type_keywords = []
        self.status_filter_enabled = False

        # Create output directory if it doesn't exist
        if not os.path.exists(self.output_directory):
            os.makedirs(self.output_directory)

    def select_file(self):
        """Open a file dialog for user to select an Excel file"""
        file_path = filedialog.askopenfilename(
            title="Select Excel File with URLs",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )

        if not file_path:  # User canceled the dialog
            print("No file selected. Exiting...")
            return False

        print(f"Selected file: {file_path}")
        try:
            self.excel_data = pd.read_excel(file_path)
            # Validate that the required column exists
            if 'URL' not in self.excel_data.columns:
                print("Error: Missing required 'URL' column in Excel file.")
                print("The Excel file must contain a 'URL' column.")
                return False

            return True
        except Exception as e:
            print(f"Error reading Excel file: {e}")
            return False

    def setup_filters(self):
        """Setup the two simple filters"""
        print("\n===== FILTER SETUP =====")

        # Filter 1: RFP Type Selection
        print("\n1. RFP TYPE FILTER")
        print("What type of RFPs are you looking for?")
        print("Examples: construction, IT services, consulting, equipment, maintenance, etc.")

        rfp_type_input = input("Enter RFP type keywords (comma-separated, leave empty for all types): ").strip()
        if rfp_type_input:
            self.rfp_type_keywords = [kw.strip().lower() for kw in rfp_type_input.split(',')]
            print(f"   ✓ Looking for: {', '.join(self.rfp_type_keywords)}")
        else:
            print("   ✓ No type filter - all RFP types will be included")

        # Filter 2: Status Filter (Open/Active/Pending only)
        print("\n2. STATUS FILTER")
        status_choice = input("Only show OPEN/ACTIVE/PENDING RFPs? (exclude closed/expired) (y/n): ").strip().lower()

        if status_choice == 'y':
            self.status_filter_enabled = True
            print("   ✓ Will exclude closed/expired RFPs")
        else:
            self.status_filter_enabled = False
            print("   ✓ All statuses will be included")

        print("\n✅ Filter setup complete!")
        return True

    def apply_filters(self, df):
        """Apply the two configured filters to the dataframe"""
        if df is None or df.empty:
            return df

        original_count = len(df)
        filtered_df = df.copy().reset_index(drop=True)  # Reset index to avoid alignment issues

        # Filter 1: RFP Type Filter
        if self.rfp_type_keywords:
            type_indices = []

            for idx, row in filtered_df.iterrows():
                # Combine title, description, and other text fields
                row_text = ' '.join(str(value).lower() for value in row.values if pd.notna(value))

                # Check if any of the RFP type keywords are present
                if any(keyword in row_text for keyword in self.rfp_type_keywords):
                    type_indices.append(idx)

            filtered_df = filtered_df.iloc[type_indices].reset_index(drop=True)
            print(f"   🔍 RFP Type filter: {len(filtered_df)} entries match your criteria (from {original_count})")

        # Filter 2: Status Filter (exclude closed/expired)
        if self.status_filter_enabled and not filtered_df.empty:
            status_indices = []

            # Keywords that indicate closed/expired RFPs
            closed_keywords = ['closed', 'expired', 'cancelled', 'canceled', 'awarded', 'completed', 'inactive']

            for idx, row in filtered_df.iterrows():
                row_text = ' '.join(str(value).lower() for value in row.values if pd.notna(value))

                # Default: keep the entry
                keep_entry = True

                # If any closed keywords are found, check further
                if any(keyword in row_text for keyword in closed_keywords):
                    # Double check - if it also has open/active/pending, keep it
                    open_keywords = ['open', 'active', 'pending', 'available', 'current']
                    if not any(open_keyword in row_text for open_keyword in open_keywords):
                        keep_entry = False

                if keep_entry:
                    status_indices.append(idx)

            filtered_df = filtered_df.iloc[status_indices].reset_index(drop=True)
            print(f"   📊 Status filter: {len(filtered_df)} open/active/pending entries remaining")

        if len(filtered_df) < original_count:
            print(f"   ✅ Total filtered out: {original_count - len(filtered_df)} entries")

        return filtered_df

    def initialize_driver(self):
        """Initialize the Selenium WebDriver with options."""
        chrome_options = Options()
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920x1080")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")  # Bypass bot detection
        chrome_options.add_argument("start-maximized")  # Open in full-screen mode
        chrome_options.add_argument("--headless")  # Run in headless mode to avoid showing browser

        service = Service(ChromeDriverManager().install())
        return webdriver.Chrome(service=service, options=chrome_options)

    def scroll_page(self, driver, times=10):
        """Scroll down the page to load dynamic content."""
        body = driver.find_element(By.TAG_NAME, "body")
        for _ in range(times):
            body.send_keys(Keys.PAGE_DOWN)
            time.sleep(1)
        time.sleep(5)  # Allow JavaScript to load

    def extract_table_data(self, soup, source_url):
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
            else:
                # If no tbody, get all rows
                for row in table.find_all("tr"):
                    cells = [td.text.strip() for td in row.find_all("td")]
                    if cells:
                        # Add source URL to each row
                        cells.append(source_url)
                        rows.append(cells)

            if rows:
                if headers:
                    headers.append("Source URL")
                    df = pd.DataFrame(rows, columns=headers)
                else:
                    # Create generic column names if no headers
                    num_cols = len(rows[0]) if rows else 0
                    headers = [f"Column_{i + 1}" for i in range(num_cols)]
                    df = pd.DataFrame(rows, columns=headers)

                extracted_data.append(df)

        return extracted_data

    def extract_div_data(self, soup, source_url):
        """Extracts div-based data specifically for Central Auction House and similar sites."""
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

            extracted_data.append(
                [title, link, status, agency, bid_id, due_date, broadcast_date, planholders, source_url])

        return pd.DataFrame(extracted_data,
                            columns=["Title", "Link", "Status", "Agency", "Bid ID", "Due Date", "Broadcast Date",
                                     "Planholders", "Source URL"]) if extracted_data else None

    def extract_generic_data(self, soup, source_url):
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

            description = div.get_text(strip=True)[:200] + "..." if len(
                div.get_text(strip=True)) > 200 else div.get_text(
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

    def extract_web_data(self, url, driver):
        """Extracts structured data from a webpage."""
        try:
            driver.get(url)
            print(f"\n🔄 Loading: {url}")

            # Scroll down to ensure content is loaded
            self.scroll_page(driver)
            # Get the page source after JavaScript execution
            soup = BeautifulSoup(driver.page_source, "html.parser")

            # Try extracting data using different methods
            all_dataframes = []

            # 1. Try extracting table-based data
            table_data = self.extract_table_data(soup, url)
            if table_data:
                print(f"✅ Extracted {len(table_data)} table(s) from {url}")
                all_dataframes.extend(table_data)

            # 2. Try extracting div-based data for specific formats (Central Auction House)
            div_data = self.extract_div_data(soup, url)
            if div_data is not None and not div_data.empty:
                print(f"✅ Extracted div-based data ({len(div_data)} entries) from {url}")
                all_dataframes.append(div_data)

            # 3. Try generic data extraction if previous methods didn't yield results
            if not all_dataframes:
                generic_data = self.extract_generic_data(soup, url)
                if generic_data is not None and not generic_data.empty:
                    print(f"✅ Extracted generic data ({len(generic_data)} entries) from {url}")
                    all_dataframes.append(generic_data)
                else:
                    print(f"⚠️ No data could be extracted from {url}")

            return all_dataframes

        except Exception as e:
            print(f"❌ Error processing {url}: {str(e)}")
            return []

    def process_rfps(self):
        """Process URLs from Excel file and extract filtered RFP data."""
        if self.excel_data is None:
            print("No data loaded. Please select a file first.")
            return

        driver = self.initialize_driver()
        all_data = []
        filtered_data = []

        try:
            total_urls = len(self.excel_data)

            for idx, row in self.excel_data.iterrows():
                url = row['URL']

                print(f"\n📋 Processing ({idx + 1}/{total_urls}): {url}")

                # Extract data from the URL
                dataframes = self.extract_web_data(url, driver)

                if not dataframes:
                    continue

                # Combine all dataframes from this URL
                combined_df = pd.concat(dataframes, ignore_index=True)

                # Add URL source information if not already present
                if 'Source URL' not in combined_df.columns:
                    combined_df['Source URL'] = url

                all_data.append(combined_df)

                # Apply filters to this batch
                filtered_batch = self.apply_filters(combined_df)
                if not filtered_batch.empty:
                    filtered_data.append(filtered_batch)
                    print(f"✅ {len(filtered_batch)} entries match your filters from {url}")
                else:
                    print(f"🔍 No entries matched filters from {url}")

                # Add a small delay to be nice to servers
                time.sleep(2)

            # Save results
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

            # Save all data (unfiltered)
            if all_data:
                all_df = pd.concat(all_data, ignore_index=True)
                unfiltered_file = os.path.join(self.output_directory, f"All_Data_{timestamp}.xlsx")
                all_df.to_excel(unfiltered_file, index=False)
                print(f"\n📊 All extracted data: {len(all_df)} total entries")
                print(f"💾 Unfiltered data saved to: {unfiltered_file}")

            # Save filtered data
            if filtered_data:
                final_filtered_df = pd.concat(filtered_data, ignore_index=True)
                filtered_file = os.path.join(self.output_directory, f"Filtered_RFPs_{timestamp}.xlsx")
                final_filtered_df.to_excel(filtered_file, index=False)

                print(f"\n🎉 SUCCESS! Found {len(final_filtered_df)} RFPs matching your criteria!")
                print(f"✅ Filtered results saved to: {filtered_file}")

                # Display summary
                if 'Source URL' in final_filtered_df.columns:
                    print("\n📊 Results by Source:")
                    for source in final_filtered_df['Source URL'].unique():
                        count = len(final_filtered_df[final_filtered_df['Source URL'] == source])
                        print(f"   • {source}: {count} RFPs")

                # Show sample of results
                if len(final_filtered_df) > 0:
                    print(f"\n📋 Sample of filtered RFPs:")
                    sample_cols = ['Title', 'Status', 'Due Date',
                                   'Agency'] if 'Title' in final_filtered_df.columns else final_filtered_df.columns[:4]
                    available_cols = [col for col in sample_cols if col in final_filtered_df.columns]
                    if available_cols:
                        print(final_filtered_df[available_cols].head(3).to_string(index=False))

            else:
                print(f"\n⚠️ No RFPs matched your filter criteria.")
                if self.rfp_type_keywords or self.status_filter_enabled:
                    print("Consider adjusting your filters and trying again.")

        finally:
            driver.quit()

    def run(self):
        """Main method to run the application"""
        print("===== RFP Extractor with Simple Filtering =====")
        print("This tool extracts RFP data with two simple filters:")
        print("1. RFP Type (what kind of opportunities you want)")
        print("2. Status Filter (exclude closed/expired RFPs)")
        print("\nExcel file must contain a 'URL' column.")

        if self.select_file():
            print(f"\n📋 Found {len(self.excel_data)} URLs to process.")

            # Show preview of URLs
            print("\nPreview of your URLs:")
            for idx, row in self.excel_data.head().iterrows():
                print(f"   {idx + 1}. {row['URL']}")

            if len(self.excel_data) > 5:
                print(f"   ... and {len(self.excel_data) - 5} more URLs")

            # Setup filters
            use_filters = input("\nDo you want to set up filters? (y/n): ").strip().lower()

            if use_filters == 'y':
                if not self.setup_filters():
                    print("Filter setup canceled.")
                    return
            else:
                print("No filters applied - all RFP data will be extracted.")

            confirm = input("\nDo you want to proceed with extraction? (y/n): ").strip().lower()

            if confirm == 'y':
                self.process_rfps()
            else:
                print("Extraction canceled.")

        print("\nThank you for using RFP Extractor!")


if __name__ == "__main__":
    extractor = RFPExtractor()
    extractor.run()