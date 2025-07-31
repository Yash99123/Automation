import warnings

warnings.filterwarnings("ignore")

import time
import re
import requests
import pandas as pd
import os
from datetime import datetime
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager


class RFPExtractor:
    def __init__(self):
        self.output_directory = os.path.join(os.path.expanduser("~"), "Desktop", "RFP_Extractions")

        # Create output directory if it doesn't exist
        if not os.path.exists(self.output_directory):
            os.makedirs(self.output_directory)

        # Setup Selenium WebDriver
        self.driver = None
        self.wait = None

        # Keywords for filtering bids
        self.keywords = [
            "IT", "IT Services", "IT Staffing", "Research",
            "Engineering", "Software", "Security", "Information Technology", "Data",
            "software", "technology", "technical", "SAP", "Services"
        ]

    def initialize_driver(self):
        """Initialize the Selenium WebDriver with options."""
        options = Options()
        options.add_argument("--disable-gpu")
        options.add_argument("--window-size=1920x1080")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--start-maximized")
        options.headless = False  # Change to True to run headless

        self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        self.wait = WebDriverWait(self.driver, 30)
        return True

    def is_keyword_match(self, text):
        """Check if text contains any of the specified keywords."""
        if not text:
            return False
        return any(re.search(r'\b' + re.escape(k) + r'\b', text, re.IGNORECASE) for k in self.keywords)

    def create_empty_dataframe(self):
        """Create an empty dataframe with standard columns."""
        return pd.DataFrame(
            columns=["Source", "Title", "Company/Customer", "Expiration/Closing Date", "Status", "Link", "Extra Info"])

    # ====================== HELPER METHODS FOR FLORIDA EXTRACTION ======================

    def get_text_or_na(self, xpath):
        """Helper method for Florida extraction to get text or return N/A."""
        try:
            elem = self.wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
            return elem.text.strip()
        except TimeoutException:
            return "N/A"

    def get_b_tag_datetime(self, label):
        """Helper method for Florida extraction to get datetime from b tags."""
        try:
            b_elem = self.wait.until(EC.presence_of_element_located(
                (By.XPATH, f"//mfmp-bid-detail//b[contains(text(),'{label}')]")
            ))
            full_text = b_elem.get_attribute("innerText").strip()
            if ":" in full_text:
                return full_text.split(":", 1)[1].strip()
            return full_text
        except TimeoutException:
            return "N/A"

    def get_ad_title(self):
        """Helper method for Florida extraction to get ad title."""
        # Try multiple XPaths for the ad title
        xpaths = [
            "//mfmp-bid-detail//h1[contains(@class, 'mat-headline')]",
            "//mfmp-bid-detail//h1",
            "//h1[contains(@class, 'mat-headline')]",
            "//h1"
        ]
        for xp in xpaths:
            title = self.get_text_or_na(xp)
            if title != "N/A" and title.strip() != "":
                return title
        return "N/A"

    # ====================== FLORIDA MARKETPLACE EXTRACTOR ======================

    def extract_florida_marketplace(self):
        """Extract from MyFloridaMarketplace"""
        print("Extracting from MyFloridaMarketplace...")
        try:
            # 1. Login
            self.driver.get("https://vendor.myfloridamarketplace.com/login")
            self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[formcontrolname="username"]'))).send_keys("eitacies")
            self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[formcontrolname="password"]'))).send_keys("Eita78900$%")
            self.wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//span[contains(text(), 'Login')]]"))).click()

            # 2. Wait until logged in (toolbar presence)
            self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'mat-toolbar.mat-toolbar')))

            # 3. Click Advertisements menu
            self.wait.until(EC.element_to_be_clickable((By.XPATH, "//a[.//span[contains(text(), 'Advertisements')]]"))).click()

            # 4. Wait for Recommended Advertisements card
            recommended_ads_ul_xpath = "//mat-card[.//mat-card-title[contains(text(),'Recommended Advertisements')]]//ul"
            self.wait.until(EC.presence_of_element_located((By.XPATH, recommended_ads_ul_xpath)))

            ads_data = []

            # 5. Get all recommended ads links
            recommended_ads = self.wait.until(EC.presence_of_all_elements_located((By.XPATH, recommended_ads_ul_xpath + "/li/a")))
            print(f"Found {len(recommended_ads)} recommended ads to process.")

            for i in range(len(recommended_ads)):
                recommended_ads = self.wait.until(EC.presence_of_all_elements_located((By.XPATH, recommended_ads_ul_xpath + "/li/a")))
                ad = recommended_ads[i]

                ad_title = ad.get_attribute("title") or ad.text.strip()
                print(f"\nProcessing ad {i+1}/{len(recommended_ads)}: {ad_title}")

                # Check if ad title matches keywords before processing
                if not self.is_keyword_match(ad_title):
                    print(f"Skipping ad - no keyword match: {ad_title}")
                    continue

                self.driver.execute_script("arguments[0].scrollIntoView(true);", ad)
                time.sleep(0.5)
                ad.click()

                self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "mfmp-bid-detail")))
                time.sleep(1)

                ad_href = self.driver.current_url

                ad_title_detail = self.get_ad_title()

                ad_no_text = self.get_text_or_na("//mfmp-bid-detail//div[contains(text(),'Advertisement Number')]")
                ad_no = ad_no_text.split(":", 1)[1].strip() if ":" in ad_no_text else ad_no_text

                start_date = self.get_b_tag_datetime("Start Date/Time")
                end_date = self.get_b_tag_datetime("End Date/Time")

                print(f"Ad Title: {ad_title_detail}")
                print(f"Ad No: {ad_no}")
                print(f"Start Date: {start_date}")
                print(f"End Date: {end_date}")
                print(f"Ad Link: {ad_href}")

                ads_data.append([
                    "MyFloridaMarketplace",
                    ad_title_detail,
                    "Florida State",
                    end_date,
                    "Active",
                    ad_href,
                    f"Ad No: {ad_no}; Start Date: {start_date}"
                ])

                self.driver.back()
                self.wait.until(EC.presence_of_element_located((By.XPATH, recommended_ads_ul_xpath)))
                time.sleep(1)

            df = pd.DataFrame(ads_data, columns=["Source", "Title", "Company/Customer", "Expiration/Closing Date", "Status", "Link", "Extra Info"])
            return df

        except Exception as e:
            print(f"MyFloridaMarketplace extraction error: {e}")
            return self.create_empty_dataframe()

    # ====================== PREDEFINED SITE EXTRACTORS ======================

    def extract_lausd(self):
        """Extract from LAUSD"""
        print("Extracting from LAUSD...")
        url = "https://psd.lausd.net/vendors/RFPList.aspx?RFPStatus=Current"
        grid_id = "CurrentList"
        self.driver.get(url)
        try:
            table = self.wait.until(EC.visibility_of_element_located((By.ID, grid_id)))
            rows = table.find_elements(By.TAG_NAME, "tr")
            filtered_data = []
            for row in rows:
                cells = row.find_elements(By.TAG_NAME, "td")
                if cells:
                    row_text = ' '.join(cell.text for cell in cells)
                    if self.is_keyword_match(row_text):
                        filtered_data.append([cell.text for cell in cells])

            if filtered_data:
                col_count = len(filtered_data[0])
                col_names = ["Title", "Company/Customer", "Expiration/Closing Date", "Status", "Link"]
                if col_count < len(col_names):
                    col_names = col_names[:col_count]
                elif col_count > len(col_names):
                    col_names += [f"Column_{i}" for i in range(len(col_names) + 1, col_count + 1)]

                df = pd.DataFrame(filtered_data, columns=col_names)
                df.insert(0, "Source", "LAUSD")
                df["Extra Info"] = ""

                expected_cols = ["Source", "Title", "Company/Customer", "Expiration/Closing Date", "Status", "Link",
                                 "Extra Info"]
                for col in expected_cols:
                    if col not in df.columns:
                        df[col] = ""

                return df[expected_cols]
            else:
                print("No matching LAUSD data.")
                return self.create_empty_dataframe()
        except Exception as e:
            print(f"LAUSD extraction error: {e}")
            return self.create_empty_dataframe()

    def extract_stlmsd(self):
        """Extract from STLMSD Planroom"""
        print("Extracting from STLMSD Planroom...")
        base_url = "https://www.stlmsdplanroom.com/projects/public?page="
        titles, dates, companies, links = [], [], [], []

        for page in range(1, 3):
            try:
                resp = requests.get(base_url + str(page))
                if resp.status_code == 200:
                    soup = BeautifulSoup(resp.content, 'html.parser')
                    projects = soup.find_all('a', class_='row')
                    for proj in projects:
                        name_tag = proj.find('div', class_='name')
                        company_tag = proj.find('div', class_='company')
                        if name_tag and company_tag:
                            full_name = name_tag.get_text(strip=True)
                            company = company_tag.get_text(strip=True)
                            link = proj['href']
                            if self.is_keyword_match(full_name):
                                if "Posting expires" in full_name:
                                    title_part = full_name.split("Posting expires")[0].strip(" -")
                                    exp_date = full_name.split("Posting expires")[-1].strip()
                                else:
                                    title_part = full_name
                                    exp_date = ""
                                titles.append(title_part)
                                companies.append(company)
                                dates.append(exp_date)
                                links.append(link)
                else:
                    print(f"Failed to load STLMSD page {page}")
            except Exception as e:
                print(f"STLMSD page {page} error: {e}")

        df = pd.DataFrame({
            "Source": "STLMSD",
            "Title": titles,
            "Company/Customer": companies,
            "Expiration/Closing Date": dates,
            "Status": "",
            "Link": links,
            "Extra Info": ""
        })
        return df

    def extract_ksu(self):
        """Extract from Kansas State University"""
        print("Extracting from Kansas State University...")
        url = "https://bidportal.ksu.edu/Module/Tenders/en/Login/Index/68558aea-8748-49c0-95c7-34c4bf7cb29b"
        self.driver.get(url)
        try:
            self.wait.until(EC.presence_of_element_located((By.ID, "Username"))).send_keys("bdm@eitacies.com")
            self.driver.find_element(By.ID, "Password").send_keys("EIta7890$%")
            self.driver.find_element(By.XPATH, "//button[@type='submit']").click()
            time.sleep(5)
            self.wait.until(EC.presence_of_element_located((By.ID, "ext-gen77")))
            soup = BeautifulSoup(self.driver.page_source, 'html.parser')
            rows = soup.select('.x-grid3-row, .x-grid3-row-alt')
            data = []
            for row in rows:
                customer = row.select_one('.x-grid3-td-Customer')
                title = row.select_one('.x-grid3-td-Title')
                status = row.select_one('.x-grid3-td-BidStatus')
                closing = row.select_one('.x-grid3-td-ClosingDateStr')
                submission = row.select_one('.x-grid3-td-SubmissionStatusDetails')

                title_text = title.text.strip() if title else 'N/A'
                if self.is_keyword_match(title_text):
                    data.append([
                        "Kansas State University",
                        title_text,
                        customer.text.strip() if customer else 'N/A',
                        closing.text.strip() if closing else '',
                        status.text.strip() if status else '',
                        title.find('a')['href'] if title and title.find('a') else '',
                        submission.text.strip() if submission else ''
                    ])

            df = pd.DataFrame(data, columns=["Source", "Title", "Company/Customer", "Expiration/Closing Date", "Status",
                                             "Link", "Extra Info"])
            return df
        except Exception as e:
            print(f"KSU extraction error: {e}")
            return self.create_empty_dataframe()



    def extract_ionwave(self):
        """Extract from IonWave"""
        print("Extracting from IonWave...")
        url = "https://supplier.ionwave.net/VendorResponse/ResponseList.aspx?status=AVAILABLE&cid=Tt1adTkP0qI_"
        self.driver.get(url)
        try:
            self.wait.until(EC.presence_of_element_located((By.ID, "divLoginRegistration")))
            self.wait.until(EC.element_to_be_clickable((By.NAME, "txtUserName"))).send_keys("eita123")
            self.wait.until(EC.element_to_be_clickable((By.NAME, "txtPassword"))).send_keys("Eita1627$%3849")
            self.wait.until(EC.element_to_be_clickable((By.NAME, "btnLogin"))).click()
            self.wait.until(EC.presence_of_element_located((By.CLASS_NAME, "rgMasterTable")))
            time.sleep(5)
            soup = BeautifulSoup(self.driver.page_source, 'html.parser')
            rows = soup.find_all('tr', class_=['rgRow', 'rgAltRow'])
            data = []
            for row in rows:
                cells = row.find_all('td')
                if len(cells) >= 9:
                    title = cells[3].text.strip()
                    if self.is_keyword_match(title):
                        data.append([
                            "IonWave",
                            title,
                            cells[1].text.strip(),
                            cells[5].text.strip(),
                            cells[7].text.strip(),
                            cells[0].find('a')['href'] if cells[0].find('a') else '',
                            f"RFP ID: {cells[2].text.strip()}; Start Date: {cells[4].text.strip()}; Duration: {cells[6].text.strip()}; View Status: {cells[8].text.strip()}"
                        ])
            df = pd.DataFrame(data, columns=["Source", "Title", "Company/Customer", "Expiration/Closing Date", "Status",
                                             "Link", "Extra Info"])
            return df
        except Exception as e:
            print(f"IonWave extraction error: {e}")
            return self.create_empty_dataframe()



    def extract_vendorline(self):
        """Extract from Vendorline"""
        print("Extracting from Vendorline...")
        url = "https://vendorline.planetbids.com/app/sub-bid-ads#"
        try:
            self.driver.get(url)
            time.sleep(6)
            self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(3)
            soup = BeautifulSoup(self.driver.page_source, 'html.parser')
            ads_wrapper = soup.select_one('.sub-bid-ads-wrapper.general-ads')
            if not ads_wrapper:
                print("No ad wrapper found.")
                return self.create_empty_dataframe()

            ad_buttons = ads_wrapper.find_all("button")
            data = []
            for ad in ad_buttons:
                try:
                    title_text = ad.select_one('.bid-ad-title').text.strip()
                    prime_text = ad.select_one('.bid-ad-display-value').text.strip()
                    if self.is_keyword_match(title_text) or self.is_keyword_match(prime_text):
                        data.append([
                            "Vendorline",
                            title_text,
                            prime_text,
                            ad.select_one('.bid-ad-response-due').text.strip() if ad.select_one(
                                '.bid-ad-response-due') else '',
                            ad.select_one('.bid-ad-stage').text.strip() if ad.select_one('.bid-ad-stage') else '',
                            '',
                            f"State (Full): {ad.select_one('.gen-bid-ad-state-full').text.strip() if ad.select_one('.gen-bid-ad-state-full') else ''}, State (Abbr): {ad.select_one('.gen-bid-ad-state-abbr').text.strip() if ad.select_one('.gen-bid-ad-state-abbr') else ''}"
                        ])
                except Exception as inner:
                    print(f"Error parsing vendorline ad: {inner}")
            df = pd.DataFrame(data, columns=["Source", "Title", "Company/Customer", "Expiration/Closing Date", "Status",
                                             "Link", "Extra Info"])
            return df
        except Exception as e:
            print(f"Vendorline error: {e}")
            return self.create_empty_dataframe()

    def extract_houston_isd(self):
        """Extract from Houston ISD"""
        print("Extracting from Houston ISD...")
        url = "https://houstonisd.ionwave.net//VendorResponse/ResponseList.aspx?status=AVAILABLE&cid=3Qnbg1KG-Dc_"
        self.driver.get(url)
        try:
            self.wait.until(EC.presence_of_element_located((By.ID, "divLoginRegistration")))
            self.wait.until(EC.element_to_be_clickable((By.NAME, "txtUserName"))).send_keys("eitaciesinc")
            self.wait.until(EC.element_to_be_clickable((By.NAME, "txtPassword"))).send_keys("Eita125689$%")
            self.wait.until(EC.element_to_be_clickable((By.NAME, "btnLogin"))).click()
            self.wait.until(EC.presence_of_element_located((By.CLASS_NAME, "rgMasterTable")))
            time.sleep(5)
            soup = BeautifulSoup(self.driver.page_source, 'html.parser')
            rows = soup.find_all('tr', class_=['rgRow', 'rgAltRow'])
            data = []
            for row in rows:
                cells = row.find_all('td')
                if len(cells) >= 9:
                    title = cells[3].text.strip()
                    company = cells[1].text.strip()
                    rfp_id = cells[2].text.strip()
                    if self.is_keyword_match(title) or self.is_keyword_match(company) or self.is_keyword_match(rfp_id):
                        data.append([
                            "Houston ISD",
                            title,
                            company,
                            cells[5].text.strip(),
                            cells[7].text.strip(),
                            cells[0].find('a')['href'] if cells[0].find('a') else '',
                            f"RFP ID: {rfp_id}; Start Date: {cells[4].text.strip()}; Duration: {cells[6].text.strip()}; View Status: {cells[8].text.strip()}"
                        ])
            df = pd.DataFrame(data, columns=["Source", "Title", "Company/Customer", "Expiration/Closing Date", "Status",
                                             "Link", "Extra Info"])
            return df
        except Exception as e:
            print(f"Houston ISD extraction error: {e}")
            return self.create_empty_dataframe()

    # ====================== EXCEL SHEET URL EXTRACTORS ======================

    def scroll_page(self, times=10):
        """Scroll down the page to load dynamic content."""
        try:
            body = self.driver.find_element(By.TAG_NAME, "body")
            for _ in range(times):
                body.send_keys(Keys.PAGE_DOWN)
                time.sleep(1)
            time.sleep(5)
        except:
            pass

    def extract_generic_site_data(self, url, site_name, keywords):
        """Generic extraction method for sites from Excel list."""
        print(f"Extracting from {site_name}...")
        try:
            self.driver.get(url)
            time.sleep(5)

            # Scroll to load dynamic content
            self.scroll_page()

            soup = BeautifulSoup(self.driver.page_source, 'html.parser')
            data = []

            # Look for common RFP/bid listing patterns
            # Pattern 1: Table-based listings
            tables = soup.find_all("table")
            for table in tables:
                rows = table.find_all("tr")
                for row in rows:
                    cells = row.find_all(["td", "th"])
                    if len(cells) > 1:
                        row_text = ' '.join(cell.get_text(strip=True) for cell in cells)
                        if self.is_keyword_match(row_text):
                            # Try to extract link from the row
                            link_elem = row.find("a")
                            link = link_elem["href"] if link_elem and "href" in link_elem.attrs else ""
                            if link and not link.startswith(("http://", "https://")):
                                base_url = "/".join(url.split("/")[:3])
                                link = f"{base_url}{link}" if link.startswith("/") else f"{base_url}/{link}"

                            # Extract title (usually first or second cell)
                            title = cells[0].get_text(strip=True) if cells else row_text[:100]

                            data.append([
                                site_name,
                                title,
                                "N/A",  # Company/Customer
                                "N/A",  # Expiration Date
                                "Active",  # Status
                                link,
                                row_text[:200] + "..." if len(row_text) > 200 else row_text
                            ])

            # Pattern 2: Div-based listings
            if not data:
                listing_divs = soup.find_all("div", class_=lambda c: c and any(
                    term in c.lower() for term in
                    ["opportunity", "bid", "rfp", "procurement", "tender", "listing", "item"]))

                for div in listing_divs:
                    div_text = div.get_text(strip=True)
                    if self.is_keyword_match(div_text):
                        # Try to find title
                        title_elem = div.find(["h1", "h2", "h3", "h4", "h5", "a", "strong"])
                        title = title_elem.get_text(strip=True) if title_elem else div_text[:100]

                        # Try to find link
                        link_elem = div.find("a")
                        link = link_elem["href"] if link_elem and "href" in link_elem.attrs else ""
                        if link and not link.startswith(("http://", "https://")):
                            base_url = "/".join(url.split("/")[:3])
                            link = f"{base_url}{link}" if link.startswith("/") else f"{base_url}/{link}"

                        # Try to find date
                        date_text = "N/A"
                        date_patterns = ["due", "deadline", "close", "closing", "expires", "expiration"]
                        for pattern in date_patterns:
                            if pattern in div_text.lower():
                                # Extract text around the date pattern
                                import re
                                date_match = re.search(rf'{pattern}[:\s]*([^,\n]+)', div_text, re.IGNORECASE)
                                if date_match:
                                    date_text = date_match.group(1).strip()
                                    break

                        data.append([
                            site_name,
                            title,
                            "N/A",
                            date_text,
                            "Active",
                            link,
                            div_text[:200] + "..." if len(div_text) > 200 else div_text
                        ])

            # Pattern 3: List-based items
            if not data:
                list_items = soup.find_all("li")
                for item in list_items:
                    item_text = item.get_text(strip=True)
                    if len(item_text) > 20 and self.is_keyword_match(item_text):  # Filter out short navigation items
                        link_elem = item.find("a")
                        link = link_elem["href"] if link_elem and "href" in link_elem.attrs else ""
                        if link and not link.startswith(("http://", "https://")):
                            base_url = "/".join(url.split("/")[:3])
                            link = f"{base_url}{link}" if link.startswith("/") else f"{base_url}/{link}"

                        title = item_text[:100] + "..." if len(item_text) > 100 else item_text

                        data.append([
                            site_name,
                            title,
                            "N/A",
                            "N/A",
                            "Active",
                            link,
                            item_text[:200] + "..." if len(item_text) > 200 else item_text
                        ])

            # If still no data, try to find any text matching keywords
            if not data:
                all_text = soup.get_text()
                if self.is_keyword_match(all_text):
                    data.append([
                        site_name,
                        f"Keyword match found on {site_name}",
                        "N/A",
                        "N/A",
                        "Requires Manual Review",
                        url,
                        "Keywords found on page but specific opportunities need manual review"
                    ])

            df = pd.DataFrame(data, columns=["Source", "Title", "Company/Customer", "Expiration/Closing Date", "Status",
                                             "Link", "Extra Info"])

            if not df.empty:
                print(f"✅ Found {len(df)} matching entries from {site_name}")
            else:
                print(f"⚠️ No matching data found from {site_name}")

            return df

        except Exception as e:
            print(f"❌ Error extracting from {site_name}: {e}")
            return self.create_empty_dataframe()

    def extract_planetbids_portal_40988(self):
        """Extract from PlanetBids Portal 40988"""
        return self.extract_generic_site_data(
            "https://vendors.planetbids.com/portal/40988/bo/bo-search",
            "PlanetBids Portal 40988",
            ["IT", "Cyber security", "Software", "AI", "ML"]
        )

    def extract_cps_bonfire_portal(self):
        """Extract from CPS Bonfire Portal"""
        return self.extract_generic_site_data(
            "https://cps.bonfirehub.com/portal/?tab=openOpportunities",
            "CPS Bonfire Portal",
            ["IT", "Cyber security", "Software", "AI", "ML"]
        )

    def extract_cps_bonfire_opportunity(self):
        """Extract from CPS Bonfire Specific Opportunity"""
        return self.extract_generic_site_data(
            "https://cps.bonfirehub.com/opportunities/87850",
            "CPS Bonfire Opportunity",
            ["IT", "Cyber security", "Software", "AI", "ML"]
        )

    def extract_txdot_bonfire(self):
        """Extract from TxDOT Bonfire"""
        return self.extract_generic_site_data(
            "https://txdot.bonfirehub.com/portal/?tab=openOpportunities",
            "TxDOT Bonfire",
            ["IT", "Cyber security", "Software", "AI", "ML"]
        )

    def extract_sunnyvale_gov(self):
        """Extract from Sunnyvale Government"""
        return self.extract_generic_site_data(
            "https://www.sunnyvale.ca.gov/business-and-development/bid-on-city-projects",
            "Sunnyvale Government",
            ["IT", "Cyber security", "Software", "AI", "ML"]
        )

    def extract_metrohealth_infor(self):
        """Extract from MetroHealth Infor"""
        return self.extract_generic_site_data(
            "https://metrohealthprod-lm01.cloud.infor.com:1442/lmscm/SourcingSupplier/list/SourcingEvent.OpenForBid?csk.CHP=LMPROC&csk.SupplierGroup=MHS&menu=EventManagement.BrowseOpenEvents",
            "MetroHealth Infor",
            ["IT", "Cyber security", "Software", "AI", "ML"]
        )

    def extract_cuyahoga_county_infor(self):
        """Extract from Cuyahoga County Infor"""
        return self.extract_generic_site_data(
            "https://ccprod-lm01.cloud.infor.com:1442/lmscm/SourcingSupplier/list/SourcingEvent.OpenForBid?csk.CHP=LMPROC&csk.SupplierGroup=CUYA&menu=EventManagement.BrowseOpenEvents",
            "Cuyahoga County Infor",
            ["IT", "Cyber security", "Software", "AI", "ML"]
        )

    def extract_cal_eprocure(self):
        """Extract from California eProcure"""
        return self.extract_generic_site_data(
            "https://caleprocure.ca.gov/pages/Events-BS3/event-search.aspx",
            "California eProcure",
            ["IT", "Cyber security", "Software", "AI", "ML"]
        )

    def extract_roswell_bonfire(self):
        """Extract from Roswell Bonfire"""
        return self.extract_generic_site_data(
            "https://roswellgov.bonfirehub.com/portal/?tab=openOpportunities",
            "Roswell Bonfire",
            ["IT", "Cyber security", "Software", "AI", "ML"]
        )

    # ====================== MAIN PROCESSING METHODS ======================

    def extract_all_sites(self):
        """Extract data from all predefined sites."""
        print("=" * 80)
        print("                    RFP DATA EXTRACTOR")
        print("=" * 80)
        print("Extracting from predefined sites...")
        print("-" * 80)

        extractors = [
            # Original predefined sites
            self.extract_lausd,
            self.extract_stlmsd,
            self.extract_ksu,

            self.extract_ionwave,

            self.extract_vendorline,
            self.extract_houston_isd,

            # NEW: Florida Marketplace
            self.extract_florida_marketplace,

            # Excel sheet URLs
            self.extract_planetbids_portal_40988,
            self.extract_cps_bonfire_portal,
            self.extract_cps_bonfire_opportunity,
            self.extract_txdot_bonfire,
            self.extract_sunnyvale_gov,
            self.extract_metrohealth_infor,
            self.extract_cuyahoga_county_infor,
            self.extract_cal_eprocure,
            self.extract_roswell_bonfire,
        ]

        all_results = []
        for extractor in extractors:
            try:
                print(f"\n🔄 Running {extractor.__name__.replace('extract_', '').upper()}...")
                result = extractor()
                if result is not None and not result.empty:
                    all_results.append(result)
                    print(f"✅ Found {len(result)} matching entries")
                else:
                    print("⚠️ No matching data found")
                time.sleep(2)  # Be nice to servers
            except Exception as e:
                print(f"❌ Error in {extractor.__name__}: {e}")
                continue

        return all_results

    def run(self):
        """Main method to run the extraction."""
        try:
            print("🚀 Initializing web driver...")
            if not self.initialize_driver():
                print("❌ Failed to initialize web driver.")
                return

            # Extract from all sites
            results = self.extract_all_sites()

            if results:
                # Combine all results
                combined_df = pd.concat(results, ignore_index=True)
                combined_df.drop_duplicates(inplace=True)

                # Save results
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_file = os.path.join(self.output_directory, f"rfp_data_{timestamp}.xlsx")
                combined_df.to_excel(output_file, index=False)

                print("\n" + "=" * 80)
                print("🎉 EXTRACTION COMPLETED SUCCESSFULLY!")
                print("=" * 80)
                print(f"📊 Total RFPs extracted: {len(combined_df)}")
                print(f"💾 Results saved to: {output_file}")

                # Display summary by source
                print("\n📋 Summary by Source:")
                source_counts = combined_df['Source'].value_counts()
                for source, count in source_counts.items():
                    print(f"   • {source}: {count} RFPs")

                print("\n🔍 Keywords used for filtering:")
                print("   " + ", ".join(self.keywords))

                return combined_df
            else:
                print("\n⚠️ No RFPs found across all sources.")
                return None

        except KeyboardInterrupt:
            print("\n\n❌ Process interrupted by user.")
        except Exception as e:
            print(f"\n❌ An unexpected error occurred: {e}")
            import traceback
            traceback.print_exc()
        finally:
            if self.driver:
                print("🔒 Closing web driver...")
                self.driver.quit()


def main():
    """Main function to run the extractor."""
    extractor = RFPExtractor()
    extractor.run()


if __name__ == "__main__":
    main()