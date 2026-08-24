import requests
import time
import csv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime

class AmazonVacuumCollector:
    def __init__(self, api_key, csv_file_path):
        self.api_key = api_key
        self.csv_file_path = csv_file_path
        self.asins = []
        self.all_products = []
        self.base_url = "https://api.rainforestapi.com/request"

        # Initialize Selenium for seller info extraction
        self.driver = None
        self._init_selenium()
        self._load_asins_from_csv()

    def _load_asins_from_csv(self):
        """Load ASINs from CSV file"""
        try:
            with open(self.csv_file_path, 'r', encoding='utf-8') as file:
                csv_reader = csv.reader(file)
                for row in csv_reader:
                    # Skip headers and empty rows
                    if row and row[0].strip() and not row[0].startswith('Brand') and not row[0].startswith('Row') and row[0].strip() != 'Grand Total':
                        asin = row[0].strip()
                        # Check if it looks like a valid ASIN (starts with B and is alphanumeric)
                        if asin.startswith('B') and len(asin) >= 10:
                            self.asins.append(asin)

            print(f"✓ Loaded {len(self.asins)} ASINs from CSV file")
            return self.asins
        except FileNotFoundError:
            print(f"❌ CSV file not found: {self.csv_file_path}")
            return []
        except Exception as e:
            print(f"⚠️ Error loading CSV: {str(e)}")
            return []

    def _init_selenium(self):
        """Initialize Selenium WebDriver for scraping seller info"""
        try:
            options = webdriver.ChromeOptions()
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--disable-blink-features=AutomationControlled')
            options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36')

            self.driver = webdriver.Chrome(options=options)
            print("✓ Selenium WebDriver initialized")
        except Exception as e:
            print(f"⚠️ Warning: Could not initialize Selenium: {str(e)}")
            print("  Seller info will be set to 'Not Available'")
            self.driver = None

    def extract_seller_info_from_page(self, asin):
        """
        Extract Fulfilled by, Sold by, and Shipper/Seller information from Amazon.ae product page
        Returns: (fulfilled_by, sold_by, shipper_seller)
        """
        if not self.driver:
            return "Not Available", "Not Available", "Not Available"

        try:
            url = f"https://www.amazon.ae/dp/{asin}"
            self.driver.get(url)

            # Wait for page to load
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            time.sleep(1)

            # Scroll to load dynamic content
            self.driver.execute_script("window.scrollBy(0, 2000);")
            time.sleep(0.5)

            # Get page text
            page_text = self.driver.find_element(By.TAG_NAME, "body").text
            lines = page_text.split('\n')

            fulfilled_by = "Not Available"
            sold_by = "Not Available"
            shipper_seller = "Not Available"

            # Parse seller information from page text
            for i, line in enumerate(lines):
                line_stripped = line.strip()
                next_line = lines[i + 1].strip() if i + 1 < len(lines) else ''

                # Look for "Fulfilled by"
                if 'Fulfilled by' in line_stripped and fulfilled_by == "Not Available":
                    if ':' in line_stripped:
                        fulfilled_by = line_stripped.split(':', 1)[1].strip()[:100]
                    elif len(line_stripped) > len('Fulfilled by'):
                        fulfilled_by = line_stripped.replace('Fulfilled by', '').strip()[:100]
                    else:
                        fulfilled_by = next_line[:100] if next_line else "Not Available"

                # Look for "Sold by"
                if 'Sold by' in line_stripped and sold_by == "Not Available":
                    if ':' in line_stripped:
                        sold_by = line_stripped.split(':', 1)[1].strip()[:100]
                    elif len(line_stripped) > len('Sold by'):
                        sold_by = line_stripped.replace('Sold by', '').strip()[:100]
                    else:
                        sold_by = next_line[:100] if next_line else "Not Available"

                # Look for "Shipper / Seller" (with space variations)
                if 'Shipper' in line_stripped and '/' in line_stripped and 'Seller' in line_stripped:
                    if shipper_seller == "Not Available":
                        if ':' in line_stripped:
                            shipper_seller = line_stripped.split(':', 1)[1].strip()[:100]
                        elif len(line_stripped) > len('Shipper / Seller'):
                            shipper_seller = line_stripped.replace('Shipper / Seller', '').replace('Shipper/Seller', '').strip()[:100]
                        else:
                            shipper_seller = next_line[:100] if next_line else "Not Available"

            return fulfilled_by, sold_by, shipper_seller

        except Exception as e:
            print(f"  ⚠️ Error extracting seller info for {asin}: {str(e)}")
            return "Not Available", "Not Available", "Not Available"

    def fetch_product_details(self, asin):
        """Fetch product details from Rainforest API"""
        try:
            params = {
                "api_key": self.api_key,
                "amazon_domain": "amazon.ae",
                "type": "product",
                "asin": asin
            }

            response = requests.get(self.base_url, params=params, timeout=30)
            response.raise_for_status()
            data = response.json()

            if not data.get("request_info", {}).get("success"):
                print(f"  ⚠️ API Error for {asin}: {data.get('request_info', {}).get('error_message', 'Unknown error')}")
                return None

            product = data.get("product", {})
            if not product:
                print(f"  ⚠️ No product data for {asin}")
                return None

            # Extract price info (Rainforest API returns the live price under
            # buybox_winner.price, not a top-level "prices" list)
            price_info = product.get("buybox_winner", {}).get("price") or product.get("price") or {}
            price = price_info.get("raw", "N/A")

            # Extract brand (top-level field, falling back to the specifications table)
            brand = product.get("brand")
            if not brand:
                for spec in product.get("specifications", []):
                    if spec.get("name") in ("Brand", "Brand Name"):
                        brand = spec.get("value")
                        break
            brand = brand or "N/A"

            # Extract seller info from page using Selenium
            print(f"    Extracting seller info for {asin}...", end=" ")
            fulfilled_by, sold_by, shipper_seller = self.extract_seller_info_from_page(asin)
            print(f"✓")

            product_info = {
                "Brand": brand,
                "ASIN": asin,
                "Title": product.get("title", "N/A"),
                "Price": price,
                "Currency": price_info.get("currency", "AED"),
                "Fulfilled by": fulfilled_by,
                "Sold by": sold_by,
                "Shipper/Seller": shipper_seller,
                "Specs": product.get("specifications_flat") or product.get("description", ""),
                "Rating": product.get("rating", "N/A"),
                "Reviews": product.get("ratings_total", 0),
                "URL": f"https://www.amazon.ae/dp/{asin}"
            }

            return product_info

        except requests.exceptions.RequestException as e:
            print(f"  ⚠️ Network error for {asin}: {str(e)}")
            return None
        except Exception as e:
            print(f"  ⚠️ Error fetching {asin}: {str(e)}")
            return None

    def collect_all_products(self):
        """Collect products for all ASINs from CSV"""
        print("\n" + "="*70)
        print("🚀 AMAZON.AE VACUUM CLEANER DATA COLLECTION")
        print("   (Using filtered ASINs from CSV + Enhanced seller info extraction)")
        print("="*70)
        print(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"API Key: {self.api_key[:10]}...")
        print(f"Total ASINs to process: {len(self.asins)}\n")

        for idx, asin in enumerate(self.asins, 1):
            print(f"[{idx}/{len(self.asins)}] Processing {asin}...", end=" ")
            product = self.fetch_product_details(asin)

            if product:
                self.all_products.append(product)
                print(f"✓ {product['Title'][:40]}...")
            else:
                print("✗ Failed")

            # Respectful delay to avoid rate limiting
            time.sleep(1)

        print(f"\n✅ Total products collected: {len(self.all_products)}/{len(self.asins)}")
        return self.all_products

    def create_excel_file(self, output_filename="Amazon_Vacuum_Cleaners_Filtered.xlsx"):
        """Create Excel workbook with product data"""
        print(f"\n📝 Creating Excel file: {output_filename}")

        wb = Workbook()

        # Remove default sheet
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]

        # Create Summary Sheet
        summary_sheet = wb.create_sheet("Summary", 0)
        summary_sheet['A1'] = "Data Collection Summary"
        summary_sheet['A1'].font = Font(bold=True, size=14, color="FFFFFF")
        summary_sheet['A1'].fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        summary_sheet.merge_cells('A1:D1')

        summary_sheet['A3'] = "Metric"
        summary_sheet['B3'] = "Value"

        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")

        for cell in ['A3', 'B3']:
            summary_sheet[cell].fill = header_fill
            summary_sheet[cell].font = header_font

        summary_data = [
            ["Total Products Collected", len(self.all_products)],
            ["Total ASINs in CSV", len(self.asins)],
            ["Collection Date", datetime.now().strftime('%Y-%m-%d')],
            ["Collection Time", datetime.now().strftime('%H:%M:%S')],
            ["Avg Rating", f"{sum([float(p['Rating']) if isinstance(p['Rating'], (int, float)) else 0 for p in self.all_products]) / len(self.all_products) if self.all_products else 0:.1f}"]
        ]

        for i, (metric, value) in enumerate(summary_data, start=4):
            summary_sheet[f'A{i}'] = metric
            summary_sheet[f'B{i}'] = value

        summary_sheet.column_dimensions['A'].width = 25
        summary_sheet.column_dimensions['B'].width = 25

        # Create All Products Sheet
        products_sheet = wb.create_sheet("All Products", 1)

        headers = ["Brand", "ASIN", "Title", "Price", "Currency", "Fulfilled by", "Sold by", "Shipper/Seller", "Specs", "Rating", "Reviews", "URL"]
        products_sheet.append(headers)

        # Format header row
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF", size=12)

        for cell in products_sheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # Add product data
        for product in self.all_products:
            products_sheet.append([
                product["Brand"],
                product["ASIN"],
                product["Title"],
                product["Price"],
                product["Currency"],
                product["Fulfilled by"],
                product["Sold by"],
                product["Shipper/Seller"],
                product["Specs"],
                product["Rating"],
                product["Reviews"],
                product["URL"]
            ])

        # Set column widths
        products_sheet.column_dimensions['A'].width = 15
        products_sheet.column_dimensions['B'].width = 12
        products_sheet.column_dimensions['C'].width = 40
        products_sheet.column_dimensions['D'].width = 12
        products_sheet.column_dimensions['E'].width = 10
        products_sheet.column_dimensions['F'].width = 20
        products_sheet.column_dimensions['G'].width = 20
        products_sheet.column_dimensions['H'].width = 20
        products_sheet.column_dimensions['I'].width = 50
        products_sheet.column_dimensions['J'].width = 10
        products_sheet.column_dimensions['K'].width = 10
        products_sheet.column_dimensions['L'].width = 45

        # Freeze header row
        products_sheet.freeze_panes = "A2"

        # Save workbook
        wb.save(output_filename)
        print(f"✅ Excel file created successfully: {output_filename}")
        return output_filename

    def close(self):
        """Clean up Selenium driver"""
        if self.driver:
            self.driver.quit()
            print("\n✓ Browser closed")

def main():
    # Configuration
    API_KEY = "AECB78A27B374B34A2F037C3A6E1AA3B"
    CSV_FILE_PATH = "ASIN.csv"  # Path to your CSV file with filtered ASINs

    collector = None
    try:
        collector = AmazonVacuumCollector(API_KEY, CSV_FILE_PATH)

        if not collector.asins:
            print("\n❌ No ASINs loaded from CSV file. Exiting.")
            return

        collector.collect_all_products()

        if collector.all_products:
            output_file = collector.create_excel_file()
            print(f"\n🎉 SUCCESS! File ready: {output_file}")
            print("="*70)
            print("Features:")
            print("  ✅ Data from filtered ASINs in your CSV")
            print("  ✅ Enhanced seller info (Fulfilled by, Sold by, Shipper/Seller)")
            print("  ✅ 3-column seller info extraction")
            print("  ✅ Summary sheet with collection stats")
            print("="*70)
        else:
            print("\n❌ No products were collected. Please check your CSV file and API key.")

    except Exception as e:
        print(f"\n❌ Error: {str(e)}")
        print("Please ensure you have:")
        print("  - Internet access")
        print("  - Valid Rainforest API key")
        print("  - CSV file with ASINs in the same directory")
        print("  - Google Chrome installed")
        print("  - ChromeDriver installed or in PATH")

    finally:
        if collector:
            collector.close()

if __name__ == "__main__":
    main()