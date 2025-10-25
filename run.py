from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import time

def save_to_excel(data, output_file):
    # Create a new workbook and select active sheet
    wb = Workbook()
    ws = wb.active
    ws.title = "NAFDAC Greenbook"
    
    # Define headers
    headers = [
        "Product Name", "Active Ingredient", "Dosage Form",
        "Product Category", "NAFDAC Reg No", "Applicant", 
        "Manufacturer", "Approval Date"
    ]
    
    # Write headers with styling
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")
    border = Border(
        left=Side(style='thin'), 
        right=Side(style='thin'), 
        top=Side(style='thin'), 
        bottom=Side(style='thin')
    )
    
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = border
        
    # Write data
    for row_idx, row_data in enumerate(data, 2):
        for col_idx, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.border = border
            cell.alignment = Alignment(vertical="center")
    
    # Adjust column widths
    for col in range(1, len(headers) + 1):
        column_letter = get_column_letter(col)
        max_length = 0
        column = ws[column_letter]
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2) * 1.2
        ws.column_dimensions[column_letter].width = min(adjusted_width, 50)
    
    # Save workbook
    wb.save(output_file)

def scrape_greenbook(output_file="nafdac_greenbook.xlsx"):
    # Setup Chrome options
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")  # Larger window for better rendering
    
    # Initialize driver
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    
    try:
        # Load the page
        print("Loading page...")
        driver.get("https://greenbook.nafdac.gov.ng/")
        
        # Wait for table to load
        wait = WebDriverWait(driver, 20)
        print("Waiting for table to appear...")
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.dataTable")))
        time.sleep(2)  # Allow table to fully render
        
        # Initialize data collection
        data = []
        page = 1
        
        while True:
            print(f"Scraping page {page}...")
            
            # Get rows from current page
            rows = driver.find_elements(By.CSS_SELECTOR, "table.dataTable tbody tr")
            print(f"Found {len(rows)} rows on page {page}")
            
            # Extract data from rows
            page_data = []
            for row in rows:
                try:
                    cells = row.find_elements(By.TAG_NAME, "td")
                    row_data = [cell.text.strip() for cell in cells]
                    if row_data:
                        page_data.append(row_data)
                except:
                    continue
            
            # Add to main data list
            data.extend(page_data)
            print(f"Total records collected: {len(data)}")
            
            # Save progress every 5 pages
            if page % 5 == 0:
                print(f"Saving checkpoint at page {page}...")
                save_to_excel(data, output_file)
            
            # Look for next button
            try:
                next_button = driver.find_element(By.CSS_SELECTOR, "li.page-item.next:not(.disabled) a")
            except:
                print("No more pages found.")
                break
                
            # Click next button and wait for table update
            try:
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_button)
                time.sleep(0.5)
                next_button.click()
                wait.until(EC.staleness_of(rows[0]))  # Wait for first row to change
                time.sleep(1)  # Small pause for table to fully update
                page += 1
            except Exception as e:
                print(f"Error navigating to next page: {e}")
                break
        
        # Save final data to Excel
        print(f"Saving {len(data)} records to {output_file}")
        save_to_excel(data, output_file)
        
        print("Scraping completed successfully")
        
    except Exception as e:
        print(f"Error occurred: {str(e)}")
    finally:
        driver.quit()

if __name__ == "__main__":
    scrape_greenbook()
