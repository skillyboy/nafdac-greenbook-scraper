from selenium import webdriver
import argparse
import os
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, UnexpectedAlertPresentException, NoAlertPresentException, StaleElementReferenceException
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from selenium.webdriver.common.alert import Alert
import time

def save_to_excel(data, output_file):
    import os
    from datetime import datetime
    
    # Try to save to the main file first
    def try_save(file_path, workbook):
        try:
            workbook.save(file_path)
            return True
        except PermissionError:
            return False
            
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
    
    # Try to save the workbook with retry mechanism
    if not try_save(output_file, wb):
        # If main file is locked, create a backup file
        backup_file = f"nafdac_greenbook_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        print(f"Main file is locked. Saving to backup file: {backup_file}")
        if try_save(backup_file, wb):
            print("Successfully saved to backup file")
        else:
            # If both fails, try saving to temp directory
            temp_file = os.path.join(os.environ.get('TEMP', ''), backup_file)
            if try_save(temp_file, wb):
                print(f"Saved to temporary location: {temp_file}")
            else:
                raise Exception("Could not save file in any location")

def load_existing_data(output_file):
    try:
        from openpyxl import load_workbook
        wb = load_workbook(output_file)
        ws = wb.active
        data = []
        for row in ws.iter_rows(min_row=2, values_only=True):  # Skip header row
            if any(cell is not None for cell in row):  # Only include non-empty rows
                data.append(list(row))
        
        # Calculate the page number based on records (10 records per page)
        last_page = (len(data) // 10) + 1 if data else 1
        print(f"Loaded {len(data)} existing records (approximately page {last_page})")
        return data, last_page
    except Exception as e:
        print(f"No existing data found or error loading file: {e}")
        return [], 1

def scrape_greenbook(output_file="nafdac_greenbook.xlsx", end_page=876, resume=True, driver_path=None):
    # Setup Chrome options
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")  # Larger window for better rendering
    
    # Initialize driver
    if driver_path:
        # Use provided ChromeDriver binary
        driver = webdriver.Chrome(service=Service(driver_path), options=options)
    else:
        # Fallback to webdriver_manager (may query OS for browser version)
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
        
        # Load existing data and determine start page
        data, last_page = load_existing_data(output_file)
        page = last_page if resume else 1
        print(f"Starting from page {page}")
        
        # Navigate to start page
        if page > 1:
            print(f"Navigating to page {page}...")
            current_page = 1
            
            def handle_alerts():
                try:
                    alert = Alert(driver)
                    alert.accept()
                    print("Cleared alert during navigation")
                    time.sleep(1)
                except NoAlertPresentException:
                    pass

            def safe_click_next():
                handle_alerts()
                next_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "li.page-item.next:not(.disabled) a")))
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_button)
                time.sleep(1)
                try:
                    driver.execute_script("arguments[0].click();", next_button)
                except:
                    handle_alerts()
                    driver.execute_script("arguments[0].click();", next_button)
                time.sleep(2)  # Increased wait time after click
                handle_alerts()

            def go_to_page(target_page):
                """Try to jump to `target_page` using the DataTables JS API. Returns True on success."""
                for attempt in range(3):
                    try:
                        # Clear any alerts first
                        try:
                            Alert(driver).accept()
                        except:
                            pass

                        # Try to call DataTables API via JS. Use page index (0-based).
                        js = ("(function(){"
                              "var tblEl = document.querySelector('table.dataTable');"
                              "if(!tblEl) return 'no-table';"
                              "var dt = null;"
                              "try{ dt = $(tblEl).DataTable(); }catch(e){}"
                              "if(!dt){ try{ dt = $.fn.dataTable.Api(tblEl); }catch(e){} }"
                              "if(!dt) return 'no-dt';"
                              "try{ dt.page(%d).draw(false); return 'ok'; }catch(e){ return 'err'; }"
                              "})();") % (target_page - 1)

                        res = driver.execute_script(js)
                        if res == 'ok':
                            # Wait a bit for redraw
                            time.sleep(2)
                            # Wait until active page number updates
                            try:
                                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "li.page-item.active a.page-link")))
                            except:
                                pass
                            return True
                        else:
                            print(f"DataTables jump returned: {res}")
                            return False

                    except UnexpectedAlertPresentException:
                        try:
                            Alert(driver).accept()
                        except:
                            pass
                        time.sleep(1)
                    except Exception as e:
                        print(f"go_to_page attempt error: {e}")
                        time.sleep(1)
                return False
                
            while current_page < page:
                try:
                    # Wait for navigation elements to be present
                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "ul.pagination")))
                    time.sleep(1)  # Give extra time for pagination to stabilize
                    
                    # Get current visible page number
                    active_page = driver.find_element(By.CSS_SELECTOR, "li.page-item.active a.page-link")
                    current_page = int(active_page.text)
                    print(f"Currently at page {current_page}")

                    # Try DataTables API jump directly to the target page (fast) as first attempt
                    try:
                        if go_to_page(page):
                            print(f"Jumped directly to page {page} using DataTables API")
                            current_page = page
                            break
                    except Exception:
                        pass
                    
                    # Try to find the last visible page number button
                    page_buttons = driver.find_elements(By.CSS_SELECTOR, "li.page-item a.page-link")
                    page_numbers = []
                    
                    for btn in page_buttons:
                        try:
                            if btn.text.strip().isdigit():
                                page_numbers.append(int(btn.text.strip()))
                        except:
                            continue
                    
                    if page_numbers:
                        # Find the highest available page number we can click
                        available_targets = [n for n in page_numbers if n > current_page]
                        if available_targets:
                            target = min(available_targets)  # Take the next available page
                            print(f"Found clickable page {target}")
                            
                            # Find and click the target page button
                            for btn in page_buttons:
                                if btn.text.strip().isdigit() and int(btn.text.strip()) == target:
                                    print(f"Clicking page {target} button...")
                                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
                                    time.sleep(0.5)
                                    driver.execute_script("arguments[0].click();", btn)
                                    time.sleep(2)
                                    handle_alerts()
                                    
                                    # Verify the page changed by checking active page number
                                    try:
                                        prev = active_page.text
                                    except:
                                        prev = None

                                    def page_changed(drv):
                                        try:
                                            cur = drv.find_element(By.CSS_SELECTOR, "li.page-item.active a.page-link").text
                                            return cur != prev
                                        except:
                                            return False

                                    try:
                                        wait.until(page_changed)
                                    except:
                                        # Fallback: wait for table rows to be present
                                        try:
                                            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.dataTable tbody tr")))
                                        except:
                                            pass
                                    break
                            continue  # Skip the next button click for this iteration
                    
                    # If no page numbers were clickable or found, use next button
                    print(f"Using next button to advance from page {current_page}...")
                    safe_click_next()
                    
                    # Verify the page changed
                    wait.until(EC.staleness_of(active_page))
                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.dataTable tbody tr")))
                    
                    # Small pause to let the page stabilize
                    time.sleep(1)
                        
                except Exception as e:
                    print(f"Error during navigation: {e}")
                    handle_alerts()
                    time.sleep(2)
                    
                    try:
                        # Get current page after error
                        active_page = driver.find_element(By.CSS_SELECTOR, "li.page-item.active a.page-link")
                        current_page = int(active_page.text)
                        print(f"After error recovery, we are on page {current_page}")
                        
                        if current_page < page:
                            # Refresh the page if we encounter persistent errors
                            print("Refreshing page...")
                            driver.refresh()
                            time.sleep(3)
                            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.dataTable")))
                        else:
                            print("Reached target page despite error")
                            break
                            
                    except:
                        print("Failed to recover, stopping script")
                        return
        
        # Per-page failure tracking so we can skip problematic pages
        page_failures = {}

        while page <= end_page:
            print(f"Scraping page {page}...")
            try:
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
                    except Exception:
                        continue

                # Add to main data list
                data.extend(page_data)
                print(f"Total records collected: {len(data)}")

                # Save progress every 5 pages
                if page % 5 == 0:
                    print(f"Saving checkpoint at page {page}...")
                    save_to_excel(data, output_file)

            except Exception as e:
                # Record failure for this page and try to recover/skip
                fail_count = page_failures.get(page, 0) + 1
                page_failures[page] = fail_count
                print(f"Error scraping page {page}: {e} (failure {fail_count})")

                # Accept any alert and try to continue
                try:
                    Alert(driver).accept()
                    print("Accepted alert during page scrape")
                except Exception:
                    pass

                # If failures exceed threshold, skip this page
                if fail_count >= 3:
                    print(f"Page {page} failing repeatedly — skipping to next page")
                    # Try to jump to the next page using DataTables API; fallback to clicking next
                    try:
                        if go_to_page(page + 1):
                            page += 1
                            continue
                    except Exception:
                        pass

                    # Try clicking next if available
                    try:
                        nb = driver.find_element(By.CSS_SELECTOR, "li.page-item.next:not(.disabled) a")
                        try:
                            driver.execute_script("arguments[0].click();", nb)
                            time.sleep(1)
                            page += 1
                            continue
                        except Exception:
                            pass
                    except Exception:
                        pass

                    # If all else fails, increment page number (best-effort skip)
                    page += 1
                    continue
            
            # Look for next button
            try:
                next_button = driver.find_element(By.CSS_SELECTOR, "li.page-item.next:not(.disabled) a")
            except:
                # Try using DataTables API to go to next page as a fallback
                print("Next button not found — attempting DataTables API jump as fallback...")
                try:
                    if go_to_page(page + 1):
                        page += 1
                        continue
                    else:
                        print("DataTables API jump failed. No more pages found.")
                        break
                except Exception as e:
                    print(f"Fallback DataTables jump error: {e}\nNo more pages found.")
                    break
                
            # Click next button and wait for table update
            max_retries = 3
            retry_count = 0
            while retry_count < max_retries:
                try:
                    # Handle any existing alerts before proceeding
                    try:
                        alert = Alert(driver)
                        alert.accept()
                        print("Cleared existing alert")
                    except NoAlertPresentException:
                        pass

                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_button)
                    time.sleep(1)  # Increased pause before click
                    
                    # Click using JavaScript to avoid potential intercepted click
                    # click and wait until the active page number changes (more reliable than staleness_of)
                    try:
                        prev_active = driver.find_element(By.CSS_SELECTOR, "li.page-item.active a.page-link").text
                    except:
                        prev_active = None

                    driver.execute_script("arguments[0].click();", next_button)

                    def active_page_changed(drv):
                        try:
                            cur = drv.find_element(By.CSS_SELECTOR, "li.page-item.active a.page-link").text
                            return cur != prev_active
                        except:
                            return False

                    try:
                        wait.until(active_page_changed)
                    except:
                        # Fallback: wait for table rows to be present
                        try:
                            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.dataTable tbody tr")))
                        except:
                            pass

                    time.sleep(1)  # small pause for table to fully update
                    page += 1
                    break  # Success, exit retry loop
                    
                except UnexpectedAlertPresentException:
                    print(f"Alert encountered, accepting it and retrying... (Attempt {retry_count + 1}/{max_retries})")
                    try:
                        alert = Alert(driver)
                        alert.accept()
                    except:
                        pass
                    time.sleep(2)
                    retry_count += 1
                    
                except Exception as e:
                    print(f"Error navigating to next page: {e}")
                    if retry_count < max_retries - 1:
                        print(f"Retrying... (Attempt {retry_count + 1}/{max_retries})")
                        time.sleep(2)
                        retry_count += 1
                    else:
                        print("Max retries reached, stopping scrape")
                        return
        
        # Save final data to Excel
        print(f"Saving {len(data)} records to {output_file}")
        save_to_excel(data, output_file)
        
        print("Scraping completed successfully")
        
    except Exception as e:
        print(f"Error occurred: {str(e)}")
    finally:
        driver.quit()

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="NAFDAC Greenbook scraper")
    parser.add_argument("--start", type=int, help="Start page (overrides resume detection)")
    parser.add_argument("--end", type=int, default=876, help="End page")
    parser.add_argument("--driver", type=str, help="Full path to chromedriver executable to avoid auto-download")
    parser.add_argument("--file", type=str, default="nafdac_greenbook.xlsx", help="Output Excel file path")
    args = parser.parse_args()

    # If user provided a start page, we'll overwrite resume behavior and start there
    if args.start:
        # call with resume=False and set last_page to start
        scrape_greenbook(output_file=args.file, end_page=args.end, resume=False, driver_path=args.driver)
    else:
        scrape_greenbook(output_file=args.file, end_page=args.end, resume=True, driver_path=args.driver)
