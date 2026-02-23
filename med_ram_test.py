"""
Madinah Ramadan Room Availability Bot
Automates copying room availability data from bookingarabian.com to Google Sheets
Each hotel tab has its own date range - bot fetches data accordingly
"""

import time
import re
import os
import json
import tempfile
from datetime import datetime, timedelta
from playwright.sync_api import sync_playwright
import gspread

# ==================== CONFIGURATION ====================
USERNAME = os.environ.get("BOT_USERNAME", "ai")
PASSWORD = os.environ.get("BOT_PASSWORD", "")
WEBSITE_URL = os.environ.get("WEBSITE_URL", "https://bookingarabian.com/")

CITY = "MADINAH"

# Google Sheets Configuration
# If GOOGLE_CREDENTIALS_JSON env var is set (GitHub Actions), write it to a temp file
_creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON", "")
if _creds_json:
    _tmp = tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False)
    _tmp.write(_creds_json)
    _tmp.close()
    CREDENTIALS_FILE = _tmp.name
else:
    CREDENTIALS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "credentials.json")

SPREADSHEET_ID = os.environ.get("MED_SPREADSHEET_ID")
if not SPREADSHEET_ID:
    raise ValueError("MED_SPREADSHEET_ID environment variable must be set")

# Hotel mapping: Google Sheet Tab Name -> Website Hotel Name
HOTEL_MAPPING = {
    "Maden Hotel": "Maden Hotel",
    "AMP": "ANWAR",
    "SAJA": "SAJA MADINAH",
    "AL HARAM": "DAR AL EIMAN AL HARAM",
    "PULLMAN": "PULLMAN ZAMZAM MADINAH",
    "AQEEQ ": "AL AQEEQ MADINAH",
    " FRONT": "TAIBA FRONT HOTEL",
    " HARTHIYA": "FRONTEL AL HARITHIA",
    "MONA KAREEM": "LEADER AL MUNA KAREEM",
    "SAFWAT": "SAFWAT",
    "BADAR MAQAM": "GRAND PLAZA BADR",
    "RUA INT": "RUA INT",
    " KAYAN INT": "KAYAN INT",
    "ANSAR TULIP": "AL ANSAR GOLDEN TULIP",
    "MADINAH CONCORD": "MADINAH CONCORD",
    " CONCORD KHAIR": "CONCORDE HOTEL DAR AL KHAIR",
    "ABRAJ TABAH": "ABRAJ TABA",
    "VALLEY HOTEL": "VALY HOTEL",
    "MAIEN TAIBA HOTEL": "MAIEN TAIBA HOTEL",
    "GULNAR": "GULNAR",
    "NUSUK": "NUSUK",
    "RAMA": "RAMA AL MADINA",
    "TAJ WARD": "TAJ WARD",
    "JAWHRAT RASHEED": "JAWHARAT AL RASHEED",
    "MUKHTARA GOLDEN": "MUKHTARA GOLDEN",
    "RUA DIYAFAH": "RUA AL DIYAFAH HOTEL",
    "TAIBAH HILLZ": "TAIBAH HILLS HOTEL",
    "BIR": "BIR",
    "MIRAMAR": "MIRAMAR",
    "SHAZA": "SHAZA",
    "QADAT": "QADAT",
}

# Google Sheet structure settings
DATE_COLUMN = 2           # Column B - dates
TOT_RMS_COLUMN = 3        # Column C - Total Allotment goes here
SOLD_RMS_COLUMN = 4       # Column D - Total Sales goes here
DATA_START_ROW = 17       # First data row in Google Sheets


def get_sheet_date_ranges():
    """Read each tab from Google Sheets and extract its date range from column B"""
    print("\n📌 Reading Google Sheets to get date ranges...")
    
    if not SPREADSHEET_ID:
        print("❌ SPREADSHEET_ID not configured!")
        return {}
    
    gc = gspread.service_account(filename=CREDENTIALS_FILE)
    spreadsheet = gc.open_by_key(SPREADSHEET_ID)
    print(f"   Connected to: {spreadsheet.title}")
    
    sheet_info = {}
    
    for tab_name, website_name in HOTEL_MAPPING.items():
        try:
            worksheet = spreadsheet.worksheet(tab_name)
        except gspread.exceptions.WorksheetNotFound:
            # Try stripped version
            stripped = tab_name.strip()
            try:
                worksheet = spreadsheet.worksheet(stripped)
            except gspread.exceptions.WorksheetNotFound:
                print(f"   ⚠️ Tab '{tab_name}' not found in Google Sheets - skipping")
                continue
        
        # Read ALL values from column B (with rate limit delay)
        date_col = worksheet.col_values(DATE_COLUMN)
        time.sleep(2)  # Avoid Google Sheets API rate limit
        
        # Scan from row 15 onward, auto-detect where dates start
        dates = []
        found_dates = False
        for row_idx in range(14, len(date_col)):  # 0-indexed, start from row 15
            val = date_col[row_idx].strip() if row_idx < len(date_col) else ''
            
            # Skip empty cells and header cells before dates start
            if not val:
                if found_dates:
                    break  # Empty cell after dates = end of data
                continue
            
            # Try to parse as date
            parsed_date = None
            for fmt in ['%d-%b-%y', '%d-%b-%Y', '%d/%m/%Y', '%d/%m/%y', '%Y-%m-%d',
                        '%d %b %Y', '%d %b %y', '%m/%d/%Y', '%m/%d/%y',
                        '%d-%B-%y', '%d-%B-%Y', '%B %d, %Y', '%b %d, %Y']:
                try:
                    parsed_date = datetime.strptime(val, fmt)
                    if parsed_date.year < 100:
                        parsed_date = parsed_date.replace(year=parsed_date.year + 2000)
                    if parsed_date.year == 2025:
                        parsed_date = parsed_date.replace(year=2026)
                    break
                except (ValueError, TypeError):
                    continue
            
            if parsed_date:
                found_dates = True
                actual_row = row_idx + 1  # Convert back to 1-indexed
                dates.append((actual_row, parsed_date))
            elif found_dates:
                break  # Non-date cell after dates = end of data
        
        if dates:
            first_date = dates[0][1]
            last_date = dates[-1][1]
            sheet_info[tab_name] = {
                'hotel_name': website_name,
                'dates': dates,
                'first_date': first_date,
                'last_date': last_date,
                'num_days': len(dates),
            }
            print(f"   {tab_name}: {first_date.strftime('%d/%m/%Y')} to {last_date.strftime('%d/%m/%Y')} ({len(dates)} days)")
        else:
            print(f"   {tab_name}: No data rows found - skipping")
    
    return sheet_info


def extract_sales_and_allotment(page):
    """Extract dates, Total Sales, and Total Allotment from the table.
    Returns a dict: {date_str: {'sales': value, 'allotment': value}}
    """
    raw = page.evaluate('''
        () => {
            const result = { dates: [], totalSales: [], totalAllotment: [] };
            const table = document.querySelector('table');
            if (!table) return result;
            
            // Get dates from the first header row (th elements)
            const headerRows = table.querySelectorAll('thead tr, tr');
            for (const row of headerRows) {
                const ths = row.querySelectorAll('th');
                if (ths.length > 2) {
                    for (let i = 1; i < ths.length; i++) {
                        const text = ths[i].innerText.trim();
                        const match = text.match(/(\\d{1,2})-(\\d{1,2})/);
                        if (match) {
                            result.dates.push(match[0]);
                        }
                    }
                    if (result.dates.length > 0) break;
                }
            }
            
            // Find Total Sales and Total Allotment rows
            const allRows = table.querySelectorAll('tr');
            for (const row of allRows) {
                const cells = row.querySelectorAll('td');
                if (cells.length > 1) {
                    const label = cells[0].innerText.trim().toUpperCase();
                    
                    if (label.includes('TOTAL SALES') || label === 'TOTAL SALES') {
                        for (let i = 1; i < cells.length; i++) {
                            const num = parseInt(cells[i].innerText.trim());
                            result.totalSales.push(isNaN(num) ? 0 : num);
                        }
                    }
                    
                    if (label.includes('TOTAL ALLOTMENT') || label === 'TOTAL ALLOTMENT') {
                        for (let i = 1; i < cells.length; i++) {
                            const num = parseInt(cells[i].innerText.trim());
                            result.totalAllotment.push(isNaN(num) ? 0 : num);
                        }
                    }
                }
            }
            
            return result;
        }
    ''')
    
    # Build date-to-values mapping
    date_value_map = {}
    dates = raw.get('dates', [])
    sales = raw.get('totalSales', [])
    allotment = raw.get('totalAllotment', [])
    
    current_year = datetime.now().year
    
    for i, date_short in enumerate(dates):
        parts = date_short.split('-')
        if len(parts) == 2:
            full_date = f"{parts[0]}/{parts[1]}/{current_year}"
            date_value_map[full_date] = {
                'sales': sales[i] if i < len(sales) else 0,
                'allotment': allotment[i] if i < len(allotment) else 0
            }
    
    return date_value_map


def update_google_sheets(sheet_info, hotel_sales_data):
    """Update Google Sheets with Total Sales and Total Allotment data.
    Uses the date rows already read from Google Sheets (sheet_info) 
    to write extracted data back to the correct rows.
    """
    print("\n📌 Updating Google Sheets...")
    
    if not SPREADSHEET_ID:
        print("   ⚠️ SPREADSHEET_ID not configured - skipping")
        return False
    
    try:
        gc = gspread.service_account(filename=CREDENTIALS_FILE)
        spreadsheet = gc.open_by_key(SPREADSHEET_ID)
        print(f"   Connected to: {spreadsheet.title}")
        
        total_updates = 0
        
        for tab_name, info in sheet_info.items():
            print(f"\n   Processing: {tab_name}")
            
            date_values = hotel_sales_data.get(tab_name)
            if not date_values:
                print(f"      ⚠️ No data extracted for {tab_name}")
                continue
            
            try:
                worksheet = spreadsheet.worksheet(tab_name)
            except gspread.exceptions.WorksheetNotFound:
                stripped = tab_name.strip()
                try:
                    worksheet = spreadsheet.worksheet(stripped)
                except gspread.exceptions.WorksheetNotFound:
                    print(f"      ⚠️ Tab '{tab_name}' not found in Google Sheets")
                    continue
            
            cells_to_update = []
            
            for row, gs_date in info['dates']:
                # Match with extracted dates
                for extracted_date, values in date_values.items():
                    try:
                        parts = extracted_date.split('/')
                        if len(parts) == 3:
                            ext_day = int(parts[0])
                            ext_month = int(parts[1])
                            if gs_date.day == ext_day and gs_date.month == ext_month:
                                allotment_val = values.get('allotment', 0)
                                sales_val = values.get('sales', 0)
                                cells_to_update.append(gspread.Cell(row, TOT_RMS_COLUMN, allotment_val))
                                cells_to_update.append(gspread.Cell(row, SOLD_RMS_COLUMN, sales_val))
                                break
                    except Exception:
                        continue
            
            if cells_to_update:
                worksheet.update_cells(cells_to_update)
                total_updates += len(cells_to_update) // 2
                print(f"      ✅ Updated {len(cells_to_update) // 2} rows (TOT RMS + SOLD RMS)")
            else:
                print(f"      ⚠️ No cells to update")
        
        print(f"\n✅ Google Sheets updated! Total rows: {total_updates}")
        print(f"📊 View at: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}")
        return True
        
    except Exception as e:
        print(f"\n❌ Google Sheets error: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    print("=" * 60)
    print("Madinah Ramadan Room Availability Bot")
    print("=" * 60)
    
    # Step 0: Read Google Sheets to get date ranges for each hotel
    sheet_info = get_sheet_date_ranges()
    
    if not sheet_info:
        print("❌ No valid sheets with dates found!")
        return
    
    # Print date ranges per hotel
    print(f"\n📅 Date ranges per hotel ({len(sheet_info)} hotels):")
    for sn, info in sheet_info.items():
        days = (info['last_date'] - info['first_date']).days + 1
        print(f"   {sn}: {info['first_date'].strftime('%d/%m/%Y')} to {info['last_date'].strftime('%d/%m/%Y')} ({days} days)")
    
    # Start browser
    playwright = sync_playwright().start()
    browser = playwright.chromium.launch(headless=True, slow_mo=500)
    context = browser.new_context(viewport={'width': 1920, 'height': 1080})
    page = context.new_page()
    page.set_default_timeout(60000)
    
    try:
        # Step 1: Login
        print("\n📌 Step 1: Logging in...")
        page.goto(WEBSITE_URL)
        page.wait_for_load_state("networkidle")
        time.sleep(2)
        
        page.click('input[name="username"]')
        time.sleep(0.3)
        page.fill('input[name="username"]', USERNAME)
        print(f"   Entered username: {USERNAME}")
        
        page.click('input[name="password"]')
        time.sleep(0.3)
        page.keyboard.type(PASSWORD)
        print("   Entered password")
        
        page.click('button:has-text("Sign in")')
        page.wait_for_load_state("networkidle")
        time.sleep(5)
        print("✅ Login successful!")
        
        # Step 2: Navigate to Availability Consolidated
        print("\n📌 Step 2: Navigating to Availability Consolidated...")
        time.sleep(2)
        page.locator("section").get_by_text("Availability Consolidated").click()
        page.wait_for_load_state("networkidle")
        time.sleep(3)
        print("✅ Availability Consolidated page loaded")
        
        # Step 5: Extract data for each hotel
        print(f"\n📌 Step 5: Extracting data for {len(sheet_info)} hotels...")
        hotel_sales_data = {}
        
        for tab_name, hotel_info in sheet_info.items():
            website_hotel_name = HOTEL_MAPPING.get(tab_name)
            if not website_hotel_name:
                print(f"\n   ⚠️ No website mapping for '{tab_name}' - skipping")
                continue
            
            print(f"\n   🏨 Processing hotel: {tab_name} -> {website_hotel_name}")
            
            hotel_start = hotel_info['first_date']
            hotel_end = hotel_info['last_date']
            hotel_days = (hotel_end - hotel_start).days + 1
            num_chunks = (hotel_days + 29) // 30
            
            print(f"   📅 Date range: {hotel_start.strftime('%d/%m/%Y')} to {hotel_end.strftime('%d/%m/%Y')} ({hotel_days} days, {num_chunks} chunk{'s' if num_chunks > 1 else ''})")
            
            try:
                # Break date range into 30-day chunks
                chunk_start = hotel_start
                chunk_num = 0
                hotel_date_values = {}
                
                while chunk_start <= hotel_end:
                    chunk_num += 1
                    chunk_end_calc = chunk_start + timedelta(days=29)
                    chunk_end = min(chunk_end_calc, hotel_end)
                    chunk_days = (chunk_end - chunk_start).days + 1
                    
                    print(f"\n   📦 Chunk {chunk_num}: {chunk_start.strftime('%d/%m/%Y')} to {chunk_end.strftime('%d/%m/%Y')} ({chunk_days} days)")
                    
                    # Step A: Set the date for this chunk
                    time.sleep(2)
                    date_str = chunk_start.strftime('%d/%m/%Y')
                    print(f"   Setting date to {date_str}...")
                    try:
                        date_input = page.locator('input.mat-datepicker-input, input[matinput], input[placeholder*="Date"], input[aria-label*="Date"]').first
                        date_input.wait_for(state="visible", timeout=10000)
                        date_input.click()
                        time.sleep(1)
                        
                        page.keyboard.press("Control+A")
                        page.keyboard.press("Backspace")
                        time.sleep(0.5)
                        page.keyboard.type(date_str)
                        time.sleep(0.3)
                        page.keyboard.press("Enter")
                        
                        print(f"   ✅ Date set to {date_str}")
                        time.sleep(2)
                    except Exception as e:
                        print(f"   ⚠️ Date insertion error: {e}")
                    
                    # Step B: Select the hotel
                    page.locator(".dropdown-toggle").click()
                    time.sleep(1)
                    
                    hotel_input = page.locator('input[type="text"]').last
                    hotel_input.fill(website_hotel_name)
                    time.sleep(2)
                    
                    try:
                        option = page.get_by_role("option", name=website_hotel_name).first
                        option.wait_for(state="visible", timeout=5000)
                        option.click()
                    except Exception:
                        option = page.get_by_role("link", name=website_hotel_name).first
                        option.wait_for(state="visible", timeout=5000)
                        option.click()
                    time.sleep(1)
                    
                    print(f"   ✅ Selected hotel: {website_hotel_name}")
                    
                    # Step C: Click View Record
                    print("   Clicking View Record...")
                    time.sleep(1)
                    page.get_by_role("button", name="View Record").click()
                    page.wait_for_load_state("networkidle")
                    time.sleep(5)
                    print(f"   ✅ Data loaded for chunk {chunk_num}!")
                    
                    # Step D: Extract data from this chunk
                    page_data = extract_sales_and_allotment(page)
                    print(f"   Found data for {len(page_data)} dates in this chunk")
                    
                    # Merge into hotel's data
                    hotel_date_values.update(page_data)
                    
                    # Move to next chunk (start from day 31, 61, etc.)
                    chunk_start = chunk_start + timedelta(days=30)
                
                hotel_sales_data[tab_name] = hotel_date_values
                print(f"\n   ✅ Total data extracted for {tab_name}: {len(hotel_date_values)} dates")
                for d, v in list(hotel_date_values.items())[:5]:
                    print(f"      {d} = {v}")
                if len(hotel_date_values) > 5:
                    print(f"      ... and {len(hotel_date_values) - 5} more")
                
            except Exception as e:
                print(f"   ⚠️ Error processing hotel {website_hotel_name}: {e}")
                continue
        
        # Show summary
        print(f"\n✅ Extracted data for {len(hotel_sales_data)} hotels:")
        for name, data in hotel_sales_data.items():
            print(f"   - {name}: {len(data)} dates")
        
        # Step 6: Update Google Sheets
        update_google_sheets(sheet_info, hotel_sales_data)
        
        print("\n" + "=" * 60)
        print("✅ BOT COMPLETED SUCCESSFULLY!")
        print(f"   Google Sheet: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}")
        print("=" * 60)
        
    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()
        
        try:
            page.screenshot(path="error_screenshot.png")
            print("📸 Error screenshot saved")
        except:
            pass
    
    finally:
        print("\n📌 Closing browser...")
        context.close()
        browser.close()
        playwright.stop()
        print("✅ Done")


if __name__ == "__main__":
    main()
