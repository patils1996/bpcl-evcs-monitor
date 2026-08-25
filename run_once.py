from playwright.sync_api import sync_playwright
import os
import sys
import time
from datetime import datetime

# Add local dir to path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

import monitor

def main():
    state_file = "state.json"
    if not os.path.exists(state_file):
        print(f"Error: {state_file} not found.")
        sys.exit(1)
        
    print("Running one-shot charger check...")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            storage_state=state_file,
            viewport={"width": 1920, "height": 1080}
        )
        page = context.new_page()
        
        print(f"Navigating to {monitor.URL}...")
        page.goto(monitor.URL, wait_until="load", timeout=60000)
        time.sleep(5)
        
        # Change page size to 100
        monitor.change_rows_per_page_to_100(page)
        
        # Scrape data
        scraped_data = monitor.scrape_chargers(page)
        print(f"Scraped {len(scraped_data)} chargers total.")
        
        if scraped_data:
            excel_path = monitor.EXCEL_PATH
            live_statuses = {c["OCPP ID"]: c["Status"] for c in scraped_data}
            
            # Update Excel file
            chargers_list = monitor.update_excel_vlookup_and_summary(excel_path, live_statuses)
            
            # Update HTML Dashboard
            monitor.update_dashboard_html(chargers_list, excel_path)
            print("Excel sheet and HTML Dashboard successfully updated!")
        else:
            print("No data scraped.")
            
        browser.close()

if __name__ == "__main__":
    main()
