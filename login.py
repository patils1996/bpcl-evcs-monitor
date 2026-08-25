import os
import sys
from playwright.sync_api import sync_playwright

def run():
    print("=" * 60)
    print("BPCL Charger Status Monitor - Login Session Capture")
    print("=" * 60)
    
    with sync_playwright() as p:
        # Launch headed browser so the user can interact with the page
        print("Launching browser. Please look for the opened browser window...")
        browser = p.chromium.launch(headless=False)
        
        # Create context
        context = browser.new_context(viewport={"width": 1280, "height": 800})
        page = context.new_page()
        
        url = "https://csms.bpcl.statiq.co.in/charger-management/all-chargers"
        print(f"Navigating to {url}...")
        page.goto(url)
        
        print("\n" + "*" * 60)
        print("ACTION REQUIRED:")
        print("1. In the browser window, log in using your Phone Number and Password.")
        print("2. Complete any Multi-Factor Authentication (OTP) if prompted.")
        print("3. Navigate to 'Charger Management' -> 'Chargers' if not redirected automatically.")
        print("4. Make sure the table of chargers is visible on the screen.")
        print("5. Once you are successfully logged in, come back to this terminal and press ENTER.")
        print("*" * 60 + "\n")
        
        # Wait for the user to press Enter in the console
        input("Press ENTER after you have logged in and can see the chargers list...")
        
        # Save storage state
        state_path = "state.json"
        context.storage_state(path=state_path)
        print(f"\n[SUCCESS] Session state successfully saved to '{state_path}'!")
        
        browser.close()

if __name__ == "__main__":
    try:
        run()
    except KeyboardInterrupt:
        print("\nOperation cancelled by user.")
        sys.exit(0)
    except Exception as e:
        print(f"\nAn error occurred: {e}")
        sys.exit(1)
