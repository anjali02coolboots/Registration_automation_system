import os
from datetime import datetime, timedelta

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import credentials_store as cs


def ensure_dirs():
    os.makedirs(cs.OUTPUT_DIR, exist_ok=True)


def is_captcha_page(page) -> bool:
    t = page.content().lower()
    return ("are you a human being" in t) or ("enter captcha" in t) or ("captcha" in t and "proceed" in t)


def ensure_not_captcha(page):
    if not is_captcha_page(page):
        return

    print("\n[CAPTCHA] CAPTCHA detected. Solve it in the open browser and click Proceed.")
    
    # In GitHub Actions, we can't solve CAPTCHA manually
    if os.getenv('GITHUB_ACTIONS'):
        print("[CAPTCHA] Running in GitHub Actions - cannot solve CAPTCHA manually!")
        print("[CAPTCHA] Consider using authenticated session or API if available.")
        raise RuntimeError("CAPTCHA detected in automated environment. Cannot proceed.")
    
    page.wait_for_function(
        """() => {
            const t = (document.body && document.body.innerText || "").toLowerCase();
            return !(t.includes("are you a human being") || t.includes("enter captcha") || (t.includes("captcha") && t.includes("proceed")));
        }""",
        timeout=10 * 60 * 1000
    )
    print("[CAPTCHA] Solved. Continuing...\n")


def login(page):
    print("[LOGIN] Navigating to login page...")
    page.goto(cs.LOGIN_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(1500)  # Increased wait time
    ensure_not_captcha(page)

    print("[LOGIN] Looking for login fields...")
    login_id = page.locator('xpath=//*[contains(normalize-space(),"Login ID")]/following::input[1]')
    pwd = page.locator('xpath=//*[contains(normalize-space(),"Password")]/following::input[1]')
    
    # Wait for fields to be visible
    try:
        login_id.first.wait_for(state="visible", timeout=10000)
        pwd.first.wait_for(state="visible", timeout=10000)
    except PlaywrightTimeoutError:
        print("[LOGIN] ERROR: Login fields not visible")
        print("[LOGIN] Page content preview:")
        print(page.content()[:500])
        raise RuntimeError("Login fields not found on login page.")
    
    if login_id.count() == 0 or pwd.count() == 0:
        raise RuntimeError("Login fields not found on login page.")

    print("[LOGIN] Filling credentials...")
    login_id.first.fill(cs.TECHGIG_USERNAME)
    pwd.first.fill(cs.TECHGIG_PASSWORD)

    print("[LOGIN] Submitting login form...")
    submit = page.locator('xpath=//button[normalize-space()="Submit"] | //input[@type="submit"]')
    if submit.count() > 0:
        submit.first.click()
    else:
        pwd.first.press("Enter")

    try:
        page.wait_for_load_state("networkidle", timeout=cs.NAV_TIMEOUT_MS)
    except PlaywrightTimeoutError:
        print("[LOGIN] Network idle timeout (continuing anyway)")
        pass

    page.wait_for_timeout(1500)
    ensure_not_captcha(page)
    
    print("[LOGIN] Login completed successfully")


def _set_select_and_verify(page, css_id: str, value: str):
    """Set dropdown value and verify it was set correctly."""
    print(f"[DROPDOWN] Setting {css_id} to {value}")
    
    # Wait for element to be visible and enabled
    try:
        page.locator(css_id).wait_for(state="visible", timeout=5000)
        page.wait_for_timeout(300)
    except PlaywrightTimeoutError:
        print(f"[DROPDOWN] WARNING: {css_id} not visible within timeout")
        raise
    
    page.locator(css_id).select_option(value=value)
    page.wait_for_timeout(500)  # Increased wait time
    
    actual = page.locator(css_id).evaluate("el => el.value")
    if str(actual) != str(value):
        raise RuntimeError(f"Dropdown {css_id} did not change. Expected {value}, got {actual}.")
    
    print(f"[DROPDOWN] ✓ {css_id} = {actual}")


def wait_for_loading_indicators(page):
    """Wait for common loading indicators to disappear."""
    try:
        page.wait_for_selector('.loading, .spinner, .overlay', state='hidden', timeout=3000)
    except PlaywrightTimeoutError:
        pass


def set_date_range_and_search(page, start_dt: datetime, end_dt: datetime):
    print(f"\n[DATE RANGE] Setting dates: {start_dt.strftime('%Y-%m-%d')} to {end_dt.strftime('%Y-%m-%d')}")
    
    print("[DATE RANGE] Navigating to stats page...")
    page.goto(cs.STATS_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(2000)  # Increased wait time for page load
    ensure_not_captcha(page)

    print("[DATE RANGE] Waiting for date selectors...")
    
    # Wait for all date selectors with better error handling
    selectors = ["#start_day", "#start_month", "#start_year", "#end_day", "#end_month", "#end_year"]
    
    for selector in selectors:
        try:
            page.wait_for_selector(selector, state="visible", timeout=15000)
            print(f"[DATE RANGE] ✓ Found {selector}")
        except PlaywrightTimeoutError:
            print(f"[DATE RANGE] ✗ TIMEOUT waiting for {selector}")
            print(f"[DATE RANGE] Page URL: {page.url}")
            print(f"[DATE RANGE] Page title: {page.title()}")
            
            # Try to find any select elements on the page
            all_selects = page.locator("select").count()
            print(f"[DATE RANGE] Total <select> elements found: {all_selects}")
            
            # Save screenshot for debugging
            screenshot_path = os.path.join(cs.OUTPUT_DIR, "error_screenshot.png")
            page.screenshot(path=screenshot_path)
            print(f"[DATE RANGE] Screenshot saved: {screenshot_path}")
            
            # Save page HTML for debugging
            html_path = os.path.join(cs.OUTPUT_DIR, "error_page.html")
            with open(html_path, 'w', encoding='utf-8') as f:
                f.write(page.content())
            print(f"[DATE RANGE] HTML saved: {html_path}")
            
            raise RuntimeError(f"Date selector {selector} not found. Check screenshot and HTML in exports folder.")

    page.wait_for_timeout(1000)

    sd, sm, sy = str(start_dt.day), str(start_dt.month), str(start_dt.year)
    ed, em, ey = str(end_dt.day), str(end_dt.month), str(end_dt.year)

    print(f"[DATE RANGE] Setting start date: {sy}-{sm}-{sd}")
    _set_select_and_verify(page, "#start_year", sy)
    page.wait_for_timeout(500)
    
    _set_select_and_verify(page, "#start_month", sm)
    page.wait_for_timeout(500)
    
    _set_select_and_verify(page, "#start_day", sd)
    page.wait_for_timeout(500)

    print(f"[DATE RANGE] Setting end date: {ey}-{em}-{ed}")
    _set_select_and_verify(page, "#end_year", ey)
    page.wait_for_timeout(500)
    
    _set_select_and_verify(page, "#end_month", em)
    page.wait_for_timeout(500)
    
    _set_select_and_verify(page, "#end_day", ed)
    page.wait_for_timeout(500)

    wait_for_loading_indicators(page)
    page.wait_for_timeout(1000)

    print("[DATE RANGE] Verifying dates before search...")
    start_year_val = page.locator("#start_year").evaluate("el => el.value")
    start_month_val = page.locator("#start_month").evaluate("el => el.value")
    start_day_val = page.locator("#start_day").evaluate("el => el.value")
    end_year_val = page.locator("#end_year").evaluate("el => el.value")
    end_month_val = page.locator("#end_month").evaluate("el => el.value")
    end_day_val = page.locator("#end_day").evaluate("el => el.value")
    
    print(f"[DATE RANGE] Verified start: {start_year_val}-{start_month_val}-{start_day_val}")
    print(f"[DATE RANGE] Verified end: {end_year_val}-{end_month_val}-{end_day_val}")

    print("[DATE RANGE] Looking for Search button...")
    container = page.locator("#start_day").locator("xpath=ancestor::div[1]")
    search_btn = container.locator(
        'xpath=.//button[normalize-space()="Search"] | .//input[@value="Search"] | .//a[normalize-space()="Search"]'
    )

    if search_btn.count() == 0:
        container = page.locator("#start_day").locator("xpath=ancestor::div[2]")
        search_btn = container.locator(
            'xpath=.//button[normalize-space()="Search"] | .//input[@value="Search"] | .//a[normalize-space()="Search"]'
        )

    if search_btn.count() == 0:
        search_btn = page.locator(
            'xpath=//button[normalize-space()="Search"] | //input[@value="Search"] | //a[normalize-space()="Search"]'
        ).first

    print("[DATE RANGE] Clicking Search button...")
    search_btn.click()

    page.wait_for_timeout(2000)  # Increased wait time

    try:
        page.wait_for_load_state("networkidle", timeout=cs.NAV_TIMEOUT_MS)
    except PlaywrightTimeoutError:
        print("[DATE RANGE] Network idle timeout (continuing anyway)")
        pass

    page.wait_for_timeout(2000)
    
    wait_for_loading_indicators(page)
    ensure_not_captcha(page)
    
    print("[DATE RANGE] Search completed\n")


def click_third_row_total_and_download(page) -> str:
    """
    Click on the 3rd row (excluding header and Average row) of Total of Registration column.
    This will trigger a download of yesterday's data.
    """
    print(f"[DOWNLOAD] Looking for 3rd row in Total of Registration column...")
    
    # Wait for table to be visible
    page.wait_for_selector('table', state='visible', timeout=10000)
    page.wait_for_timeout(2000)  # Increased wait time
    
    # Find all clickable links in the "Total of Registration" column (2nd column)
    all_total_links = page.locator('xpath=//table//tr//td[2]//a')
    
    total_count = all_total_links.count()
    print(f"[DOWNLOAD] Found {total_count} links in Total of Registration column")
    
    if total_count == 0:
        # Try alternative: look for any table cell with a link
        print("[DOWNLOAD] Trying alternative method to find links...")
        all_total_links = page.locator('xpath=//table//tbody//tr//td//a[normalize-space()]')
        total_count = all_total_links.count()
        print(f"[DOWNLOAD] Found {total_count} total links in table")
    
    if total_count < 3:
        # Save screenshot for debugging
        screenshot_path = os.path.join(cs.OUTPUT_DIR, "download_error_screenshot.png")
        page.screenshot(path=screenshot_path)
        print(f"[DOWNLOAD] Screenshot saved: {screenshot_path}")
        
        raise RuntimeError(f"Found only {total_count} links, need at least 3. Cannot find 3rd row.")
    
    # Get the 2nd link (0-indexed, so index 1) - this represents yesterday
    target_link = all_total_links.nth(1)
    
    # Get the text of the link to confirm we're clicking the right one
    link_text = target_link.text_content()
    print(f"[DOWNLOAD] Found target row link with value: {link_text}")
    
    # Set up download expectation before clicking
    with page.expect_download(timeout=60000) as dl_info:
        target_link.click()
        page.wait_for_timeout(3000)  # Increased wait time
    
    dl = dl_info.value
    ext = os.path.splitext(dl.suggested_filename)[1]
    
    # Save with the specified name
    filename = f"Registered_User_Source_Summary{ext}"
    path = os.path.join(cs.OUTPUT_DIR, filename)
    dl.save_as(path)
    
    print(f"[DOWNLOAD] Download completed: {filename}")
    return path


def main():
    ensure_dirs()

    today = datetime.today()
    yesterday = today - timedelta(days=1)
    
    # Check if running in GitHub Actions
    is_github_actions = os.getenv('GITHUB_ACTIONS') == 'true'
    
    if is_github_actions:
        print("\n" + "="*60)
        print("RUNNING IN GITHUB ACTIONS ENVIRONMENT")
        print("="*60)
        print("Headless mode: Enabled")
        print("CAPTCHA handling: Will fail if detected")
        print("="*60 + "\n")

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=cs.HEADLESS,
            args=['--no-sandbox', '--disable-setuid-sandbox'] if is_github_actions else []
        )
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        page.set_default_timeout(cs.NAV_TIMEOUT_MS)

        try:
            login(page)

            print("\n" + "="*60)
            print("DOWNLOADING YESTERDAY'S REGISTRATION REPORT")
            print("="*60)
            
            # Set date range to include yesterday
            # 7-day window to ensure we have enough rows in the table
            start_date = yesterday - timedelta(days=6)
            set_date_range_and_search(page, start_date, today)
            
            # Click on the 3rd row's total and download
            report_path = click_third_row_total_and_download(page)
            print("✅ Downloaded (yesterday's data):", report_path)

        except Exception as e:
            print("\n" + "="*60)
            print("ERROR DURING SCRAPING")
            print("="*60)
            print(f"Error: {str(e)}")
            
            # Save debug information
            try:
                screenshot_path = os.path.join(cs.OUTPUT_DIR, "final_error_screenshot.png")
                page.screenshot(path=screenshot_path)
                print(f"Screenshot saved: {screenshot_path}")
            except:
                pass
            
            raise
        finally:
            context.close()
            browser.close()

    return report_path


if __name__ == "__main__":
    main()