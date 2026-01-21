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
    page.wait_for_function(
        """() => {
            const t = (document.body && document.body.innerText || "").toLowerCase();
            return !(t.includes("are you a human being") || t.includes("enter captcha") || (t.includes("captcha") && t.includes("proceed")));
        }""",
        timeout=10 * 60 * 1000
    )
    print("[CAPTCHA] Solved. Continuing...\n")


def login(page):
    page.goto(cs.LOGIN_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(600)
    ensure_not_captcha(page)

    login_id = page.locator('xpath=//*[contains(normalize-space(),"Login ID")]/following::input[1]')
    pwd = page.locator('xpath=//*[contains(normalize-space(),"Password")]/following::input[1]')
    if login_id.count() == 0 or pwd.count() == 0:
        raise RuntimeError("Login fields not found on login page.")

    login_id.first.fill(cs.TECHGIG_USERNAME)
    pwd.first.fill(cs.TECHGIG_PASSWORD)

    submit = page.locator('xpath=//button[normalize-space()="Submit"] | //input[@type="submit"]')
    if submit.count() > 0:
        submit.first.click()
    else:
        pwd.first.press("Enter")

    try:
        page.wait_for_load_state("networkidle", timeout=cs.NAV_TIMEOUT_MS)
    except PlaywrightTimeoutError:
        pass

    page.wait_for_timeout(800)
    ensure_not_captcha(page)


def _set_select_and_verify(page, css_id: str, value: str):
    """Set dropdown value and verify it was set correctly."""
    page.locator(css_id).select_option(value=value)
    page.wait_for_timeout(200)
    
    actual = page.locator(css_id).evaluate("el => el.value")
    if str(actual) != str(value):
        raise RuntimeError(f"Dropdown {css_id} did not change. Expected {value}, got {actual}.")


def wait_for_loading_indicators(page):
    """Wait for common loading indicators to disappear."""
    try:
        page.wait_for_selector('.loading, .spinner, .overlay', state='hidden', timeout=3000)
    except PlaywrightTimeoutError:
        pass


def set_date_range_and_search(page, start_dt: datetime, end_dt: datetime):
    print(f"\n[DATE RANGE] Setting dates: {start_dt.strftime('%Y-%m-%d')} to {end_dt.strftime('%Y-%m-%d')}")
    
    page.goto(cs.STATS_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(800)
    ensure_not_captcha(page)

    page.wait_for_selector("#start_day", state="visible")
    page.wait_for_selector("#start_month", state="visible")
    page.wait_for_selector("#start_year", state="visible")
    page.wait_for_selector("#end_day", state="visible")
    page.wait_for_selector("#end_month", state="visible")
    page.wait_for_selector("#end_year", state="visible")

    page.wait_for_timeout(500)

    sd, sm, sy = str(start_dt.day), str(start_dt.month), str(start_dt.year)
    ed, em, ey = str(end_dt.day), str(end_dt.month), str(end_dt.year)

    print(f"[DATE RANGE] Setting start date: {sy}-{sm}-{sd}")
    _set_select_and_verify(page, "#start_year", sy)
    page.wait_for_timeout(300)
    
    _set_select_and_verify(page, "#start_month", sm)
    page.wait_for_timeout(300)
    
    _set_select_and_verify(page, "#start_day", sd)
    page.wait_for_timeout(300)

    print(f"[DATE RANGE] Setting end date: {ey}-{em}-{ed}")
    _set_select_and_verify(page, "#end_year", ey)
    page.wait_for_timeout(300)
    
    _set_select_and_verify(page, "#end_month", em)
    page.wait_for_timeout(300)
    
    _set_select_and_verify(page, "#end_day", ed)
    page.wait_for_timeout(300)

    wait_for_loading_indicators(page)
    page.wait_for_timeout(800)

    print("[DATE RANGE] Verifying dates before search...")
    start_year_val = page.locator("#start_year").evaluate("el => el.value")
    start_month_val = page.locator("#start_month").evaluate("el => el.value")
    start_day_val = page.locator("#start_day").evaluate("el => el.value")
    end_year_val = page.locator("#end_year").evaluate("el => el.value")
    end_month_val = page.locator("#end_month").evaluate("el => el.value")
    end_day_val = page.locator("#end_day").evaluate("el => el.value")
    
    print(f"[DATE RANGE] Verified start: {start_year_val}-{start_month_val}-{start_day_val}")
    print(f"[DATE RANGE] Verified end: {end_year_val}-{end_month_val}-{end_day_val}")

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

    page.wait_for_timeout(1000)

    try:
        page.wait_for_load_state("networkidle", timeout=cs.NAV_TIMEOUT_MS)
    except PlaywrightTimeoutError:
        print("[DATE RANGE] Network idle timeout (continuing anyway)")
        pass

    page.wait_for_timeout(1500)
    
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
    page.wait_for_timeout(1000)
    
    # Find all clickable links in the "Total of Registration" column (2nd column)
    # The column header is "Total of Registration", and we need links under it
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
        raise RuntimeError(f"Found only {total_count} links, need at least 3. Cannot find 3rd row.")
    
    # Get the 3rd link (0-indexed, so index 2)
    target_link = all_total_links.nth(1)
    
    # Get the text of the link to confirm we're clicking the right one
    link_text = target_link.text_content()
    print(f"[DOWNLOAD] Found 3rd row total link with value: {link_text}")
    
    # Set up download expectation before clicking
    with page.expect_download(timeout=60000) as dl_info:
        target_link.click()
        page.wait_for_timeout(2000)
    
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

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=cs.HEADLESS)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        page.set_default_timeout(cs.NAV_TIMEOUT_MS)

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

        context.close()
        browser.close()

    return report_path


if __name__ == "__main__":
    main()