#!/usr/bin/env python3
"""
gst_portal_bot.py
==================
Engine behind the "GST Returns & Notices — Bulk Download" Excel/VBA tool.

Called as:   python gst_portal_bot.py <job_file.txt>

The job file is a simple KEY=VALUE text file written by the Excel macro, with:
    GSTIN, LOGIN_ID, PASSWORD, FY, ACTION, PARTY_NAME, DOWNLOAD_FOLDER
ACTION is one of: GSTR1, GSTR3B, BOTH, NOTICES

Writes a result file at <job_file.txt>.result.txt:
    line 1: STATUS   (SUCCESS / PARTIAL / FAILED)
    line 2: MESSAGE  (one-line human-readable summary)
    line 3+: one downloaded file path per line (if any)

----------------------------------------------------------------------------
IMPORTANT — WHAT IS VERIFIED VS. BEST-EFFORT IN THIS SCRIPT
----------------------------------------------------------------------------
This script was built after a live session actually browsing
services.gst.gov.in / return.gst.gov.in. The following were directly
observed and are reflected accurately below:
  - The portal is an Angular single-page app. Navigating by typing a URL
    directly (driver.get on an internal page) invalidates the session and
    throws "Access Denied — session expired". Every internal navigation
    in this script is therefore done by clicking the actual on-page link/
    button, never by driver.get() after the first login page load.
  - The "View Filed Returns" screen (Services > Returns > View Filed
    Returns) has a Financial Year select, a Return Filing Period select
    (Annual/Half Yearly/Quarterly/Monthly), a Month select, a Return Type
    select (options include "GSTR3B", "GSTR-1/IFF/GSTR-1A", etc.) and a
    SEARCH button, matching the selectors used below.
  - The "Track Return Status" and "Notices and Orders" screens, their
    column layout, and the fact that a notice's "View" link opens in a
    SEPARATE browser window/tab (not the same tab) were directly observed.

The following were NOT observed live (the only account tested had zero
filed returns, so no "download a filed return" screen was ever actually
rendered) and are best-effort based on the portal's general conventions.
TREAT THESE AS THE FIRST THINGS TO VERIFY AND FIX on your own GSTIN once
it has real filed returns:
  - The exact button/icon used to download a filed GSTR-1 / GSTR-3B as a
    PDF from the results table (DOWNLOAD_RETURN_PDF_XPATHS below).
  - The exact login page field identifiers (LOGIN_USERNAME_XPATHS /
    LOGIN_PASSWORD_XPATHS / LOGIN_BUTTON_XPATHS below) — these use
    several fallback strategies, but confirm on first run.

All such selectors are centralised at the top of this file so you (or
whoever maintains this) can fix one place when the portal's HTML changes.
----------------------------------------------------------------------------
"""

import os
import re
import sys
import time
import glob
from datetime import datetime

try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.support.ui import WebDriverWait, Select
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import (
        TimeoutException, NoSuchElementException, WebDriverException
    )
except ImportError:
    print("SELENIUM NOT INSTALLED. Run: pip install -r requirements.txt")
    sys.exit(1)


# ============================================================================
# CONFIG — centralised selectors. Update here if the portal's HTML changes.
# ============================================================================

LOGIN_URL = "https://services.gst.gov.in/services/login"

LOGIN_USERNAME_XPATHS = [
    "//input[@id='username']",
    "//input[contains(@formcontrolname,'user')]",
    "//label[contains(normalize-space(.),'Username')]/following::input[1]",
]
LOGIN_PASSWORD_XPATHS = [
    "//input[@id='user_pass']",
    "//input[@type='password']",
    "//label[contains(normalize-space(.),'Password')]/following::input[1]",
]
LOGIN_BUTTON_XPATHS = [
    "//button[normalize-space()='LOGIN']",
    "//button[contains(translate(normalize-space(.), 'login', 'LOGIN'),'LOGIN')]",
]
# Something on the post-login page that proves we're in (dashboard / welcome banner / user name chip)
LOGGED_IN_MARKERS = [
    "//a[contains(@href,'/services/auth/')]",
    "//*[contains(normalize-space(.),'Last logged in')]",
]

SERVICES_MENU_XPATH = "//a[normalize-space()='Services'] | //button[normalize-space()='Services']"
RETURNS_SUBMENU_XPATH = "//a[normalize-space()='Returns']"
VIEW_FILED_RETURNS_LINK_XPATH = "//a[normalize-space()='View Filed Returns']"
NOTICES_QUICKLINK_XPATH = "//a[contains(normalize-space(.),'Notices and Orders')]"

FY_SELECT_XPATH = "//label[contains(.,'Financial year') or contains(.,'Financial Year')]/following::select[1]"
RETURN_FILING_PERIOD_SELECT_XPATH = "//label[contains(.,'Return Filing Period')]/following::select[1]"
RETURN_TYPE_SELECT_XPATH = "//label[contains(.,'Return Type')]/following::select[1]"
SEARCH_BUTTON_XPATH = "//button[normalize-space()='SEARCH' or normalize-space()='Search']"

# Best-effort — not verified live (see module docstring)
DOWNLOAD_RETURN_PDF_XPATHS = [
    "//a[contains(@class,'download') or contains(normalize-space(.),'Download')]",
    "//button[contains(normalize-space(.),'Download')]",
    "//a[contains(normalize-space(.),'View')]",
]

NOTICES_TABLE_ROW_XPATH = "//table//tr[td]"
NOTICE_VIEW_LINK_XPATH = ".//a[normalize-space()='View']"

GSTR_TYPE_OPTION_TEXT = {
    "GSTR1": "GSTR-1",     # matched with "contains" against select option text
    "GSTR3B": "GSTR3B",
}

WAIT_SECS = 20


# ============================================================================
# Helpers
# ============================================================================

def read_job(path):
    job = {}
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.rstrip("\n")
            if not line or "=" not in line:
                continue
            k, v = line.split("=", 1)
            job[k.strip()] = v.strip()
    required = ["GSTIN", "LOGIN_ID", "PASSWORD", "FY", "ACTION"]
    missing = [k for k in required if not job.get(k)]
    if missing:
        raise ValueError(f"Job file missing required fields: {missing}")
    return job


def write_result(result_path, status, message, files=None):
    files = files or []
    with open(result_path, "w", encoding="utf-8") as f:
        f.write(status + "\n")
        f.write(message.replace("\n", " ") + "\n")
        for fp in files:
            f.write(fp + "\n")


def safe_name(s):
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", s).strip("_")


def find_first(driver, xpaths, timeout=WAIT_SECS):
    """Try a list of XPaths in order, return the first visible element found."""
    last_err = None
    for xp in xpaths:
        try:
            el = WebDriverWait(driver, timeout).until(
                EC.visibility_of_element_located((By.XPATH, xp))
            )
            return el
        except TimeoutException as e:
            last_err = e
            continue
    raise NoSuchElementException(f"None of the selectors matched: {xpaths}") from last_err


def click_first(driver, xpaths, timeout=WAIT_SECS):
    el = find_first(driver, xpaths, timeout)
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    try:
        el.click()
    except WebDriverException:
        driver.execute_script("arguments[0].click();", el)
    return el


def select_by_visible_text_fragment(select_element, text_fragment):
    """Selects the first <option> whose visible text contains text_fragment (case-insensitive)."""
    sel = Select(select_element)
    for opt in sel.options:
        if text_fragment.lower() in opt.text.strip().lower():
            sel.select_by_visible_text(opt.text)
            return opt.text
    raise NoSuchElementException(f"No option containing '{text_fragment}' in this <select>")


def new_chrome_driver(download_folder):
    os.makedirs(download_folder, exist_ok=True)
    options = Options()
    options.add_argument("--start-maximized")
    # Never run headless: a human must see the page to type the CAPTCHA.
    prefs = {
        "download.default_directory": download_folder,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True,  # force PDFs to download, not open in-browser viewer
        "profile.default_content_settings.popups": 0,
    }
    options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(options=options)
    return driver


def wait_for_manual_login(driver):
    """
    Pauses here. The CAPTCHA cannot and must not be automated — the human
    operator types it into the real browser window and clicks the portal's
    own Login button, then confirms here.
    """
    print("\n" + "=" * 78)
    print(" ACTION NEEDED: a Chrome window is open on the GST login page.")
    print(" Username and Password have been filled in for you.")
    print(" >>> Please type the CAPTCHA shown in the browser and click LOGIN. <<<")
    print(" Once you see the GST Dashboard / Welcome page, come back here and")
    print(" press ENTER to let the automation continue.")
    print("=" * 78)
    input(" Press ENTER once you are logged in... ")

    # Confirm we actually reached a logged-in page before proceeding.
    try:
        find_first(driver, LOGGED_IN_MARKERS, timeout=10)
        return True
    except NoSuchElementException:
        print(" Could not confirm a logged-in page yet — waiting a little longer...")
        try:
            find_first(driver, LOGGED_IN_MARKERS, timeout=20)
            return True
        except NoSuchElementException:
            return False


def navigate_to_view_filed_returns(driver):
    """Services > Returns > View Filed Returns — all by clicking, never driver.get()."""
    click_first(driver, [SERVICES_MENU_XPATH])
    click_first(driver, [RETURNS_SUBMENU_XPATH])
    click_first(driver, [VIEW_FILED_RETURNS_LINK_XPATH])


def download_return(driver, job, return_type_key, downloaded_files, errors):
    """
    return_type_key: "GSTR1" or "GSTR3B"
    Searches the given Financial Year for that return type and attempts to
    download every filed period found. See module docstring: the actual
    download step is best-effort / unverified and the first thing to check.
    """
    try:
        navigate_to_view_filed_returns(driver)

        fy_select = find_first(driver, [FY_SELECT_XPATH])
        select_by_visible_text_fragment(fy_select, job["FY"])

        # Return Type
        try:
            rtype_select = find_first(driver, [RETURN_TYPE_SELECT_XPATH], timeout=10)
            select_by_visible_text_fragment(rtype_select, GSTR_TYPE_OPTION_TEXT[return_type_key])
        except NoSuchElementException:
            pass  # some portal versions show Return Type only after Return Filing Period is chosen

        # Return Filing Period -> Monthly (GSTR-1 and GSTR-3B are filed monthly/quarterly;
        # defaulting to Monthly here — adjust if a client is on the QRMP quarterly scheme)
        try:
            period_select = find_first(driver, [RETURN_FILING_PERIOD_SELECT_XPATH], timeout=10)
            select_by_visible_text_fragment(period_select, "Monthly")
        except NoSuchElementException:
            pass

        click_first(driver, [SEARCH_BUTTON_XPATH])
        time.sleep(2)

        rows = driver.find_elements(By.XPATH, NOTICES_TABLE_ROW_XPATH)
        if not rows:
            errors.append(f"{return_type_key} ({job['FY']}): no filed returns found / no results table.")
            return

        before_files = set(glob.glob(os.path.join(job["_download_folder"], "*")))

        for idx in range(len(rows)):
            # re-find rows each loop: DOM can refresh after a click
            rows = driver.find_elements(By.XPATH, NOTICES_TABLE_ROW_XPATH)
            if idx >= len(rows):
                break
            row = rows[idx]
            try:
                link = find_first(driver, DOWNLOAD_RETURN_PDF_XPATHS, timeout=5) \
                    if idx == 0 else row.find_element(By.XPATH, ".//a | .//button")
                handles_before = driver.window_handles
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", link)
                try:
                    link.click()
                except WebDriverException:
                    driver.execute_script("arguments[0].click();", link)
                time.sleep(2)

                # If it opened a new tab/window (as Notices "View" does), try to print it to PDF and close it.
                handles_after = driver.window_handles
                new_handles = [h for h in handles_after if h not in handles_before]
                if new_handles:
                    driver.switch_to.window(new_handles[0])
                    time.sleep(1)
                    save_current_page_as_pdf(
                        driver, job["_download_folder"],
                        f"{safe_name(job['PARTY_NAME'] or job['GSTIN'])}_{return_type_key}_{job['FY']}_{idx+1}.pdf"
                    )
                    driver.close()
                    driver.switch_to.window(handles_before[0])
            except Exception as row_err:  # noqa: BLE001 - log and continue with next row
                errors.append(f"{return_type_key} row {idx+1}: {row_err}")

        after_files = set(glob.glob(os.path.join(job["_download_folder"], "*")))
        downloaded_files.extend(sorted(after_files - before_files))

    except Exception as e:  # noqa: BLE001
        errors.append(f"{return_type_key}: {e}")


def save_current_page_as_pdf(driver, download_folder, filename):
    """Uses Chrome DevTools Protocol to print the current tab to a PDF file."""
    result = driver.execute_cdp_cmd("Page.printToPDF", {"printBackground": True})
    import base64
    data = base64.b64decode(result["data"])
    out_path = os.path.join(download_folder, filename)
    with open(out_path, "wb") as f:
        f.write(data)
    return out_path


def download_notices(driver, job, downloaded_files, errors):
    try:
        click_first(driver, [SERVICES_MENU_XPATH])
        click_first(driver, ["//a[normalize-space()='User Services']"])
        click_first(driver, [NOTICES_QUICKLINK_XPATH])
        time.sleep(2)

        rows = driver.find_elements(By.XPATH, NOTICES_TABLE_ROW_XPATH)
        if not rows:
            errors.append("Notices: no rows found in the Notices and Orders table.")
            return

        n_rows = len(rows)
        for idx in range(n_rows):
            rows = driver.find_elements(By.XPATH, NOTICES_TABLE_ROW_XPATH)
            if idx >= len(rows):
                break
            row = rows[idx]
            try:
                row_text_cells = row.find_elements(By.XPATH, "./td")
                notice_id = row_text_cells[0].text.strip() if row_text_cells else f"notice_{idx+1}"
                view_link = row.find_element(By.XPATH, NOTICE_VIEW_LINK_XPATH)

                handles_before = driver.window_handles
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", view_link)
                try:
                    view_link.click()
                except WebDriverException:
                    driver.execute_script("arguments[0].click();", view_link)
                time.sleep(2)

                handles_after = driver.window_handles
                new_handles = [h for h in handles_after if h not in handles_before]
                if new_handles:
                    driver.switch_to.window(new_handles[0])
                    time.sleep(1)
                    fname = f"{safe_name(job['PARTY_NAME'] or job['GSTIN'])}_Notice_{safe_name(notice_id)}.pdf"
                    path = save_current_page_as_pdf(driver, job["_download_folder"], fname)
                    downloaded_files.append(path)
                    driver.close()
                    driver.switch_to.window(handles_before[0])
                else:
                    errors.append(f"Notice {notice_id}: View did not open a new window as expected.")
            except Exception as row_err:  # noqa: BLE001
                errors.append(f"Notice row {idx+1}: {row_err}")

    except Exception as e:  # noqa: BLE001
        errors.append(f"Notices: {e}")


def main():
    if len(sys.argv) != 2:
        print("Usage: python gst_portal_bot.py <job_file.txt>")
        sys.exit(1)

    job_path = sys.argv[1]
    result_path = job_path + ".result.txt"

    try:
        job = read_job(job_path)
    except Exception as e:  # noqa: BLE001
        write_result(result_path, "FAILED", f"Could not read job file: {e}")
        sys.exit(1)

    download_folder = job.get("DOWNLOAD_FOLDER") or os.path.join(os.path.expanduser("~"), "Downloads")
    job["_download_folder"] = download_folder

    downloaded_files = []
    errors = []
    driver = None

    try:
        driver = new_chrome_driver(download_folder)
        driver.get(LOGIN_URL)  # the ONE allowed driver.get() — first page load only

        user_field = find_first(driver, LOGIN_USERNAME_XPATHS)
        user_field.clear()
        user_field.send_keys(job["LOGIN_ID"])

        pass_field = find_first(driver, LOGIN_PASSWORD_XPATHS)
        pass_field.clear()
        pass_field.send_keys(job["PASSWORD"])

        ok = wait_for_manual_login(driver)
        if not ok:
            write_result(result_path, "FAILED", "Login was not confirmed (CAPTCHA/login step not completed).")
            return

        action = job["ACTION"].upper()
        if action in ("GSTR1", "BOTH"):
            download_return(driver, job, "GSTR1", downloaded_files, errors)
        if action in ("GSTR3B", "BOTH"):
            download_return(driver, job, "GSTR3B", downloaded_files, errors)
        if action == "NOTICES":
            download_notices(driver, job, downloaded_files, errors)

        if downloaded_files and not errors:
            write_result(result_path, "SUCCESS", f"Downloaded {len(downloaded_files)} file(s).", downloaded_files)
        elif downloaded_files and errors:
            write_result(result_path, "PARTIAL", "; ".join(errors)[:500], downloaded_files)
        else:
            write_result(result_path, "FAILED", "; ".join(errors)[:500] or "No files were downloaded.")

    except Exception as e:  # noqa: BLE001
        write_result(result_path, "FAILED", f"Unhandled error: {e}", downloaded_files)
    finally:
        if driver is not None:
            try:
                time.sleep(2)
                driver.quit()
            except Exception:  # noqa: BLE001
                pass


if __name__ == "__main__":
    main()
