#!/usr/bin/env python3
"""
Microsoft 365 Admin Center Invoice Downloader

This script automates downloading invoices from the Microsoft 365 Admin Center.
It opens a browser, waits for you to authenticate via SSO, then downloads all
available invoices to an output folder.
"""

import os
import shutil
import time
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# Configuration
INVOICE_URL = "https://admin.cloud.microsoft/?#/billoverview/invoice-list"
OUTPUT_DIR = Path(__file__).parent / "output"
USER_DATA_DIR = Path(__file__).parent / ".browser_data"
DATE_RANGE = "Past 6 months"  # Options: "Past 3 months", "Past 6 months", "Specify date range"


def setup_directories():
    """Create output and browser data directories if they don't exist."""
    OUTPUT_DIR.mkdir(exist_ok=True)
    USER_DATA_DIR.mkdir(exist_ok=True)
    print(f"Output directory: {OUTPUT_DIR}")


def wait_for_authentication(page):
    """Wait for the user to complete SSO authentication."""
    print("\n" + "=" * 60)
    print("Please complete the SSO authentication in the browser.")
    print("The script will continue automatically once you're logged in.")
    print("=" * 60 + "\n")

    # Wait for the invoice grid to appear (indicates successful login)
    try:
        page.wait_for_selector('[aria-label="Invoice grid"]', timeout=300000)  # 5 minute timeout
        print("Authentication successful!")
    except PlaywrightTimeoutError:
        print("Authentication timeout. Please try again.")
        raise


def change_date_range(page, date_range: str):
    """Change the date range filter to show more invoices."""
    try:
        # Click on the date range menu
        date_menu = page.get_by_role("menuitem", name="Past 3 months")
        if date_menu.is_visible():
            date_menu.click()
            time.sleep(0.5)

            # Select the desired date range
            page.get_by_role("menuitemcheckbox", name=date_range).click()
            time.sleep(2)  # Wait for the grid to reload
            print(f"Date range changed to: {date_range}")
    except Exception as e:
        print(f"Could not change date range (may already be set): {e}")


def close_dialogs(page):
    """Close any popup dialogs that may appear."""
    try:
        close_btn = page.get_by_role("button", name="Close this coachmark")
        if close_btn.is_visible(timeout=2000):
            close_btn.click()
            time.sleep(0.5)
    except:
        pass  # No dialog to close


def get_invoice_count(page):
    """Count the number of invoices in the grid."""
    rows = page.locator('[aria-label="Invoice grid"] [role="row"]').all()
    # Subtract 1 for the header row
    return max(0, len(rows) - 1)


def download_all_invoices(page, download_path: Path):
    """Click all download buttons and save the invoices."""
    # Find all download buttons
    download_buttons = page.get_by_role("button", name="Download invoice").all()
    total = len(download_buttons)

    if total == 0:
        print("No invoices found to download.")
        return 0

    print(f"Found {total} invoices to download.")
    downloaded = 0

    for i, button in enumerate(download_buttons, 1):
        try:
            # Get the invoice ID from the row for naming
            row = button.locator("xpath=ancestor::*[@role='row']")
            row_header = row.locator("[role='rowheader']")
            invoice_id = row_header.inner_text().split()[0] if row_header.count() > 0 else f"invoice_{i}"

            print(f"[{i}/{total}] Downloading invoice {invoice_id}...")

            # Start waiting for download before clicking
            with page.expect_download(timeout=60000) as download_info:
                button.click()

            download = download_info.value

            # Save with original filename or invoice ID
            original_name = download.suggested_filename
            if original_name:
                save_path = download_path / original_name
            else:
                save_path = download_path / f"{invoice_id}.pdf"

            # Handle duplicate filenames
            if save_path.exists():
                base = save_path.stem
                ext = save_path.suffix
                counter = 1
                while save_path.exists():
                    save_path = download_path / f"{base}_{counter}{ext}"
                    counter += 1

            download.save_as(save_path)
            print(f"    Saved: {save_path.name}")
            downloaded += 1

            # Small delay between downloads to avoid overwhelming the server
            time.sleep(1)

        except PlaywrightTimeoutError:
            print(f"    Timeout waiting for download. Skipping...")
        except Exception as e:
            print(f"    Error: {e}")

    return downloaded


def main():
    """Main entry point."""
    print("Microsoft 365 Invoice Downloader")
    print("-" * 40)

    setup_directories()

    with sync_playwright() as p:
        # Launch browser with persistent context to save authentication
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(USER_DATA_DIR),
            headless=False,  # Must be visible for SSO login
            accept_downloads=True,
            viewport={"width": 1280, "height": 900},
        )

        page = context.new_page()

        try:
            print(f"Navigating to {INVOICE_URL}")
            page.goto(INVOICE_URL, wait_until="networkidle")

            # Check if we need to authenticate
            if "login.microsoftonline.com" in page.url or "login.microsoft.com" in page.url:
                wait_for_authentication(page)

            # Wait for the page to fully load
            page.wait_for_load_state("networkidle")
            time.sleep(2)

            # Close any popup dialogs
            close_dialogs(page)

            # Change date range to get more invoices
            if DATE_RANGE != "Past 3 months":
                change_date_range(page, DATE_RANGE)

            # Download all invoices
            downloaded = download_all_invoices(page, OUTPUT_DIR)

            print("\n" + "=" * 40)
            print(f"Download complete! {downloaded} invoices saved to {OUTPUT_DIR}")
            print("=" * 40)

        except KeyboardInterrupt:
            print("\nCancelled by user.")
        except Exception as e:
            print(f"\nError: {e}")
            raise
        finally:
            # Keep browser open for a moment to see results
            input("\nPress Enter to close the browser...")
            context.close()


if __name__ == "__main__":
    main()
