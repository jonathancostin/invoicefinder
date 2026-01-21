#!/usr/bin/env python3
"""
Microsoft 365 Admin Center Invoice Downloader

This script automates downloading invoices from the Microsoft 365 Admin Center.
It opens a browser, waits for you to authenticate via SSO, then downloads all
available invoices to an output folder.
"""

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
        # Try common date range options that might be currently selected
        date_options = ["Past 3 months", "Past 6 months", "Past 12 months", "All"]
        date_menu = None

        for option in date_options:
            menu = page.get_by_role("menuitem", name=option)
            if menu.is_visible(timeout=1000):
                date_menu = menu
                break

        if date_menu is None:
            print("Could not find date range menu. Continuing with current selection.")
            return

        date_menu.click()
        page.wait_for_timeout(500)

        # Select the desired date range
        target_option = page.get_by_role("menuitemcheckbox", name=date_range)
        if target_option.is_visible(timeout=2000):
            target_option.click()
            time.sleep(2)  # Allow grid to refresh
            print(f"Date range changed to: {date_range}")
        else:
            print(f"Date range option '{date_range}' not found. Using current selection.")

    except PlaywrightTimeoutError:
        print("Date range menu interaction timed out. Continuing with current selection.")
    except Exception as e:
        print(f"Could not change date range: {e}")


def close_dialogs(page):
    """Close any popup dialogs that may appear."""
    try:
        close_btn = page.get_by_role("button", name="Close this coachmark")
        if close_btn.is_visible(timeout=2000):
            close_btn.click()
            time.sleep(0.5)
    except PlaywrightTimeoutError:
        pass  # No dialog to close
    except Exception:
        pass  # Dialog interaction failed, continue anyway


def download_all_invoices(page, download_path: Path):
    """Click all download buttons and save the invoices."""
    # Get initial count of download buttons
    total = page.get_by_role("button", name="Download invoice").count()

    if total == 0:
        print("No invoices found to download.")
        return 0

    print(f"Found {total} invoices to download.")
    downloaded = 0

    for i in range(1, total + 1):
        try:
            # Re-query buttons each iteration to avoid stale element references
            # after DOM updates from previous downloads
            buttons = page.get_by_role("button", name="Download invoice")
            button = buttons.nth(i - 1)

            # Get the invoice ID from the row for naming
            row = button.locator("xpath=ancestor::*[@role='row']")
            row_header = row.locator("[role='rowheader']")

            # Safely extract invoice ID, handling empty text
            invoice_id = f"invoice_{i}"
            if row_header.count() > 0:
                header_text = row_header.inner_text().strip()
                if header_text:
                    parts = header_text.split()
                    if parts:
                        invoice_id = parts[0]

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

            # Verify file was saved successfully
            if not save_path.exists():
                print(f"    Warning: File may not have saved correctly: {save_path.name}")
                continue

            file_size = save_path.stat().st_size
            if file_size == 0:
                print(f"    Warning: File is empty: {save_path.name}")
                save_path.unlink()  # Remove empty file
                continue

            print(f"    Saved: {save_path.name} ({file_size:,} bytes)")
            downloaded += 1

            # Small delay between downloads to avoid overwhelming the server
            time.sleep(1)

        except PlaywrightTimeoutError:
            print(f"    Timeout waiting for download {i}. Skipping...")
        except Exception as e:
            print(f"    Error downloading invoice {i}: {e}")

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
            page.goto(INVOICE_URL, wait_until="domcontentloaded")

            # Check if we need to authenticate
            if "login.microsoftonline.com" in page.url or "login.microsoft.com" in page.url:
                wait_for_authentication(page)
            else:
                # Already authenticated - wait for invoice grid to appear
                print("Session found, waiting for invoice list to load...")
                page.wait_for_selector('[aria-label="Invoice grid"]', timeout=60000)

            # Give the page a moment to finish rendering
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
            try:
                input("\nPress Enter to close the browser...")
            except EOFError:
                pass  # Running non-interactively, just close
            context.close()


if __name__ == "__main__":
    main()
