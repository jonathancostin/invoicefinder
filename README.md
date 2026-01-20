# Microsoft 365 Invoice Downloader

A Python script that automates downloading invoices from the Microsoft 365 Admin Center billing page.

## Features

- Opens a browser and navigates to the Microsoft 365 billing portal
- Waits for manual SSO authentication (supports any SSO provider)
- Persists browser session data for faster subsequent runs
- Configurable date range filter (3 months, 6 months, or custom)
- Downloads all invoices as PDFs to a local output folder
- Handles duplicate filenames automatically

## Requirements

- Python 3.8+
- Playwright

## Installation

```bash
# Clone the repository
git clone https://github.com/YOUR_USERNAME/invoicefinder.git
cd invoicefinder

# Create virtual environment
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Install Playwright browsers (first time only)
playwright install chromium
```

## Usage

```bash
# Activate virtual environment
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Run the script
python download_invoices.py
```

On first run:
1. A browser window will open and navigate to the Microsoft 365 Admin Center
2. Complete the SSO authentication manually
3. The script will automatically detect when you're logged in
4. Invoices will be downloaded to the `output/` folder

On subsequent runs, your session may be cached, reducing the need to re-authenticate.

## Configuration

Edit the variables at the top of `download_invoices.py`:

| Variable | Default | Description |
|----------|---------|-------------|
| `INVOICE_URL` | Microsoft billing page | The URL to navigate to |
| `OUTPUT_DIR` | `./output` | Where to save downloaded invoices |
| `USER_DATA_DIR` | `./.browser_data` | Browser session data location |
| `DATE_RANGE` | `"Past 6 months"` | Filter options: `"Past 3 months"`, `"Past 6 months"`, `"Specify date range"` |

## Output

Downloaded invoices are saved to the `output/` directory with their original filenames from Microsoft (typically in the format `Invoice_GXXXXXXXXX.pdf`).

## License

MIT
