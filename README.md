
# GST Bot

GST Bot is a Python Selenium automation script for logging in to the GST portal and downloading available GSTR-2B Excel files for multiple users.

This repository contains only sample/demo credential data. Do not commit real GST usernames, passwords, downloaded GST files, or generated reports.

## Features

- Reads users from an Excel file.
- Opens the GST login page with Selenium.
- Allows manual CAPTCHA entry.
- Navigates to the Returns Dashboard.
- Downloads available GSTR-2B Excel files.
- Creates a final Excel report with user-wise status.

## Tech Stack

- Python
- Selenium
- pandas
- openpyxl
- webdriver-manager

## Setup

Install dependencies:

```bash
pip install -r requirements.txt
```

## Usage

Update `sample_users.xlsx` with sample or test credentials only, then run:

```bash
python gst_bot.py
```

CAPTCHA must be entered manually in the browser when prompted.

## Output

Downloaded files are saved in:

```text
GST_Downloads/
```

Summary reports are saved in:

```text
Reports/
```

Both folders are ignored by Git and should not be uploaded to GitHub.

## Security Notes

- Do not commit real GST login credentials.
- Do not commit downloaded GST files.
- Do not commit generated reports.
- Use `sample_users.xlsx` only for dummy/demo data.
