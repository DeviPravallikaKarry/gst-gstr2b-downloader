
# GST Bot

A Python Selenium automation script for downloading GST GSTR-2B Excel files.

## Setup

Install dependencies:

```bash
pip install -r requirements.txt
Usage
Add usernames and passwords in sample_users.xlsx, then run:

python gst_bot.py
CAPTCHA must be entered manually in the browser.

Notes
Do not commit real GST login credentials.
Downloaded files are saved in GST_Downloads/.
Reports are saved in Reports/.