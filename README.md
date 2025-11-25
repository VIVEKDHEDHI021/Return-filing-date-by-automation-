GST Filing Data Collector – README
==Overview

GST Filing Data Collector is a desktop-based Python application designed to automate the retrieval of GST return filing information directly from the official GST Portal (https://services.gst.gov.in
).
It collects the latest filing dates for:

GSTR-1 (Outward Supplies)

GSTR-3B (Monthly Summary Return)

The tool allows manual CAPTCHA entry while automating all other steps using Selenium.
Final results are saved in a clean and professionally formatted Excel file (gst_filing_data.xlsx).

1. Automated Filing Information Fetching

Automatically navigates to the GST portal

Extracts only the latest filing dates for GSTR-1 and GSTR-3B

Supports multiple GSTINs at once

2. Manual CAPTCHA Handling

User enters CAPTCHA manually (for legal compliance)

After verification, automation continues automatically

3. Automatic Page Actions

After CAPTCHA and search:

Page scrolling

Clicking Show Filing Table

Clicking filing section Search

Loading return tables

Extracting latest filing dates
