# Outlook Email Automation Script

This Python script automates sending personalized emails using Microsoft Outlook and an Excel file as the source of recipient data. It uses an Outlook email template (`.oft` file), replaces placeholders with values from Excel, sends the emails, and updates the Excel file with the send status for each lead.

## Features

- Reads recipient data from an Excel sheet.
- Uses an Outlook `.oft` template for email formatting.
- Replaces placeholders (e.g., `[FirstName]`, `[CompanyNameSimplified]`) in the template with values from Excel.
- Sends emails via Outlook (using the desktop client).
- Updates the Excel file with the status of each email (Sent, Failed, etc.).
- Adds a random delay between emails to avoid spam filters.
- Detailed logging for troubleshooting and audit.

## Requirements

- **Windows OS** (script uses `win32com` for Outlook automation)
- **Microsoft Outlook** (desktop version, must be installed and configured)
- **Python 3.7+**
- Python packages:
  - `pandas`
  - `openpyxl`
  - `pywin32`

Install dependencies with:
```bash
pip install pandas openpyxl pywin32
```

## Setup

1. **Clone this repository** or copy the script to your local machine.

2. **Prepare your Excel file**:
   - Ensure your Excel file contains the following columns (update names in the script if needed):
     - `EmailAddress`
     - `Subject`
     - `FirstName`
     - `CompanyNameSimplified`
     - `Salutation`
     - `SendStatus` (used to track which leads have been processed)
   - Place your Excel file at the path specified in `EXCEL_FILE_PATH` in the script.

3. **Prepare your Outlook template**:
   - Create an `.oft` template in Outlook with placeholders like `[FirstName]`, `[CompanyNameSimplified]`, etc.
   - Save the template and update the `TEMPLATE_FILE_PATH` in the script.

4. **Configure the script**:
   - Update the following variables at the top of the script as needed:
     - `EXCEL_FILE_PATH`
     - `SHEET_NAME`
     - `TEMPLATE_FILE_PATH`
     - `PLACEHOLDERS` dictionary if you use different placeholders/columns.

## Usage

1. **Close your Excel file** before running the script (the script needs exclusive access).
2. Run the script:
   ```bash
   python SendEmailTest.py
   ```
3. The script will:
   - Read the leads from the specified Excel sheet.
   - For each lead with an empty `SendStatus`, send a personalized email.
   - Update the `SendStatus` column in Excel after each attempt.
   - Wait a random interval (3–4 minutes by default) between emails.

4. **Interrupt the script** at any time with `Ctrl+C`. It will finish the current email and exit gracefully.

## Logging

- The script logs detailed information and errors to the console.
- You can adjust the logging level in the script (`logging.basicConfig(...)`).

## Troubleshooting

- **Excel file not found**: Check the `EXCEL_FILE_PATH`.
- **Template file not found**: Check the `TEMPLATE_FILE_PATH`.
- **Permission errors**: Make sure the Excel file is closed before running the script.
- **Missing columns**: Ensure all required columns exist in your Excel sheet.

## Notes

- The script is designed for safe updates: it only modifies the specified sheet and does not affect other sheets in the Excel file.
- You can customize the delay between emails by changing `MIN_DELAY_MINUTES` and `MAX_DELAY_MINUTES`.
- The script does not move sent emails to a specific Outlook folder by default. You can add this feature if needed.

## License

This project is provided as-is for internal or personal use. Please review and adapt for your organization's policies and requirements.

---

**Author:** Akhil Verma  
**Contact:** [Your Email or GitHub profile]

---

**If you want this as a file:**  
- Copy all the text above and paste it into a file named `README.md` in your repository.

Let me know if you want a shorter or more detailed version, or if you want to include example screenshots or sample Excel/template files!
