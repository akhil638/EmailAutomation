import pandas as pd
import win32com.client as win32
import time
import random
import os
import logging

# --- Configuration ---
EXCEL_FILE_PATH = r'C:\Users\z004m1jf\OneDrive - Siemens AG\Documents\4. Archives\Akhil Verma\Outreach\Medical Devices\Medical Devices Masters Test.xlsx' # <-- User's path
SHEET_NAME = 'Cleaned Source' # <-- User's sheet name
TEMPLATE_FILE_PATH = r'C:\Users\z004m1jf\OneDrive - Siemens AG\Documents\4. Archives\Akhil Verma\Outreach\Medical Devices\FirstOutreachPLM.oft' # <-- User's template path

# Columns to read from Excel
EMAIL_COLUMN = 'EmailAddress'
SUBJECT_COLUMN = 'Subject'
NAME_COLUMN = 'FirstName'      # For [FirstName] placeholder
COMPANY_COLUMN = 'CompanyNameSimplified' # For [CompanyNameSimplified] placeholder
SALUTATION_COLUMN = 'Salutation' # For [Salutation] placeholder

# Configure placeholders in your template and the corresponding Excel columns
PLACEHOLDERS = {
    '[FirstName]': NAME_COLUMN,
    '[CompanyNameSimplified]': COMPANY_COLUMN,
    '[Salutation]': SALUTATION_COLUMN,
}

# Status column configuration
STATUS_COLUMN = 'SendStatus' # Column to track sending status

# Delay configuration
MIN_DELAY_MINUTES = 3
MAX_DELAY_MINUTES = 4
MIN_DELAY_SECONDS = MIN_DELAY_MINUTES * 60
MAX_DELAY_SECONDS = MAX_DELAY_MINUTES * 60

# --- NEW: Outlook Folder Configuration for Sent Items ---
# Set to None or an empty string if you DON'T want to move sent emails.
#TARGET_OUTLOOK_FOLDER_NAME = "Automated Outreach Sent" # <-- IMPORTANT: Update to your folder name or set to None

# Setup logging
#logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s') 

# --- Functions ---

def send_outlook_email_from_template(template_path, recipient_email, subject, placeholder_data):
    """
    Creates and sends an email from an Outlook Template.
    """
    if not os.path.exists(template_path):
        logging.error(f"Outlook template file not found at: {template_path}")
        return False
    mail = None
    try:
        outlook_app = win32.Dispatch('Outlook.Application')
        mail = outlook_app.CreateItemFromTemplate(template_path)
        logging.debug(f"Created mail item from template: {template_path}")

        mail.To = recipient_email
        mail.Subject = subject
        logging.debug(f"Set To: {recipient_email}, Subject: {subject}")

        try:
            body_content = None; is_html = False; body_type_used = "None"
            if hasattr(mail, 'HTMLBody') and mail.HTMLBody:
                body_content = mail.HTMLBody; is_html = True; body_type_used = "HTMLBody"
            elif hasattr(mail, 'Body') and mail.Body:
                body_content = mail.Body; body_type_used = "Body"
            
            if body_content:
                logging.debug(f"Using {body_type_used} for replacements for {recipient_email}.")
                replacements_made = 0
                for placeholder, value in placeholder_data.items():
                    replacement_value = str(value) if value is not None else ''
                    if placeholder in body_content:
                        body_content = body_content.replace(placeholder, replacement_value)
                        replacements_made += 1; logging.debug(f"   -> Replaced '{placeholder}'.")
                    else:
                        logging.debug(f"   -> Placeholder '{placeholder}' not found in {body_type_used}.")
                logging.info(f"Placeholders processed for {recipient_email}. Replacements: {replacements_made}/{len(placeholder_data)}.")
                if is_html: mail.HTMLBody = body_content
                else: mail.Body = body_content
            else:
                logging.warning(f"Template for {recipient_email} has empty body.")
        except Exception as e:
            logging.error(f"Error during placeholder replacement for {recipient_email}: {e}", exc_info=True)

        mail.Send()
        logging.info(f"Successfully sent email to: {recipient_email} with subject: '{subject}'")
        return True

    except Exception as e:
        logging.error(f"Failed to create or send email for {recipient_email}: {e}", exc_info=True)
        return False

def process_leads(excel_file_path, sheet_name, template_path):
    """
    Reads leads from Excel, sends emails, and updates status on the specified sheet
    without affecting other sheets in the Excel file.
    """
    logging.info(f"--- Starting lead processing for sheet: {sheet_name} ---")
    if not os.path.exists(excel_file_path):
        logging.error(f"Excel file not found at: {excel_file_path}")
        return False

    try:
        df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
        logging.info(f"Successfully loaded sheet '{sheet_name}' from Excel file: {excel_file_path}")

        required_columns = [EMAIL_COLUMN, SUBJECT_COLUMN]
        for excel_col in PLACEHOLDERS.values():
            if excel_col and isinstance(excel_col, str) and excel_col not in required_columns:
                required_columns.append(excel_col)

        missing_cols = [col for col in required_columns if col not in df.columns]
        if missing_cols:
            logging.error(f"Missing required columns in sheet '{sheet_name}': {', '.join(missing_cols)}. Needed: {required_columns}.")
            return False

        if STATUS_COLUMN not in df.columns:
            logging.error(f"'{STATUS_COLUMN}' column is missing in sheet '{sheet_name}'. Please add this column before running the script.")
            return False

    except FileNotFoundError:
        logging.error(f"Excel file not found: {excel_file_path}")
        return False
    except PermissionError:
         logging.error(f"PERMISSION ERROR reading Excel. CLOSE file '{excel_file_path}'.")
         return False
    except ValueError as ve:
        if f"Worksheet named '{sheet_name}' not found" in str(ve): # More specific check
            logging.error(f"Sheet named '{sheet_name}' not found in Excel file '{excel_file_path}'.")
        else:
            logging.error(f"Error reading Excel (ValueError): {ve}", exc_info=True)
        return False
    except Exception as e:
        logging.error(f"Error reading or preparing Excel file: {e}", exc_info=True)
        return False

    leads_to_send = df[df[STATUS_COLUMN].isna() | (df[STATUS_COLUMN] == '')].copy()

    if leads_to_send.empty:
        logging.info(f"No new leads that has empty SendStatus field found on sheet '{sheet_name}'.")
        return True

    logging.info(f"Found {len(leads_to_send)} leads to process on sheet '{sheet_name}'.")

    for index, row in leads_to_send.iterrows():
        email_address = row[EMAIL_COLUMN]
        subject = row[SUBJECT_COLUMN]
        # Use df.index.get_loc(index) to get integer position for original df, then +2 for Excel row
        excel_row_display = df.index.get_loc(index) + 2 

        if not isinstance(email_address, str) or '@' not in email_address:
            logging.warning(f"Skipping Excel row {excel_row_display}: Invalid email '{email_address}'.")
            df.loc[index, STATUS_COLUMN] = 'Failed - Invalid Email'
            # Save status update for this row
            update_status_and_save(df, index, 'Failed - Invalid Email', EXCEL_FILE_PATH, SHEET_NAME)
            continue

        if pd.isna(subject) or not str(subject).strip():
            logging.warning(f"Skipping Excel row {excel_row_display}: Missing subject for '{email_address}'.")
            df.loc[index, STATUS_COLUMN] = 'Failed - Missing Subject'
            update_status_and_save(df, index, 'Failed - Missing Subject', EXCEL_FILE_PATH, SHEET_NAME)
            continue
        subject = str(subject)

        placeholder_values = {}
        for placeholder, excel_col_name in PLACEHOLDERS.items():
            if excel_col_name and isinstance(excel_col_name, str) and excel_col_name in df.columns:
                cell_value = row[excel_col_name]
                actual_value = '' if pd.isna(cell_value) else cell_value
                placeholder_values[placeholder] = actual_value
            else:
                placeholder_values[placeholder] = ''
                logging.warning(f"Placeholder '{placeholder}' (col: '{excel_col_name}') data missing for row {excel_row_display}.")

        logging.info(f"Processing lead from Excel row {excel_row_display}: {email_address} (Subject: '{subject}')")

        if send_outlook_email_from_template(template_path, email_address, subject, placeholder_values):
            df.loc[index, STATUS_COLUMN] = 'Sent'
            update_status_and_save(df, index, 'Sent', EXCEL_FILE_PATH, SHEET_NAME)

            delay = random.uniform(MIN_DELAY_SECONDS, MAX_DELAY_SECONDS)
            logging.info(f"Waiting for {delay / 60:.2f} minutes before next email...")
            time.sleep(delay)
        else:
            df.loc[index, STATUS_COLUMN] = 'Failed - Send Error'
            update_status_and_save(df, index, 'Failed - Send Error', EXCEL_FILE_PATH, SHEET_NAME)

    logging.info(f"Finished processing leads on sheet '{sheet_name}'.")
    return True

def update_status_and_save(df, index, status, excel_file_path, sheet_name):
    df.loc[index, STATUS_COLUMN] = status
    try:
        with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    except Exception as e:
        logging.error(f"Error saving status '{status}' for row {index}: {e}")

# --- Main Execution ---
if __name__ == "__main__":
    print("Starting Outlook Email Automation Script (Template Mode - Single Run with Safe Excel Update & Move)...")
    print(f"--- Reading leads from: {EXCEL_FILE_PATH} (Sheet: {SHEET_NAME})")
    print(f"--- Using Outlook template: {TEMPLATE_FILE_PATH}")
    print(f"--- Configured Placeholders: {list(PLACEHOLDERS.keys())}")
    print(f"--- Delay between emails: {MIN_DELAY_MINUTES}-{MAX_DELAY_MINUTES} minutes.")
    print("--- Press Ctrl+C to stop gracefully. ---")
    print("--- IMPORTANT: Ensure Excel file is CLOSED before running! ---")

    if not os.path.exists(TEMPLATE_FILE_PATH):
         print(f"\nCRITICAL ERROR: Template file '{TEMPLATE_FILE_PATH}' not found. Script aborted.")
         logging.critical(f"Template file not found. Script aborted.")
    else:

        try:
            success = process_leads(EXCEL_FILE_PATH, SHEET_NAME, TEMPLATE_FILE_PATH)
            if success:
                print("Lead processing completed.")
            else:
                print("Lead processing encountered critical errors. Please check logs.")
        except KeyboardInterrupt:
            print("\nScript stopped by user.")
        except Exception as e:
            logging.critical(f"Unexpected critical error in main execution: {e}", exc_info=True)
            print(f"\nCRITICAL ERROR: {e}. Check logs.")
        finally:
            print("Script finished.")