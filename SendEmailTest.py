import pandas as pd
import win32com.client as win32
import time
import random
import os
import logging
from datetime import datetime

# --- Configuration ---
EXCEL_FILE_PATH = r'C:\Users\z004m1jf\OneDrive - Siemens AG\Documents\4. Archives\Akhil Verma\Outreach\Medical Devices\Medical Devices Masters Test.xlsx' # <-- User's path
SHEET_NAME = 'Cleaned Source' # <-- User's sheet name
# Define paths for your templates
TEMPLATES = {
    0: r'C:\Users\z004m1jf\OneDrive - Siemens AG\Documents\4. Archives\Akhil Verma\Outreach\Medical Devices\FirstOutreachPLM.oft', # Path to Initial email template (Followup 0)
    1: r'C:\Users\z004m1jf\OneDrive - Siemens AG\Documents\4. Archives\Akhil Verma\Outreach\Medical Devices\Followup1.oft', # Path to Follow-up 1 template
    2: r'C:\Users\z004m1jf\OneDrive - Siemens AG\Documents\4. Archives\Akhil Verma\Outreach\Medical Devices\Followup2.oft', # Path to Follow-up 2 template
    3: r'C:\Users\z004m1jf\OneDrive - Siemens AG\Documents\4. Archives\Akhil Verma\Outreach\Medical Devices\Followup3.oft', # Path to Follow-up 3 template
}

# Define the delay in days before sending a follow-up
FOLLOWUP_DELAY_DAYS = 3

# Columns to read from Excel
EMAIL_COLUMN = 'EmailAddress'
SUBJECT_COLUMN = 'Subject'
NAME_COLUMN = 'FirstName'      # For [FirstName] placeholder
COMPANY_COLUMN = 'CompanyNameSimplified' # For [CompanyNameSimplified] placeholder
SALUTATION_COLUMN = 'Salutation' # For [Salutation] placeholder
LAST_SENT_DATE_COLUMN = 'LastSentDate'
FOLLOWUP_COLUMN = 'FollowupNumber'
REPLY_DATE_COLUMN = 'ReplyDate'

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

# Define paths for your templates
# IMPORTANT: Replace these with the actual paths to your templates
TEMPLATES = {
    0: r'C:\Users\z004m1jf\OneDrive - Siemens AG\Documents\4. Archives\Akhil Verma\Outreach\Medical Devices\FirstOutreachPLM.oft', # Path to Initial email template (Followup 0)
    1: r'C:\Users\z004m1jf\OneDrive - Siemens AG\Documents\4. Archives\Akhil Verma\Outreach\Medical Devices\Followup1.oft', # Path to Follow-up 1 template
    2: r'C:\Users\z004m1jf\OneDrive - Siemens AG\Documents\4. Archives\Akhil Verma\Outreach\Medical Devices\Followup2.oft', # Path to Follow-up 2 template
    3: r'C:\Users\z004m1jf\OneDrive - Siemens AG\Documents\4. Archives\Akhil Verma\Outreach\Medical Devices\Followup3.oft', # Path to Follow-up 3 template
}

# Define the delay in days before sending a follow-up
FOLLOWUP_DELAY_DAYS = 3

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


def check_for_reply(recipient_email):
    """
    Checks Outlook inbox for any email from the recipient email address using Restrict filter.
    """
    logging.debug(f"Starting efficient reply check for {recipient_email} using Restrict.")
    try:
        outlook_app = win32.Dispatch('Outlook.Application')
        namespace = outlook_app.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6) # 6 refers to the Inbox folder
        logging.debug("Accessed Outlook Inbox for Restrict search.")

        # Construct a filter string for SenderEmailAddress.
        # Using 'LIKE' with wildcards (%) makes the search more flexible
        # to match formats like "Display Name <email@address.com>".
        # We need to escape single quotes in the email address if any exist, though unlikely.
        # Simple check for single quotes:
        clean_recipient_email = recipient_email.replace("'", "''")

        # Filter string to find items where SenderEmailAddress contains the recipient's email
        # MAPI filter syntax uses "@SQL=" prefix for advanced queries like LIKE.
        # However, simple property filters like [Property] = 'Value' or [Property] LIKE 'Value'
        # often work directly without "@SQL=". Let's try the simpler syntax first.
        # filter_string = f"[SenderEmailAddress] LIKE '%{clean_recipient_email}%'" # Using LIKE
        
        # Exact match on just the email part might be more precise if SenderEmailAddress
        # consistently stores just the email. Let's try exact match first, it's usually faster.
        # If exact match doesn't work, we can switch to LIKE.
        filter_string = f"[SenderEmailAddress] = '{clean_recipient_email}'"
        logging.debug(f"Applying filter for replies: {filter_string}")

        # Apply the filter using Restrict
        restricted_items = inbox.Items.Restrict(filter_string)

        # Check if any items were found
        if restricted_items.Count > 0:
            # Found at least one email from the recipient
            logging.info(f"Reply detected from {recipient_email} (found {restricted_items.Count} items).")
            # Get the received time of the first found item (most recent due to default sort, or can sort again)
            # Although we don't *strictly* need the time for the simplified logic, returning it is good practice.
            # The items collection returned by Restrict is often already sorted by ReceivedTime descending by default,
            # but explicitly sorting is safer if we need the most recent one.
            # Let's get the first item's time just in case, although the prompt said we don't care about *when*.
            first_item_received_time = restricted_items.GetFirst().ReceivedTime

            return True, first_item_received_time
        else:
            # No emails from the recipient found
            logging.debug(f"No reply detected from {recipient_email} in inbox using filter.")
            return False, None

    except Exception as e:
        logging.error(f"Error checking Outlook inbox for replies for {recipient_email} using Restrict: {e}", exc_info=True)
        return False, None

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
        # Specify dtype for relevant columns. Read FollowupNumber as float initially to handle NaNs.
        dtype_spec = {
            STATUS_COLUMN: str,
            LAST_SENT_DATE_COLUMN: str,
            REPLY_DATE_COLUMN: str,
            FOLLOWUP_COLUMN: float # Read as float to handle NaNs
        }
        df = pd.read_excel(excel_file_path, sheet_name=sheet_name, dtype=dtype_spec)
        logging.info(f"Successfully loaded sheet '{sheet_name}' from Excel file: {excel_file_path}")

        # The checks below for missing columns are still needed
        # as dtype only applies if the column exists.

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

        if LAST_SENT_DATE_COLUMN not in df.columns:
            df[LAST_SENT_DATE_COLUMN] = ''
            df[LAST_SENT_DATE_COLUMN] = df[LAST_SENT_DATE_COLUMN].astype(str)

        # Add FollowupNumber column if missing, then fill NaNs and convert to int
        if FOLLOWUP_COLUMN not in df.columns:
            df[FOLLOWUP_COLUMN] = 0 # Default to 0 if missing

        # Fill any remaining NaNs in FollowupNumber with 0 and convert to int
        # This is necessary because reading as float can still result in NaNs if cells are empty
        df[FOLLOWUP_COLUMN] = df[FOLLOWUP_COLUMN].fillna(0).astype(int)

        # Add ReplyDate column if missing
        if REPLY_DATE_COLUMN not in df.columns:
            df[REPLY_DATE_COLUMN] = ''
            df[REPLY_DATE_COLUMN] = df[REPLY_DATE_COLUMN].astype(str)


    except FileNotFoundError:
        logging.error(f"Excel file not found: {excel_file_path}")
        return False
    except PermissionError:
         logging.error(f"PERMISSION ERROR reading Excel. CLOSE file '{excel_file_path}'.")
         return False
    except ValueError as ve:
        if f"Worksheet named '{sheet_name}' not found" in str(ve):
            logging.error(f"Sheet named '{sheet_name}' not found in Excel file '{excel_file_path}'.")
        else:
            logging.error(f"Error reading Excel (ValueError): {ve}", exc_info=True)
        return False
    except Exception as e:
        logging.error(f"Error reading or preparing Excel file: {e}", exc_info=True)
        return False

    # We'll process leads that are candidates for the initial email (empty status)
    # OR leads that have been sent an email but haven't replied and are due for a followup
    leads_to_process = df[
        (df[STATUS_COLUMN].isna() | (df[STATUS_COLUMN] == '')) | # Candidates for initial email
        ((df[STATUS_COLUMN] == 'Sent') & (df[FOLLOWUP_COLUMN] < len(TEMPLATES))) # Candidates for followups
    ].copy()


    if leads_to_process.empty:
        logging.info(f"No leads require processing on sheet '{sheet_name}'. All either replied or reached max followups.")
        return True

    logging.info(f"Found {len(leads_to_process)} leads to consider for processing on sheet '{sheet_name}'.")

    for index, row in leads_to_process.iterrows():
        email_address = row[EMAIL_COLUMN]
        subject = row[SUBJECT_COLUMN]
        current_followup = row[FOLLOWUP_COLUMN]
        last_sent_str = row[LAST_SENT_DATE_COLUMN]
        current_status = row[STATUS_COLUMN]
        excel_row_display = df.index.get_loc(index) + 2

        logging.info(f"Evaluating lead from Excel row {excel_row_display}: {email_address}")

        # --- 1. Check for Reply ---
        # Check for reply if the status is not already 'Reply Received'
        if current_status != 'Reply Received':
             try:
                 # Call the simplified check_for_reply function
                 replied, reply_date = check_for_reply(email_address)

                 if replied:
                     df.loc[index, STATUS_COLUMN] = 'Reply Received'
                     # Ensure reply_date is a datetime object or can be formatted
                     # pywintypes.time usually works with strftime
                     df.loc[index, REPLY_DATE_COLUMN] = reply_date.strftime('%Y-%m-%d %H:%M:%S') if reply_date and hasattr(reply_date, 'strftime') else datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                     logging.info(f"Lead {email_address} (row {excel_row_display}) updated to 'Reply Received'. Skipping further emails.")
                     update_status_and_save(df, index, 'Reply Received', EXCEL_FILE_PATH, SHEET_NAME)
                     continue # Skip to the next lead
                 else:
                     logging.debug(f"No reply from {email_address} (row {excel_row_display}) found in inbox.")

             except Exception as e:
                 logging.error(f"Error during reply check for {email_address} (row {excel_row_display}): {e}", exc_info=True)
                 # Continue processing, don't stop if reply check fails


        # --- 2. Determine if eligible to send the next email ---
        send_next = False
        template_to_use = None
        next_followup_number = current_followup # Initialize with current, update if sending

        if pd.isna(current_status) or current_status == '':
            # Never sent before, send initial email (Followup 0)
            send_next = True
            next_followup_number = 0
            template_to_use = TEMPLATES.get(next_followup_number)
            logging.debug(f"Lead {email_address} (row {excel_row_display}) eligible for initial email (Followup 0).")

        elif current_status == 'Sent' and current_followup < len(TEMPLATES) - 1: # Check if not the last followup
            # Sent before, check delay for next follow-up
            try:
                last_sent_date_obj = datetime.strptime(last_sent_str, '%Y-%m-%d %H:%M:%S') if isinstance(last_sent_str, str) and last_sent_str.strip() else None
                if last_sent_date_obj:
                    days_since_last_sent = (datetime.now() - last_sent_date_obj).days
                    if days_since_last_sent >= FOLLOWUP_DELAY_DAYS:
                         send_next = True
                         next_followup_number = current_followup + 1
                         template_to_use = TEMPLATES.get(next_followup_number)
                         logging.debug(f"Lead {email_address} (row {excel_row_display}) eligible for Followup {next_followup_number} ({days_since_last_sent} days since last sent).")
                    else:
                         logging.debug(f"Lead {email_address} (row {excel_row_display}) not yet eligible for next followup ({days_since_last_sent} days since last sent, requires {FOLLOWUP_DELAY_DAYS}).")

                else:
                    # Should not happen if STATUS is 'Sent', but handle defensively
                    logging.warning(f"Could not parse LastSentDate '{last_sent_str}' for {email_address} (row {excel_row_display}) with status 'Sent'. Skipping.")
                    df.loc[index, STATUS_COLUMN] = 'Failed - Date Parse Error'
                    update_status_and_save(df, index, 'Failed - Date Parse Error', EXCEL_FILE_PATH, SHEET_NAME)


            except Exception as e:
                 logging.error(f"Error calculating days since last sent for {email_address} (row {excel_row_display}): {e}", exc_info=True)
                 df.loc[index, STATUS_COLUMN] = 'Failed - Date Calc Error'
                 update_status_and_save(df, index, 'Failed - Date Calc Error', EXCEL_FILE_PATH, SHEET_NAME)


        # --- 3. Send Email if Eligible ---
        if send_next and template_to_use:
            if not isinstance(email_address, str) or '@' not in email_address:
                logging.warning(f"Skipping Excel row {excel_row_display}: Invalid email '{email_address}'.")
                df.loc[index, STATUS_COLUMN] = 'Failed - Invalid Email'
                update_status_and_save(df, index, 'Failed - Invalid Email', EXCEL_FILE_PATH, SHEET_NAME)
                continue

            if pd.isna(subject) or not str(subject).strip():
                logging.warning(f"Skipping Excel row {excel_row_display}: Missing subject for '{email_address}'.")
                df.loc[index, STATUS_COLUMN] = 'Failed - Missing Subject'
                update_status_and_save(df, index, 'Failed - Missing Subject', EXCEL_FILE_PATH, SHEET_NAME)
                continue
            subject = str(subject) # Ensure subject is string

            placeholder_values = {}
            for placeholder, excel_col_name in PLACEHOLDERS.items():
                if excel_col_name and isinstance(excel_col_name, str) and excel_col_name in df.columns:
                    cell_value = row[excel_col_name]
                    actual_value = '' if pd.isna(cell_value) else cell_value
                    placeholder_values[placeholder] = actual_value
                else:
                    placeholder_values[placeholder] = ''
                    logging.warning(f"Placeholder '{placeholder}' (col: '{excel_col_name}') data missing for row {excel_row_display}.")

            # Check if placeholder values are valid
            if not placeholder_values:
                logging.warning(f"No valid placeholder data found for row {excel_row_display}.")
                continue


            logging.info(f"Attempting to send email to: {email_address} (Subject: '{subject}', Followup: {next_followup_number}) from template: {template_to_use}") # Use next_followup_number in log

            if send_outlook_email_from_template(template_to_use, email_address, subject, placeholder_values):
                # Update status, last sent date, and increment followup number
                df.loc[index, STATUS_COLUMN] = 'Sent'
                df.loc[index, LAST_SENT_DATE_COLUMN] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                df.loc[index, FOLLOWUP_COLUMN] = next_followup_number # Use next_followup_number here
                logging.info(f"Successfully sent email to {email_address} (row {excel_row_display}). Updated to Followup {next_followup_number}.") # Use next_followup_number in log
                update_status_and_save(df, index, 'Sent', EXCEL_FILE_PATH, SHEET_NAME)

                delay = random.uniform(MIN_DELAY_SECONDS, MAX_DELAY_SECONDS)
                logging.info(f"Waiting for {delay / 60:.2f} minutes before next email...")
                time.sleep(delay)
            else:
                logging.error(f"Failed to send email to {email_address} (row {excel_row_display}). Status set to 'Failed - Send Error'.")
                df.loc[index, STATUS_COLUMN] = 'Failed - Send Error'
                # Do NOT update FollowupNumber or LastSentDate on failure
                update_status_and_save(df, index, 'Failed - Send Error', EXCEL_FILE_PATH, SHEET_NAME)
        else:
             if current_status != 'Reply Received' and current_followup >= len(TEMPLATES): # Added Reply Received check here too
                 logging.debug(f"Lead {email_address} (row {excel_row_display}) has reached max followups ({current_followup}). Skipping.")
             # Other reasons for not sending (already replied, delay not met) are logged by debug messages above

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
    print(f"--- Using Outlook template: {TEMPLATES[0]}")
    print(f"--- Configured Placeholders: {list(PLACEHOLDERS.keys())}")
    print(f"--- Delay between emails: {MIN_DELAY_MINUTES}-{MAX_DELAY_MINUTES} minutes.")
    print("--- Press Ctrl+C to stop gracefully. ---")
    print("--- IMPORTANT: Ensure Excel file is CLOSED before running! ---")

    if not os.path.exists(TEMPLATES[0]):
         print(f"\nCRITICAL ERROR: Template file '{TEMPLATES[0]}' not found. Script aborted.")
         logging.critical(f"Template file not found. Script aborted.")
    else:

        try:
            success = process_leads(EXCEL_FILE_PATH, SHEET_NAME, TEMPLATES[0])
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