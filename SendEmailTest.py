import pandas as pd
import win32com.client as win32
import time
import random
import os
import logging
from datetime import datetime
import math

# --- Campaign Configurations ---
CAMPAIGNS = [
    {
        "campaign_name": "Medical Devices",
        "excel_file_path": r"C:\Users\z004m1jf\OneDrive - Siemens AG\Documents\4. Archives\Akhil Verma\Outreach\Medical Devices\Medical Devices Masters Test.xlsx",
        "sheet_name": "Cleaned Source",
        "email_templates": [
            r"C:\Users\z004m1jf\OneDrive - Siemens AG\Documents\4. Archives\Akhil Verma\Outreach\Medical Devices\FirstOutreachPLM.oft",
            r"C:\Users\z004m1jf\OneDrive - Siemens AG\Documents\4. Archives\Akhil Verma\Outreach\Medical Devices\Followup1.oft",
            r"C:\Users\z004m1jf\OneDrive - Siemens AG\Documents\4. Archives\Akhil Verma\Outreach\Medical Devices\Followup2.oft",
            r"C:\Users\z004m1jf\OneDrive - Siemens AG\Documents\4. Archives\Akhil Verma\Outreach\Medical Devices\Followup3.oft",
        ],
        "delays": [0, 3, 4, 5],  # days before each followup (0 for initial)
    },
    # {
    #     "campaign_name": "Mendix",
    #     "excel_file_path": r"C:\Users\z004m1jf\OneDrive - Siemens AG\Documents\4. Archives\Akhil Verma\Outreach\Mendix\Mendix Campaign Sheet Test.xlsx",
    #     "sheet_name": "Cleaned Source",
    #     "email_templates": [
    #         r"C:\Users\z004m1jf\OneDrive - Siemens AG\Documents\4. Archives\Akhil Verma\Outreach\Mendix\Mendix Email Templates\0.Mendix Initial Email.oft",
    #         r"C:\Users\z004m1jf\OneDrive - Siemens AG\Documents\4. Archives\Akhil Verma\Outreach\Mendix\Mendix Email Templates\1.Mendix First Follwoup.oft",
    #         r"C:\Users\z004m1jf\OneDrive - Siemens AG\Documents\4. Archives\Akhil Verma\Outreach\Mendix\Mendix Email Templates\2.Mendix First Followup.oft",
    #         r"C:\Users\z004m1jf\OneDrive - Siemens AG\Documents\4. Archives\Akhil Verma\Outreach\Mendix\Mendix Email Templates\3.Mendix First Followup.oft",
    #     ],
    #     "delays": [0, 3, 4, 5], #days before each follwoup (0 for initial)
    # },
    # Add more campaigns as needed
]

# --- Global Settings ---
DAILY_LIMIT = 75
MIN_DELAY_MINUTES = 3
MAX_DELAY_MINUTES = 4
MIN_DELAY_SECONDS = MIN_DELAY_MINUTES * 60
MAX_DELAY_SECONDS = MAX_DELAY_MINUTES * 60

# --- Dynamic Daily Quota Settings ---
INITIALS_PER_DAY = 30      # Set your preferred initial quota before running
FOLLOWUPS_PER_DAY = 45     # Set your preferred followup quota before running
# The script will use any leftover quota for the other category if needed
# Ensure INITIALS_PER_DAY + FOLLOWUPS_PER_DAY >= DAILY_LIMIT for full flexibility

# --- Excel Column Names ---
EMAIL_COLUMN = 'EmailAddress'
SUBJECT_COLUMN = 'Subject'
NAME_COLUMN = 'FirstName'
COMPANY_COLUMN = 'CompanyNameSimplified'
SALUTATION_COLUMN = 'Salutation'
LAST_SENT_DATE_COLUMN = 'LastSentDate'
FOLLOWUP_COLUMN = 'FollowupNumber'
REPLY_DATE_COLUMN = 'ReplyDate'
STATUS_COLUMN = 'SendStatus'

# --- Placeholder Mapping ---
PLACEHOLDERS = {
    '[FirstName]': NAME_COLUMN,
    '[CompanyNameSimplified]': COMPANY_COLUMN,
    '[Salutation]': SALUTATION_COLUMN,
}

# --- Logging Setup ---
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# --- Email Sending Function ---
def send_outlook_email_from_template(template_path, recipient_email, subject, placeholder_data):
    """
    Creates and sends an email from an Outlook Template, replacing placeholders.
    """
    if not os.path.exists(template_path):
        logging.error(f"Outlook template file not found at: {template_path}")
        return False
    try:
        outlook_app = win32.Dispatch('Outlook.Application')
        mail = outlook_app.CreateItemFromTemplate(template_path)
        logging.debug(f"Created mail item from template: {template_path}")

        mail.To = recipient_email
        mail.Subject = subject
        logging.debug(f"Set To: {recipient_email}, Subject: {subject}")

        try:
            body_content = None
            is_html = False
            body_type_used = "None"
            if hasattr(mail, 'HTMLBody') and mail.HTMLBody:
                body_content = mail.HTMLBody
                is_html = True
                body_type_used = "HTMLBody"
            elif hasattr(mail, 'Body') and mail.Body:
                body_content = mail.Body
                body_type_used = "Body"
            
            if body_content:
                logging.debug(f"Using {body_type_used} for replacements for {recipient_email}.")
                replacements_made = 0
                for placeholder, value in placeholder_data.items():
                    # Replace NaN or None with empty string
                    if value is None or (isinstance(value, float) and math.isnan(value)):
                        replacement_value = ''
                    else:
                        replacement_value = str(value)
                    if placeholder in body_content:
                        body_content = body_content.replace(placeholder, replacement_value)
                        replacements_made += 1
                        logging.debug(f"   -> Replaced '{placeholder}'.")
                    else:
                        logging.debug(f"   -> Placeholder '{placeholder}' not found in {body_type_used}.")
                logging.info(f"Placeholders processed for {recipient_email}. Replacements: {replacements_made}/{len(placeholder_data)}.")
                if is_html:
                    mail.HTMLBody = body_content
                else:
                    mail.Body = body_content
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
    
def send_followup_as_reply(recipient_email, subject, placeholder_data, followup_template_path):
    """
    Sends a follow-up email as a reply to the last sent message to the recipient,
    so that the follow-up appears in the same thread/conversation.
    """
    try:
        outlook_app = win32.Dispatch('Outlook.Application')
        namespace = outlook_app.GetNamespace("MAPI")
        sent_folder = namespace.GetDefaultFolder(5)  # 5 = Sent Items

        # Find the last sent message to this recipient (search in To field)
        messages = sent_folder.Items
        messages.Sort("[SentOn]", True)  # Sort descending by sent date

        last_sent = None
        for msg in messages:
            try:
                # Check if recipient_email is in the To field (case-insensitive)
                if recipient_email.lower() in str(msg.To).lower():
                    last_sent = msg
                    break
            except Exception:
                continue

        if last_sent is None:
            logging.warning(f"No previous sent message found for {recipient_email}. Sending as new email.")
            return send_outlook_email_from_template(followup_template_path, recipient_email, subject, placeholder_data)

        reply = last_sent.Reply()

        # Load follow-up template body (from .oft file)
        if not os.path.exists(followup_template_path):
            logging.error(f"Follow-up template not found: {followup_template_path}")
            return False

        temp_mail = outlook_app.CreateItemFromTemplate(followup_template_path)
        if hasattr(temp_mail, 'HTMLBody') and temp_mail.HTMLBody:
            template_body = temp_mail.HTMLBody
            is_html = True
        elif hasattr(temp_mail, 'Body') and temp_mail.Body:
            template_body = temp_mail.Body
            is_html = False
        else:
            logging.error("Follow-up template has no body.")
            return False

        # Replace placeholders
        #logging.debug(f"Template body before replacement: {template_body}")
        for placeholder, value in placeholder_data.items():
            # Replace NaN or None with empty string
            if value is None or (isinstance(value, float) and math.isnan(value)):
                replacement_value = ''
            else:
                replacement_value = str(value)
            # Replace plain placeholder
            template_body = template_body.replace(placeholder, replacement_value)
            # Replace HTML-wrapped placeholder (for Word spellcheck tags)
            html_placeholder = f"[<span class=SpellE>{placeholder.strip('[]')}</span>]"
            template_body = template_body.replace(html_placeholder, replacement_value)
        #logging.debug(f"Template body after replacement: {template_body}")

        # Insert the follow-up template above the original message
        if is_html and hasattr(reply, 'HTMLBody'):
            reply.HTMLBody = template_body.rstrip() + reply.HTMLBody.lstrip()
        else:
            reply.Body = template_body + "\n\n" + reply.Body

        # Optionally update subject if needed
        if subject:
            reply.Subject = subject

        reply.To = recipient_email  # Ensure correct recipient
        reply.Send()
        logging.info(f"Follow-up sent as reply to {recipient_email}")
        return True

    except Exception as e:
        logging.error(f"Failed to send follow-up as reply for {recipient_email}: {e}", exc_info=True)
        return False

# --- Reply Check Function ---
def check_for_reply(recipient_email):
    """
    Checks Outlook inbox for any email from the recipient email address using Restrict filter.
    """
    try:
        outlook_app = win32.Dispatch('Outlook.Application')
        namespace = outlook_app.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6) # 6 refers to the Inbox folder
        clean_recipient_email = recipient_email.replace("'", "''")
        filter_string = f"[SenderEmailAddress] = '{clean_recipient_email}'"
        restricted_items = inbox.Items.Restrict(filter_string)
        if restricted_items.Count > 0:
            first_item_received_time = restricted_items.GetFirst().ReceivedTime
            return True, first_item_received_time
        else:
            return False, None
    except Exception as e:
        logging.error(f"Error checking Outlook inbox for replies for {recipient_email}: {e}", exc_info=True)
        return False, None

# --- Load Campaign Data ---
def load_campaign_data(campaign):
    """
    Loads campaign Excel data into a DataFrame, ensuring required columns exist.
    """
    try:
        dtype_spec = {
            STATUS_COLUMN: str,
            LAST_SENT_DATE_COLUMN: str,
            REPLY_DATE_COLUMN: str,
            FOLLOWUP_COLUMN: float
        }
        df = pd.read_excel(campaign["excel_file_path"], sheet_name=campaign["sheet_name"], dtype=dtype_spec)
        # Ensure required columns exist
        for col in [LAST_SENT_DATE_COLUMN, FOLLOWUP_COLUMN, REPLY_DATE_COLUMN, STATUS_COLUMN]:
            if col not in df.columns:
                if col == FOLLOWUP_COLUMN:
                    df[col] = 0
                else:
                    df[col] = ''
        df[FOLLOWUP_COLUMN] = df[FOLLOWUP_COLUMN].fillna(0).astype(int)
        return df
    except Exception as e:
        logging.error(f"Error loading campaign '{campaign['campaign_name']}': {e}", exc_info=True)
        return None

# --- Save Campaign Data ---
def save_campaign_data(campaign, df):
    """
    Saves the DataFrame back to the campaign's Excel file.
    """
    try:
        with pd.ExcelWriter(campaign["excel_file_path"], engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=campaign["sheet_name"], index=False)
        logging.info(f"Saved updates to '{campaign['campaign_name']}' ({campaign['excel_file_path']})")
    except Exception as e:
        logging.error(f"Error saving campaign '{campaign['campaign_name']}': {e}", exc_info=True)

# --- Update Replies for Campaign ---
def update_replies_for_campaign(df):
    """
    Checks for replies for all leads in the DataFrame and updates their status.
    """
    for idx, row in df.iterrows():
        email_address = row.get(EMAIL_COLUMN, '')
        current_status = row.get(STATUS_COLUMN, '')
        if current_status == 'Reply Received':
            continue
        replied, reply_date = check_for_reply(email_address)
        if replied:
            df.at[idx, STATUS_COLUMN] = 'Reply Received'
            df.at[idx, REPLY_DATE_COLUMN] = reply_date.strftime('%Y-%m-%d %H:%M:%S') if reply_date else datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            logging.info(f"Reply received from {email_address}. Marked as 'Reply Received'.")
    return df

# --- Count Emails Sent Today ---
def count_emails_sent_today(dfs):
    """
    Counts how many emails have already been sent today across all campaigns.
    """
    today_str = datetime.now().strftime('%Y-%m-%d')
    count = 0
    for df in dfs:
        count += df[LAST_SENT_DATE_COLUMN].apply(lambda x: str(x).startswith(today_str)).sum()
    return count

def count_initials_sent_today(dfs):
    today_str = datetime.now().strftime('%Y-%m-%d')
    count = 0
    for df in dfs:
        count += df[
            (df[LAST_SENT_DATE_COLUMN].apply(lambda x: str(x).startswith(today_str))) &
            (df[FOLLOWUP_COLUMN] == 0)
        ].shape[0]
    return count

def count_followups_sent_today(dfs):
    today_str = datetime.now().strftime('%Y-%m-%d')
    count = 0
    for df in dfs:
        count += df[
            (df[LAST_SENT_DATE_COLUMN].apply(lambda x: str(x).startswith(today_str))) &
            (df[FOLLOWUP_COLUMN] > 0)
        ].shape[0]
    return count

# --- Get Next Eligible Lead for Campaign ---
def get_next_eligible_lead(df, campaign, today_str):
    """
    Returns the next eligible lead for sending, or None if none found.
    """
    for idx, row in df.iterrows():
        email_address = row.get(EMAIL_COLUMN, '')
        subject = row.get(SUBJECT_COLUMN, '')
        current_followup = row.get(FOLLOWUP_COLUMN, 0)
        last_sent_str = row.get(LAST_SENT_DATE_COLUMN, '')
        current_status = row.get(STATUS_COLUMN, '')

        # Skip if replied or already sent today
        if current_status == 'Reply Received':
            continue
        if last_sent_str and str(last_sent_str).startswith(today_str):
            continue

        # Determine if eligible for initial or followup
        if not current_status or pd.isna(current_status):
            template_idx = 0
            delay_days = campaign["delays"][0]
        elif current_status == 'Sent' and current_followup < len(campaign["email_templates"]) - 1:
            if last_sent_str:
                try:
                    last_sent_date = datetime.strptime(last_sent_str, '%Y-%m-%d %H:%M:%S')
                    days_since_last = (datetime.now() - last_sent_date).days
                    delay_days = campaign["delays"][current_followup + 1]
                    if days_since_last < delay_days:
                        continue
                    template_idx = current_followup + 1
                except Exception as e:
                    logging.warning(f"Date parse error for {email_address}: {e}")
                    continue
            else:
                continue
        else:
            continue

        # Prepare placeholder values
        placeholder_values = {ph: row.get(col, '') for ph, col in PLACEHOLDERS.items()}
        return idx, email_address, subject, template_idx, placeholder_values
    return None

def get_all_eligible_initials(dfs, campaigns, today_str):
    initials = []
    for i, (df, campaign) in enumerate(zip(dfs, campaigns)):
        for idx, row in df.iterrows():
            current_status = row.get(STATUS_COLUMN, '')
            last_sent_str = row.get(LAST_SENT_DATE_COLUMN, '')
            if current_status == 'Reply Received':
                continue
            if last_sent_str and str(last_sent_str).startswith(today_str):
                continue
            if not current_status or pd.isna(current_status):
                email_address = row.get(EMAIL_COLUMN, '')
                subject = row.get(SUBJECT_COLUMN, '')
                template_idx = 0
                placeholder_values = {ph: row.get(col, '') for ph, col in PLACEHOLDERS.items()}
                initials.append((i, idx, email_address, subject, template_idx, placeholder_values))
    return initials

def get_all_eligible_followups(dfs, campaigns, today_str):
    followups = []
    for i, (df, campaign) in enumerate(zip(dfs, campaigns)):
        for idx, row in df.iterrows():
            current_status = row.get(STATUS_COLUMN, '')
            last_sent_str = row.get(LAST_SENT_DATE_COLUMN, '')
            current_followup = row.get(FOLLOWUP_COLUMN, 0)
            if current_status == 'Reply Received':
                continue
            if last_sent_str and str(last_sent_str).startswith(today_str):
                continue
            if current_status == 'Sent' and current_followup < len(campaign["email_templates"]) - 1:
                if last_sent_str:
                    try:
                        last_sent_date = datetime.strptime(last_sent_str, '%Y-%m-%d %H:%M:%S')
                        days_since_last = (datetime.now() - last_sent_date).days
                        delay_days = campaign["delays"][current_followup + 1]
                        if days_since_last < delay_days:
                            continue
                        template_idx = current_followup + 1
                        email_address = row.get(EMAIL_COLUMN, '')
                        subject = row.get(SUBJECT_COLUMN, '')
                        placeholder_values = {ph: row.get(col, '') for ph, col in PLACEHOLDERS.items()}
                        followups.append((i, idx, email_address, subject, template_idx, placeholder_values))
                    except Exception:
                        continue
    return followups

def get_all_eligible_initials_per_campaign(dfs, campaigns, today_str):
    initials_per_campaign = []
    for i, (df, campaign) in enumerate(zip(dfs, campaigns)):
        campaign_initials = []
        for idx, row in df.iterrows():
            current_status = row.get(STATUS_COLUMN, '')
            last_sent_str = row.get(LAST_SENT_DATE_COLUMN, '')
            if current_status == 'Reply Received':
                continue
            if last_sent_str and str(last_sent_str).startswith(today_str):
                continue
            if not current_status or pd.isna(current_status):
                email_address = row.get(EMAIL_COLUMN, '')
                subject = row.get(SUBJECT_COLUMN, '')
                template_idx = 0
                placeholder_values = {ph: row.get(col, '') for ph, col in PLACEHOLDERS.items()}
                campaign_initials.append((i, idx, email_address, subject, template_idx, placeholder_values))
        initials_per_campaign.append(campaign_initials)
    return initials_per_campaign

def get_all_eligible_followups_per_campaign(dfs, campaigns, today_str):
    followups_per_campaign = []
    for i, (df, campaign) in enumerate(zip(dfs, campaigns)):
        campaign_followups = []
        for idx, row in df.iterrows():
            current_status = row.get(STATUS_COLUMN, '')
            last_sent_str = row.get(LAST_SENT_DATE_COLUMN, '')
            current_followup = row.get(FOLLOWUP_COLUMN, 0)
            if current_status == 'Reply Received':
                continue
            if last_sent_str and str(last_sent_str).startswith(today_str):
                continue
            if current_status == 'Sent' and current_followup < len(campaign["email_templates"]) - 1:
                if last_sent_str:
                    try:
                        last_sent_date = datetime.strptime(last_sent_str, '%Y-%m-%d %H:%M:%S')
                        days_since_last = (datetime.now() - last_sent_date).days
                        delay_days = campaign["delays"][current_followup + 1]
                        if days_since_last < delay_days:
                            continue
                        template_idx = current_followup + 1
                        email_address = row.get(EMAIL_COLUMN, '')
                        subject = row.get(SUBJECT_COLUMN, '')
                        placeholder_values = {ph: row.get(col, '') for ph, col in PLACEHOLDERS.items()}
                        campaign_followups.append((i, idx, email_address, subject, template_idx, placeholder_values))
                    except Exception:
                        continue
        followups_per_campaign.append(campaign_followups)
    return followups_per_campaign

def interleave_round_robin(lists, limit):
    """
    Interleave lists in round robin fashion up to 'limit' total items.
    """
    result = []
    pointers = [0] * len(lists)
    while len(result) < limit:
        added = False
        for i, l in enumerate(lists):
            if pointers[i] < len(l):
                result.append(l[pointers[i]])
                pointers[i] += 1
                if len(result) >= limit:
                    break
                added = True
        if not added:
            break
    return result

def round_robin_send_dynamic_quota(campaigns, dfs, emails_sent_today, daily_limit, initials_quota, followups_quota, prioritize='followups'):
    today_str = datetime.now().strftime('%Y-%m-%d')

    initials_sent = count_initials_sent_today(dfs)
    followups_sent = count_followups_sent_today(dfs)
    print(f"Summary: Initials sent today: {initials_sent}, Follow-ups sent today: {followups_sent}")

    initials_left = max(0, initials_quota - initials_sent)
    followups_left = max(0, followups_quota - followups_sent)
    total_left = daily_limit - emails_sent_today

    initials_per_campaign = get_all_eligible_initials_per_campaign(dfs, campaigns, today_str)
    followups_per_campaign = get_all_eligible_followups_per_campaign(dfs, campaigns, today_str)

    initials = interleave_round_robin(initials_per_campaign, initials_left)
    followups = interleave_round_robin(followups_per_campaign, followups_left)

    # If quota not filled, use leftover for the other category
    if prioritize == 'initials':
        slots_left = total_left - len(initials)
        if slots_left > 0:
            more_followups = interleave_round_robin(
                [l[followups_left:] for l in followups_per_campaign], slots_left
            )
            followups += more_followups
    else:
        slots_left = total_left - len(followups)
        if slots_left > 0:
            more_initials = interleave_round_robin(
                [l[initials_left:] for l in initials_per_campaign], slots_left
            )
            initials += more_initials

    # Interleave initials and followups for round robin effect
    combined = []
    if prioritize == 'initials':
        max_len = max(len(initials), len(followups))
        for i in range(max_len):
            if i < len(initials):
                combined.append(('initial',) + initials[i])
            if i < len(followups):
                combined.append(('followup',) + followups[i])
    else:
        max_len = max(len(followups), len(initials))
        for i in range(max_len):
            if i < len(followups):
                combined.append(('followup',) + followups[i])
            if i < len(initials):
                combined.append(('initial',) + initials[i])

    sent_count = 0
    try:
        for entry in combined:
            if sent_count + emails_sent_today >= daily_limit:
                break
            kind, i, idx, email_address, subject, template_idx, placeholder_values = entry
            campaign = campaigns[i]
            df = dfs[i]
            template_path = campaign["email_templates"][template_idx]

            if not isinstance(email_address, str) or '@' not in email_address:
                logging.warning(f"Skipping: Invalid email '{email_address}'.")
                df.at[idx, STATUS_COLUMN] = 'Failed - Invalid Email'
                continue
            if pd.isna(subject) or not str(subject).strip():
                logging.warning(f"Skipping: Missing subject for '{email_address}'.")
                df.at[idx, STATUS_COLUMN] = 'Failed - Missing Subject'
                continue

            logging.info(f"Sending {kind} email to {email_address} (Campaign: {campaign['campaign_name']}, Template: {template_idx})")
            if kind == 'initial':
                send_success = send_outlook_email_from_template(template_path, email_address, subject, placeholder_values)
            else:
                send_success = send_followup_as_reply(email_address, subject, placeholder_values, template_path)

            if send_success:
                df.at[idx, STATUS_COLUMN] = 'Sent'
                df.at[idx, LAST_SENT_DATE_COLUMN] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                df.at[idx, FOLLOWUP_COLUMN] = template_idx
                sent_count += 1
                save_campaign_data(campaign, df)
                delay = random.uniform(MIN_DELAY_SECONDS, MAX_DELAY_SECONDS)
                logging.info(f"Waiting {delay/60:.2f} minutes before next email...")
                time.sleep(delay)
            else:
                df.at[idx, STATUS_COLUMN] = 'Failed - Send Error'
                save_campaign_data(campaign, df)
    except KeyboardInterrupt:
        logging.warning("Script interrupted by user. Saving all progress before exit...")
        for i, campaign in enumerate(campaigns):
            save_campaign_data(campaign, dfs[i])
        print("\nScript stopped by user. All progress saved.")
        return
    except Exception as e:
        logging.critical(f"Unexpected error in round robin send: {e}", exc_info=True)
        for i, campaign in enumerate(campaigns):
            save_campaign_data(campaign, dfs[i])
        print("\nCritical error occurred. All progress saved.")
        return

    # Final save after all sends
    for i, campaign in enumerate(campaigns):
        save_campaign_data(campaign, dfs[i])
    logging.info(f"Dynamic quota sending complete. Total emails sent this run: {sent_count}")

# --- Main Execution ---
if __name__ == "__main__":
    print("Starting Outlook Email Automation Script (Round Robin Multi-Campaign Mode)")
    print(f"--- Global daily send limit: {DAILY_LIMIT} emails")
    print(f"--- Delay between emails: {MIN_DELAY_MINUTES}-{MAX_DELAY_MINUTES} minutes.")
    print("--- Press Ctrl+C to stop gracefully. ---")
    print("--- IMPORTANT: Ensure all Excel files are CLOSED before running! ---")

    # 1. Load all campaign data
    dfs = []
    for campaign in CAMPAIGNS:
        print(f"Loading campaign: {campaign['campaign_name']} ({campaign['excel_file_path']}, Sheet: {campaign['sheet_name']})")
        df = load_campaign_data(campaign)
        if df is None:
            print(f"Failed to load campaign: {campaign['campaign_name']}. Skipping.")
            continue
        dfs.append(df)

    if not dfs:
        print("No campaigns loaded successfully. Exiting.")
        exit(1)

    # 2. Check for replies and update all campaigns
    print("Checking for replies and updating all campaigns...")
    for i, campaign in enumerate(CAMPAIGNS):
        dfs[i] = update_replies_for_campaign(dfs[i])
        save_campaign_data(campaign, dfs[i])

    # 3. Count emails already sent today
    emails_sent_today = count_emails_sent_today(dfs)
    print(f"Emails already sent today (across all campaigns): {emails_sent_today}")
    if emails_sent_today >= DAILY_LIMIT:
        print("Daily limit already reached. No emails will be sent.")
    else:
        # 4. Round robin send
        print("Starting round robin email sending...")
    # You can set prioritize='initials' or 'followups' as you wish
    round_robin_send_dynamic_quota(
        CAMPAIGNS, dfs, emails_sent_today, DAILY_LIMIT,
        INITIALS_PER_DAY, FOLLOWUPS_PER_DAY,
        prioritize='followups'  # or 'initials'
    )
    print("All campaigns processed or daily limit reached. Script finished.")