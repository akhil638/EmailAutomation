import unittest
from unittest.mock import patch, MagicMock
import pandas as pd
from datetime import datetime

# Import your functions from the main script
import SendEmailTest

class TestEmailAutomation(unittest.TestCase):

    @patch('win32com.client.Dispatch')
    def test_send_outlook_email_from_template(self, mock_dispatch):
        # Mock Outlook and mail item
        mock_outlook = MagicMock()
        mock_mail = MagicMock()
        mock_mail.HTMLBody = "Hello [FirstName], welcome to [CompanyNameSimplified]!"
        mock_outlook.CreateItemFromTemplate.return_value = mock_mail
        mock_dispatch.return_value = mock_outlook

        placeholder_data = {
            '[FirstName]': 'Akhil',
            '[CompanyNameSimplified]': 'Siemens'
        }

        # Should succeed and replace placeholders
        result = SendEmailTest.send_outlook_email_from_template(
            'dummy.oft', 'test@example.com', 'Subject', placeholder_data
        )
        self.assertTrue(result)
        self.assertIn('Akhil', mock_mail.HTMLBody)
        self.assertIn('Siemens', mock_mail.HTMLBody)
        mock_mail.Send.assert_called_once()

    @patch('win32com.client.Dispatch')
    def test_send_followup_as_reply(self, mock_dispatch):
        # Mock Outlook, sent items, and reply
        mock_outlook = MagicMock()
        mock_mail = MagicMock()
        mock_mail.To = 'test@example.com'
        mock_mail.HTMLBody = "Previous message"
        mock_reply = MagicMock()
        mock_reply.HTMLBody = "Previous message"
        mock_mail.Reply.return_value = mock_reply

        mock_sent_folder = MagicMock()
        mock_items = MagicMock()
        mock_items.__iter__.return_value = [mock_mail]
        mock_items.Sort = MagicMock()
        mock_sent_folder.Items = mock_items
        mock_namespace = MagicMock()
        mock_namespace.GetDefaultFolder.return_value = mock_sent_folder
        mock_outlook.GetNamespace.return_value = mock_namespace
        mock_outlook.CreateItemFromTemplate.return_value = MagicMock(HTMLBody="Followup [FirstName]", Body=None)
        mock_dispatch.return_value = mock_outlook

        placeholder_data = {'[FirstName]': 'Akhil'}

        result = SendEmailTest.send_followup_as_reply(
            'test@example.com', 'Followup Subject', placeholder_data, 'dummy_followup.oft'
        )
        self.assertTrue(result)
        mock_reply.Send.assert_called_once()

    @patch('pandas.read_excel')
    def test_load_campaign_data_missing_columns(self, mock_read_excel):
        # DataFrame missing some columns
        df = pd.DataFrame({'EmailAddress': ['a@b.com']})
        mock_read_excel.return_value = df
        campaign = SendEmailTest.CAMPAIGNS[0]
        result_df = SendEmailTest.load_campaign_data(campaign)
        # Should add missing columns
        self.assertIn(SendEmailTest.STATUS_COLUMN, result_df.columns)
        self.assertIn(SendEmailTest.LAST_SENT_DATE_COLUMN, result_df.columns)
        self.assertIn(SendEmailTest.FOLLOWUP_COLUMN, result_df.columns)
        self.assertIn(SendEmailTest.REPLY_DATE_COLUMN, result_df.columns)

    def test_get_next_eligible_lead(self):
        # Test lead selection logic
        df = pd.DataFrame([
            {
                'EmailAddress': 'a@b.com',
                'Subject': 'Hello',
                'SendStatus': '',
                'LastSentDate': '',
                'FollowupNumber': 0
            }
        ])
        campaign = {
            "email_templates": ['t1.oft', 't2.oft'],
            "delays": [0, 3]
        }
        today_str = '2025-05-24'
        result = SendEmailTest.get_next_eligible_lead(df, campaign, today_str)
        self.assertIsNotNone(result)
        idx, email, subject, template_idx, placeholder_values = result
        self.assertEqual(email, 'a@b.com')
        self.assertEqual(template_idx, 0)

    def test_count_emails_sent_today(self):
        today = datetime.now().strftime('%Y-%m-%d')
        df = pd.DataFrame([
            {SendEmailTest.LAST_SENT_DATE_COLUMN: today + ' 10:00:00'},
            {SendEmailTest.LAST_SENT_DATE_COLUMN: '2020-01-01 10:00:00'}
        ])
        count = SendEmailTest.count_emails_sent_today([df])
        self.assertEqual(count, 1)

if __name__ == '__main__':
    unittest.main()