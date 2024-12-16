import pandas as pd
import os
import base64
import pickle
from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.generator import BytesGenerator
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import io

# Constants for Gmail API
SCOPES = ['https://www.googleapis.com/auth/gmail.send']
CLIENT_SECRET_FILE = 'credentials.json'
TOKEN_PICKLE_FILE = 'token.pickle'

def get_authenticated_service():
    creds = None
    if os.path.exists(TOKEN_PICKLE_FILE):
        with open(TOKEN_PICKLE_FILE, 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_PICKLE_FILE, 'wb') as token:
            pickle.dump(creds, token)
    return build('gmail', 'v1', credentials=creds)

def create_message(sender, to, subject, body):
    """Create a message object."""
    message = MIMEMultipart('alternative')
    message['To'] = to
    message['From'] = sender
    message['Subject'] = subject
    part = MIMEText(body, 'html')
    message.attach(part)
    return message

def send_message(service, user_id, message):
    try:
        raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
        sent_message = service.users().messages().send(userId=user_id, body={'raw': raw_message}).execute()
        print(f"Message Id: {sent_message['id']} sent successfully.")
        return sent_message
    except Exception as e:
        print(f"An error occurred: {e}")
        return None

def main():
    # Read the CSV data
    df = pd.read_csv('data.csv')

    # Columns that we consider for medians
    value_columns = [
        'total_checkins',
        'total_unique_checkins',
        'months_since_first_checkin',
        'total_membership_claimed',
        'total_payment_value',
        'total_transactions',
        'avg_fly_balance_per_employee',
        'median_fly_balance_per_employee',
        'pct_employees_with_vaulted_cards',
        'pct_employees_with_fly_spent',
        'total_employees',
        'repeat_checkins_last_3_months'
    ]

    # Mapping columns to layman's terms
    column_mapping = {
        'total_checkins': 'Number of Visits',
        'total_unique_checkins': 'Number of Unique Guests',
        'months_since_first_checkin': 'Months Since First Guest Visit',
        'total_membership_claimed': 'Memberships Claimed',
        'total_payment_value': 'Total Payment Value ($)',
        'total_transactions': 'Number of Transactions',
        'avg_fly_balance_per_employee': 'Average Loyalty Points per Employee',
        'median_fly_balance_per_employee': 'Median Loyalty Points per Employee',
        'pct_employees_with_vaulted_cards': 'Percentage of Employees with Vaulted Payment Methods',
        'pct_employees_with_fly_spent': 'Percentage of Employees Using Loyalty Points',
        'total_employees': 'Total Employees',
        'repeat_checkins_last_3_months': 'Repeat Visits in the Last 3 Months'
    }

    # Convert columns to numeric if needed
    for col in value_columns:
        df[col] = df[col].astype(str).str.replace(r'[\$,]', '', regex=True)
        df[col] = pd.to_numeric(df[col], errors='coerce')

    # Compute medians
    medians = df[value_columns].median(numeric_only=True)

    # Authenticate Gmail API
    service = get_authenticated_service()
    sender_email = 'maddie.weber@blackbird.xyz'
    recipient_email = 'maddie.weber@blackbird.xyz'  # Send all attachments to Maddie

    # We'll store all attachments in this list
    all_attachments = []

    # Process each restaurant
    for _, row in df.iterrows():
        restaurant_name = row.get('restaurant_name', 'Restaurant Team')
        if pd.isnull(restaurant_name) or restaurant_name.strip() == '':
            restaurant_name = 'Restaurant Team'

        # Determine which columns meet/exceed the median
        highlight_values = {}
        for col in value_columns:
            val = row[col]
            if (pd.notnull(val) and 
                col in medians.index and 
                pd.notnull(medians[col]) and 
                val >= medians[col]):
                # Format nicely: if it's a float but an integer value, convert to int
                if isinstance(val, float) and val.is_integer():
                    val = int(val)
                highlight_values[col] = val

        # If fewer than 2 metrics meet/exceed the median, we won't show any stats.
        # This includes the case of 0 or exactly 1 stat.
        show_stats = len(highlight_values) >= 2

        # Build personalized email body
        # General message always included
        email_body = f"""
        <html>
        <body>
            <p>Hi {restaurant_name},</p>
            <p>As the holiday season sets in, we want to express our heartfelt gratitude 
            for your partnership over these past few months. Working together has been an incredible 
            journey, and we truly value the trust you've placed in us.</p>

            <p>We're looking forward to the new year ahead, especially as we step into 2025. 
            Our team is excited to continue evolving our platform to better serve your guests 
            and help grow your business. Your success remains our top priority, and we can't 
            wait to show you what's in store.</p>
        """

        # Include metrics only if show_stats is True
        if show_stats:
            email_body += """
            <p>We also wanted to highlight some of the areas where you've excelled:</p>

            <table border="1" cellspacing="0" cellpadding="5" style="border-collapse: collapse;">
                <tr><th>Metric</th><th>Value</th></tr>
            """
            for col, val in highlight_values.items():
                nice_name = column_mapping.get(col, col)
                email_body += f"<tr><td>{nice_name}</td><td>{val}</td></tr>"
            email_body += "</table>"
        
        # Conclude the email
        email_body += f"""
            <p>Thank you once again for your continued support. 
            Wishing you wonderful holidays and an even brighter 2025!</p>

            <p>Warmly,<br>
            The Blackbird Team</p>
        </body>
        </html>
        """

        subject = f"Year-End Greetings from Blackbird: Looking Forward to 2025"
        
        # Placeholder for the client recipient if needed
        # If you have a column in your CSV for the client's email, use row['email'].
        # Otherwise, just leave a placeholder or remove if not required.
        client_email = restaurant_name.replace(' ', '').lower() + "@example.com"

        client_message = create_message(sender_email, client_email, subject, email_body)

        # Convert client_message to .eml
        eml_file = io.BytesIO()
        gen = BytesGenerator(eml_file)
        gen.flatten(client_message)
        eml_content = eml_file.getvalue()

        # Create attachment
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(eml_content)
        encoders.encode_base64(part)
        filename = f"{restaurant_name.replace(' ', '_')}_email.eml"
        part.add_header('Content-Disposition', f'attachment; filename="{filename}"')

        # Add to the list of all attachments
        all_attachments.append(part)

    # After processing all restaurants, send one email with all attachments
    if all_attachments:
        subject = "All Personalized Restaurant Emails"
        body = """Hi Maddie,

Please find attached all the personalized email drafts for each of our SF launch partners. 
These emails already reflect holiday well-wishes and enthusiasm for 2025, 
and some include their key metrics (only if they met or exceeded the median in two or more areas).

Happy Holidays,
The Blackbird Team
"""
        message = MIMEMultipart()
        message['To'] = recipient_email
        message['From'] = sender_email
        message['Subject'] = subject
        message.attach(MIMEText(body, 'plain'))

        for attachment in all_attachments:
            message.attach(attachment)

        send_message(service, 'me', message)
    else:
        print("No attachments to send.")

if __name__ == '__main__':
    main()
