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

    # Columns that we consider for medians (EXCLUDING total_employees and total_membership_claimed)
    value_columns = [
        'total_checkins',
        'total_unique_checkins',
        'months_since_first_checkin',
        'total_payment_value',
        'total_transactions',
        'avg_fly_balance_per_employee',
        'median_fly_balance_per_employee',
        'pct_employees_with_vaulted_cards',
        'pct_employees_with_fly_spent',
        'repeat_checkins_last_3_months'
    ]

    # Mapping columns to layman's terms
    column_mapping = {
        'total_checkins': 'Number of Blackbird Tap-ins',
        'total_unique_checkins': 'Number of Unique Guests',
        'total_payment_value': 'Total Payment Value in $FLY (converted to $USD)',
        'total_transactions': 'Number of $FLY Transactions',
        # Updated to emphasize $FLY instead of "Loyalty Points"
        'avg_fly_balance_per_employee': 'Average $FLY per Employee (converted to $USD)',
        'median_fly_balance_per_employee': 'Median $FLY per Employee (converted to $USD)',
        # Updated phrasing for vaulted cards
        'pct_employees_with_vaulted_cards': 'Percent of employees who have added a card and are ready to use their $FLY for meals',
        # Already referencing $FLY, just dropping "loyalty points"
        'pct_employees_with_fly_spent': 'Percentage of Employees Using $FLY',
        'repeat_checkins_last_3_months': 'Repeat Visits since Launch'
    }

    # Convert columns to numeric if needed
    for col in value_columns:
        df[col] = df[col].astype(str).str.replace(r'[\$,]', '', regex=True)
        df[col] = pd.to_numeric(df[col], errors='coerce')

    # Convert monetary columns from $FLY to USD by dividing by 10^20
    if 'total_payment_value' in df.columns:
        df['total_payment_value'] = df['total_payment_value'] / (10**20)

    if 'avg_fly_balance_per_employee' in df.columns:
        df['avg_fly_balance_per_employee'] = df['avg_fly_balance_per_employee'] / (10**20)

    if 'median_fly_balance_per_employee' in df.columns:
        df['median_fly_balance_per_employee'] = df['median_fly_balance_per_employee'] / (10**20)

    percentage_columns = ['pct_employees_with_vaulted_cards', 'pct_employees_with_fly_spent']

    # Determine whether to use avg or median $FLY per employee by comparing global medians
    medians = df[value_columns].median(numeric_only=True)
    avg_median = medians.get('avg_fly_balance_per_employee', float('nan'))
    median_median = medians.get('median_fly_balance_per_employee', float('nan'))

    chosen_fly_metric = None  # Will store the chosen metric's name
    chosen_fly_label = None   # Will store the chosen metric's layman label

    if pd.notnull(avg_median) and pd.notnull(median_median):
        if avg_median >= median_median:
            # Keep avg_fly_balance_per_employee, remove median_fly_balance_per_employee
            if 'median_fly_balance_per_employee' in value_columns:
                value_columns.remove('median_fly_balance_per_employee')
            chosen_fly_metric = 'avg_fly_balance_per_employee'
        else:
            # Keep median_fly_balance_per_employee, remove avg_fly_balance_per_employee
            if 'avg_fly_balance_per_employee' in value_columns:
                value_columns.remove('avg_fly_balance_per_employee')
            chosen_fly_metric = 'median_fly_balance_per_employee'
    else:
        # Handle NaN cases
        if pd.isnull(avg_median) and pd.notnull(median_median):
            # Only median is valid
            if 'avg_fly_balance_per_employee' in value_columns:
                value_columns.remove('avg_fly_balance_per_employee')
            chosen_fly_metric = 'median_fly_balance_per_employee'
        elif pd.isnull(median_median) and pd.notnull(avg_median):
            # Only avg is valid
            if 'median_fly_balance_per_employee' in value_columns:
                value_columns.remove('median_fly_balance_per_employee')
            chosen_fly_metric = 'avg_fly_balance_per_employee'
        else:
            # Both NaN - remove both
            if 'avg_fly_balance_per_employee' in value_columns:
                value_columns.remove('avg_fly_balance_per_employee')
            if 'median_fly_balance_per_employee' in value_columns:
                value_columns.remove('median_fly_balance_per_employee')
            chosen_fly_metric = None

    # Recompute medians after possibly removing one loyalty metric
    medians = df[value_columns].median(numeric_only=True)

    # If chosen_fly_metric is determined, get its layman label
    if chosen_fly_metric:
        chosen_fly_label = column_mapping[chosen_fly_metric]

    # Authenticate Gmail API
    service = get_authenticated_service()
    sender_email = 'maddie.weber@blackbird.xyz'
    recipient_email = 'maddie.weber@blackbird.xyz'  # Send all attachments to Maddie

    all_attachments = []

    # Process each restaurant
    for _, row in df.iterrows():
        restaurant_name = row.get('restaurant_name', 'Restaurant Team')
        if pd.isnull(restaurant_name) or restaurant_name.strip() == '':
            restaurant_name = 'Restaurant Team'

        highlight_values = {}
        for col in value_columns:
            val = row[col]
            if (
                pd.notnull(val) and
                col in medians.index and
                pd.notnull(medians[col]) and
                val >= medians[col]
            ):
                # Special condition: if pct_employees_with_fly_spent is 0%, exclude it
                if col == 'pct_employees_with_fly_spent' and val == 0:
                    continue

                # Formatting
                if col in percentage_columns and pd.notnull(val):
                    # Format as XX%
                    val = f"{int(round(val))}%"
                else:
                    if isinstance(val, float) and val.is_integer():
                        val = int(val)

                highlight_values[col] = val

        # If fewer than 2 metrics meet/exceed the median, we won't show the highlight table
        show_stats = len(highlight_values) >= 2

        # Start building the email
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

        # Always show the chosen $FLY metric if it exists
        if chosen_fly_metric is not None:
            chosen_val = row[chosen_fly_metric]
            # Format chosen_val similarly
            if isinstance(chosen_val, float) and chosen_val.is_integer():
                chosen_val = int(chosen_val)
            email_body += f"""
            <p>Additionally, your employees currently hold a {chosen_fly_label}: <strong>{chosen_val}</strong></p>
            """

        email_body += f"""
            <p>Thank you once again for your continued support. 
            Wishing you wonderful holidays and an even brighter 2025!</p>

            <p>Warmly,<br>
            The Blackbird Team</p>
        </body>
        </html>
        """

        subject = f"Year-End Greetings from Blackbird: Looking Forward to 2025"
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

        all_attachments.append(part)

    # Send one email to Maddie with all attachments
    if all_attachments:
        subject = "All Personalized Restaurant Emails"
        body = """Hi Maddie,

Please find attached all the personalized email drafts for each of our SF launch partners.
These emails include holiday well-wishes, enthusiasm for 2025, and highlight metrics that meet or exceed the median. 
We've updated the wording for vaulted cards and replaced "loyalty points" with "$FLY".
Also, we are always showing the chosen $FLY metric (average or median) that employees have accumulated, 
regardless of whether it meets the median threshold or not.

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
