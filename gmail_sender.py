import os
import base64
import pickle
from io import BytesIO
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import time
import win32com.client as win32
from PIL import ImageGrab

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build


# =========================
# CONFIGURATION
# =========================
SENDER_EMAIL = "anjali.rathore@coolbootsmedia.co"
RECIPIENT_EMAIL = "ankit.k@coolbootsmedia.com"
EMAIL_SUBJECT = "📊 Automated Report - Registration Template"
EMAIL_BODY = "📊 Here is the Automated Report 4\n\nPlease find the registration template image attached."
EXCEL_FILENAME = "Registration_Template.xlsx"
SHEET_NAME = "Template"

# Gmail API scope
SCOPES = ['https://www.googleapis.com/auth/gmail.send']


# =========================
# HELPER FUNCTIONS
# =========================
def excel_to_image(excel_path, sheet_name):
    """
    Convert Excel sheet to PIL Image object
    Returns: PIL Image
    """
    print(f"📂 Opening Excel file: {excel_path}")
    
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(excel_path, ReadOnly=True)
    
    # Open the specific sheet
    try:
        ws = wb.Worksheets(sheet_name)
        print(f"✅ Opened sheet: '{sheet_name}'")
    except:
        print(f"⚠️ Sheet '{sheet_name}' not found. Using first sheet.")
        ws = wb.Worksheets(1)

    # AutoFit columns for better appearance
    ws.UsedRange.Columns.AutoFit()
    time.sleep(0.5)

    # Copy UsedRange as picture to clipboard
    ws.UsedRange.CopyPicture(Appearance=1, Format=2)
    time.sleep(0.8)

    # Grab clipboard image
    img = ImageGrab.grabclipboard()
    
    # Close Excel
    wb.Close(False)
    excel.Quit()
    
    if img is None:
        raise RuntimeError("Clipboard image capture failed. Ensure the sheet has a visible UsedRange.")

    print(f"✅ Excel sheet '{sheet_name}' converted to image")
    return img


def authenticate_gmail():
    """
    Authenticate with Gmail API using OAuth 2.0
    Returns: Gmail API service object
    """
    creds = None
    
    # Token file stores the user's access and refresh tokens
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    
    # If there are no valid credentials, let the user log in
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("🔄 Refreshing authentication token...")
            creds.refresh(Request())
        else:
            print("🔐 First-time authentication required...")
            print("   A browser window will open for Gmail login.")
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Save credentials for future runs
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
        print("✅ Authentication successful!")
    else:
        print("✅ Using saved authentication")
    
    service = build('gmail', 'v1', credentials=creds)
    return service


def create_message_with_attachment(sender, to, subject, body_text, img):
    """
    Create email message with image attachment
    
    Args:
        sender: Sender email address
        to: Recipient email address
        subject: Email subject
        body_text: Email body text
        img: PIL Image object
    
    Returns: Base64 encoded email message
    """
    message = MIMEMultipart()
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject

    # Add email body
    msg_body = MIMEText(body_text, 'plain')
    message.attach(msg_body)

    # Convert PIL Image to bytes
    img_byte_arr = BytesIO()
    img.save(img_byte_arr, format='PNG')
    img_byte_arr = img_byte_arr.getvalue()

    # Attach image
    image_attachment = MIMEImage(img_byte_arr, name='registration_template.png')
    message.attach(image_attachment)

    # Encode the message
    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode('utf-8')
    return {'raw': raw_message}


def send_email(service, message):
    """
    Send email using Gmail API
    
    Args:
        service: Gmail API service object
        message: Encoded email message
    
    Returns: Sent message object
    """
    try:
        sent_message = service.users().messages().send(
            userId='me', body=message).execute()
        print(f"✅ Email sent successfully! Message ID: {sent_message['id']}")
        return sent_message
    except Exception as e:
        print(f"❌ Error sending email: {e}")
        raise


def send_registration_template_via_gmail(
    recipient=RECIPIENT_EMAIL,
    subject=EMAIL_SUBJECT,
    body=EMAIL_BODY
):
    """
    Main function to send Registration Template via Gmail.
    Looks for Registration_Template.xlsx in the same directory as the script.
    
    Args:
        recipient: Recipient email address
        subject: Email subject
        body: Email body text
    """
    print("\n📧 Starting Gmail automation...\n")
    
    # Get the directory where this script is located
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Construct path to the Excel file
    excel_path = os.path.join(script_dir, EXCEL_FILENAME)
    
    # Check if credentials.json exists
    credentials_path = os.path.join(script_dir, 'credentials.json')
    if not os.path.exists(credentials_path):
        raise FileNotFoundError(
            f"❌ credentials.json not found!\n"
            f"   Please download it from Google Cloud Console and place it in:\n"
            f"   {script_dir}"
        )
    
    # Check if Excel file exists
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"❌ Excel file not found: {excel_path}")
    
    print(f"📂 Found Excel file: {excel_path}\n")
    
    # Step 1: Convert Excel to Image
    print("STEP 1: Converting Excel to Image...")
    img = excel_to_image(excel_path, SHEET_NAME)
    
    # Step 2: Authenticate with Gmail
    print("\nSTEP 2: Authenticating with Gmail API...")
    service = authenticate_gmail()
    
    # Step 3: Create email message
    print("\nSTEP 3: Creating email message...")
    message = create_message_with_attachment(
        sender=SENDER_EMAIL,
        to=recipient,
        subject=subject,
        body_text=body,
        img=img
    )
    print(f"   To: {recipient}")
    print(f"   Subject: {subject}")
    
    # Step 4: Send email
    print("\nSTEP 4: Sending email...")
    send_email(service, message)
    
    print("\n🎉 Gmail automation complete!\n")


# Allow script to be run standalone
if __name__ == "__main__":
    send_registration_template_via_gmail()