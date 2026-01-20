import os
import base64
import pickle
from io import BytesIO
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders

import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from PIL import Image, ImageDraw, ImageFont

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build


# =========================
# CONFIGURATION
# =========================
SENDER_EMAIL = "anjali.rathore@coolbootsmedia.co"
RECIPIENT_EMAIL = os.getenv('RECIPIENT_EMAIL', "ankit.k@coolbootsmedia.com")
EMAIL_SUBJECT = "📊 Automated Report - Registration Template"
EMAIL_BODY = "📊 Here is the Automated Report 4\n\nPlease find the registration template attached."
EXCEL_FILENAME = "Registration_Template.xlsx"
SHEET_NAME = "Template"

# Gmail API scope
SCOPES = ['https://www.googleapis.com/auth/gmail.send']


# =========================
# CROSS-PLATFORM EXCEL TO IMAGE
# =========================
def excel_to_image_cross_platform(excel_path, sheet_name):
    """
    Convert Excel sheet to PNG image (cross-platform version)
    Uses openpyxl to read data and PIL to create image
    """
    print(f"📂 Opening Excel file: {excel_path}")
    
    # Load workbook
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    
    try:
        ws = wb[sheet_name]
        print(f"✅ Opened sheet: '{sheet_name}'")
    except KeyError:
        print(f"⚠️ Sheet '{sheet_name}' not found. Using first sheet.")
        ws = wb.active
    
    # Get used range
    max_row = ws.max_row
    max_col = ws.max_column
    
    # Read data into list
    data = []
    for row in ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
        data.append([str(cell.value) if cell.value is not None else '' for cell in row])
    
    # Calculate image dimensions
    cell_width = 120
    cell_height = 35
    padding = 10
    
    img_width = max_col * cell_width + padding * 2
    img_height = max_row * cell_height + padding * 2
    
    # Create image
    img = Image.new('RGB', (img_width, img_height), color='white')
    draw = ImageDraw.Draw(img)
    
    # Try to use a better font, fall back to default if not available
    try:
        font = ImageFont.truetype("arial.ttf", 12)
        font_bold = ImageFont.truetype("arialbd.ttf", 12)
    except:
        font = ImageFont.load_default()
        font_bold = font
    
    # Draw cells
    for row_idx, row_data in enumerate(data):
        for col_idx, cell_value in enumerate(row_data):
            x = col_idx * cell_width + padding
            y = row_idx * cell_height + padding
            
            # Get cell style from Excel
            excel_cell = ws.cell(row=row_idx + 1, column=col_idx + 1)
            
            # Background color
            fill = excel_cell.fill
            if fill and fill.start_color and fill.start_color.rgb:
                rgb = fill.start_color.rgb
                if len(rgb) == 8:  # ARGB format
                    rgb = rgb[2:]  # Remove alpha
                bg_color = f'#{rgb}'
                draw.rectangle([x, y, x + cell_width, y + cell_height], 
                             fill=bg_color, outline='black', width=1)
            else:
                draw.rectangle([x, y, x + cell_width, y + cell_height], 
                             fill='white', outline='black', width=1)
            
            # Text
            is_bold = excel_cell.font and excel_cell.font.bold
            current_font = font_bold if is_bold else font
            
            # Center text in cell
            text_bbox = draw.textbbox((0, 0), cell_value, font=current_font)
            text_width = text_bbox[2] - text_bbox[0]
            text_height = text_bbox[3] - text_bbox[1]
            
            text_x = x + (cell_width - text_width) / 2
            text_y = y + (cell_height - text_height) / 2
            
            draw.text((text_x, text_y), cell_value, fill='black', font=current_font)
    
    wb.close()
    print(f"✅ Excel sheet '{sheet_name}' converted to image")
    return img


# =========================
# GMAIL AUTHENTICATION
# =========================
def authenticate_gmail():
    """Authenticate with Gmail API using OAuth 2.0"""
    creds = None
    
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("🔄 Refreshing authentication token...")
            creds.refresh(Request())
        else:
            print("🔐 First-time authentication required...")
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            # Use run_console for headless environments
            creds = flow.run_console() if os.getenv('CI') else flow.run_local_server(port=0)
        
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
        print("✅ Authentication successful!")
    else:
        print("✅ Using saved authentication")
    
    service = build('gmail', 'v1', credentials=creds)
    return service


# =========================
# EMAIL CREATION & SENDING
# =========================
def create_message_with_attachment(sender, to, subject, body_text, img):
    """Create email message with image attachment"""
    message = MIMEMultipart()
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject

    msg_body = MIMEText(body_text, 'plain')
    message.attach(msg_body)

    # Convert PIL Image to bytes
    img_byte_arr = BytesIO()
    img.save(img_byte_arr, format='PNG')
    img_byte_arr = img_byte_arr.getvalue()

    image_attachment = MIMEImage(img_byte_arr, name='registration_template.png')
    message.attach(image_attachment)

    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode('utf-8')
    return {'raw': raw_message}


def send_email(service, message):
    """Send email using Gmail API"""
    try:
        sent_message = service.users().messages().send(
            userId='me', body=message).execute()
        print(f"✅ Email sent successfully! Message ID: {sent_message['id']}")
        return sent_message
    except Exception as e:
        print(f"❌ Error sending email: {e}")
        raise


# =========================
# MAIN FUNCTION
# =========================
def send_registration_template_via_gmail(
    recipient=None,
    subject=EMAIL_SUBJECT,
    body=EMAIL_BODY
):
    """Main function to send Registration Template via Gmail"""
    print("\n📧 Starting Gmail automation...\n")
    
    recipient = recipient or RECIPIENT_EMAIL
    script_dir = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(script_dir, EXCEL_FILENAME)
    
    credentials_path = os.path.join(script_dir, 'credentials.json')
    if not os.path.exists(credentials_path):
        raise FileNotFoundError(
            f"❌ credentials.json not found!\n"
            f"   Please download it from Google Cloud Console"
        )
    
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"❌ Excel file not found: {excel_path}")
    
    print(f"📂 Found Excel file: {excel_path}\n")
    
    # Step 1: Convert Excel to Image
    print("STEP 1: Converting Excel to Image...")
    img = excel_to_image_cross_platform(excel_path, SHEET_NAME)
    
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


if __name__ == "__main__":
    send_registration_template_via_gmail()