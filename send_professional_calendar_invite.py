#!/usr/bin/env python3
"""
Send professional calendar invitation with detailed instructions
"""

import json
import logging
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
from datetime import datetime

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def create_professional_email_content():
    """Create professional email content with calendar instructions."""
    
    content = """
Dear Recipient,

I hope this email finds you well. I am pleased to share with you the comprehensive course calendar for the upcoming academic programs.

📅 Calendar Invitation Details

Please find attached the UBC Course Calendar (.ics file) containing all scheduled courses with their respective start and end dates. This calendar includes all programs: KCDS, ISPD, DASH, and NLP schedules.

How to Add This Calendar to Your Calendar Application:

For Microsoft Outlook:
1. Click on the attachment in this email "ubc_course_calendar.ics"
2. It will open a new window with the message box "The .ICS attachment can't be viewed because it contains multiple items. Do you want to add it to your calendar?"
3. Click "Yes"
4. Wait for the calendar to finish importing.

If you encouter any issues, please follow the steps below:
1. Open Outlook and go to File > Open & Export > Import/Export
2. Select "Import an iCalendar (.ics) or vCalendar file (.vcs)"
3. Browse and select the attached .ics file "ubc_course_calendar.ics"
4. Click "Next" and then "Finish"


Important Notes:
- The calendar file contains all course start and end dates
- Events are set as all-day events for easy visibility
- You can modify or delete individual events after importing
- The calendar will automatically sync across your devices if you're signed into the same account
- You will be recieving reminders for each course 14 days before the start date, 7 days before the start date, 2 days before the start date, 1 day before the start date, and 14 days after the end date.

Calendar Contents:
- All course start dates and end dates
- Program information for each course
- Location details (UBC)
- Event descriptions with course names

If you encounter any issues importing the calendar or need assistance, please don't hesitate to reach out.

Best regards,

---
This is an automated message from the Course Reminder System.
For technical support, please contact: chhavi.nayyar@ubc.ca
"""
    
    return content

def send_professional_calendar_invite():
    """Send professional calendar invitation with detailed instructions."""
    
    # Load configuration
    with open('email_config.json', 'r') as f:
        config = json.load(f)
    
    # Create message
    msg = MIMEMultipart()
    msg['From'] = config['sender_email']
    msg['To'] = config['recipients'][0]  # Send to first recipient (you)
    msg['Subject'] = "📅 UBC Course Calendar Invitation - Professional Instructions"
    
    # Add professional email content
    email_content = create_professional_email_content()
    msg.attach(MIMEText(email_content, 'plain'))
    
    # Attach the calendar file
    calendar_file = "ubc_course_calendar.ics"
    if os.path.exists(calendar_file):
        with open(calendar_file, 'rb') as f:
            attachment = MIMEBase('application', 'octet-stream')
            attachment.set_payload(f.read())
            encoders.encode_base64(attachment)
            attachment.add_header(
                'Content-Disposition',
                f'attachment; filename= {calendar_file}'
            )
            msg.attach(attachment)
        logger.info(f"✅ Attached calendar file: {calendar_file}")
    else:
        logger.error(f"❌ Calendar file not found: {calendar_file}")
        return False
    
    try:
        # Create secure connection
        context = ssl.create_default_context()
        context.check_hostname = False
        context.verify_mode = ssl.CERT_NONE
        
        with smtplib.SMTP(config['smtp_server'], config['smtp_port']) as server:
            server.starttls(context=context)
            server.login(config['sender_email'], config['sender_password'])
            
            # Send email
            text = msg.as_string()
            server.sendmail(config['sender_email'], [config['recipients'][0]], text)
            
        logger.info(f"✅ Professional calendar invitation sent to {config['recipients'][0]}")
        logger.info("📧 Email includes detailed instructions for all major calendar applications")
        return True
        
    except Exception as e:
        logger.error(f"❌ Failed to send professional calendar invitation: {e}")
        return False

def main():
    """Main function to send professional calendar invitation."""
    
    logger.info("📅 Professional Calendar Invitation System")
    logger.info("=" * 50)
    
    # Check if calendar file exists
    calendar_file = "ubc_course_calendar.ics"
    if not os.path.exists(calendar_file):
        logger.error(f"❌ Calendar file not found: {calendar_file}")
        logger.info("Please run the calendar creation script first.")
        return
    
    # Send professional email
    success = send_professional_calendar_invite()
    
    if success:
        logger.info("✅ Professional calendar invitation sent successfully!")
        logger.info("📧 Check your email for the calendar file and detailed instructions")
    else:
        logger.error("❌ Failed to send calendar invitation")

if __name__ == "__main__":
    main()
