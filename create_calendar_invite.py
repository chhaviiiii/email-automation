#!/usr/bin/env python3
"""
Create calendar invitation files (.ics) for course events
"""

import pandas as pd
import json
import logging
from datetime import datetime, timedelta
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib
import ssl

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def create_ics_file(course_name, start_date, end_date, program, description=""):
    """Create an .ics calendar file for a course event."""
    
    # Generate unique ID for the event
    import uuid
    event_id = str(uuid.uuid4())
    
    # Format dates for ICS
    start_dt = datetime.strptime(str(start_date), '%Y-%m-%d') if isinstance(start_date, str) else start_date
    end_dt = datetime.strptime(str(end_date), '%Y-%m-%d') if isinstance(end_date, str) else end_date
    
    # ICS format requires specific date format
    start_ics = start_dt.strftime('%Y%m%d')
    end_ics = end_dt.strftime('%Y%m%d')
    now_ics = datetime.now().strftime('%Y%m%dT%H%M%SZ')
    
    # Create ICS content
    ics_content = f"""BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//Course Reminder System//Course Event//EN
BEGIN:VEVENT
UID:{event_id}@course-reminder-system.com
DTSTAMP:{now_ics}
DTSTART;VALUE=DATE:{start_ics}
DTEND;VALUE=DATE:{end_ics}
SUMMARY:{course_name}
DESCRIPTION:{description or f'Course: {course_name}\\nProgram: {program}\\nStart: {start_date}\\nEnd: {end_date}'}
LOCATION:UBC
STATUS:CONFIRMED
TRANSP:OPAQUE
END:VEVENT
END:VCALENDAR"""
    
    return ics_content

def create_calendar_invitations():
    """Create calendar invitations for all courses."""
    
    # Load configuration
    with open('email_config.json', 'r') as f:
        config = json.load(f)
    
    # Load Excel data
    excel_file = config['excel_file']
    excel_data = pd.read_excel(excel_file, sheet_name=None)
    
    # Combine all sheets
    all_data = []
    for sheet_name, df in excel_data.items():
        df['Sheet'] = sheet_name
        all_data.append(df)
    
    schedule_data = pd.concat(all_data, ignore_index=True)
    logger.info(f"📊 Loaded {len(schedule_data)} courses")
    
    # Create calendar directory
    calendar_dir = "calendar_invites"
    if not os.path.exists(calendar_dir):
        os.makedirs(calendar_dir)
    
    created_files = []
    
    for i, row in schedule_data.iterrows():
        try:
            course_name = row.get('Course Name', f'Course {i+1}')
            start_date = row.get('Start Date')
            end_date = row.get('End Date')
            program = row.get('Program Name', 'Unknown Program')
            
            if pd.isna(start_date) or pd.isna(end_date):
                continue
            
            # Create ICS file
            ics_content = create_ics_file(course_name, start_date, end_date, program)
            
            # Save to file
            filename = f"{calendar_dir}/{course_name.replace(' ', '_').replace('/', '_')}.ics"
            with open(filename, 'w') as f:
                f.write(ics_content)
            
            created_files.append(filename)
            logger.info(f"✅ Created calendar invite: {filename}")
            
        except Exception as e:
            logger.error(f"❌ Error creating calendar for row {i}: {e}")
    
    logger.info(f"📅 Created {len(created_files)} calendar invitation files")
    return created_files

def send_calendar_invitation_email(recipient_email, ics_file_path, course_name):
    """Send calendar invitation via email."""
    
    # Load configuration
    with open('email_config.json', 'r') as f:
        config = json.load(f)
    
    # Create message
    msg = MIMEMultipart()
    msg['From'] = config['sender_email']
    msg['To'] = recipient_email
    msg['Subject'] = f"Calendar Invitation: {course_name}"
    
    # Email content
    content = f"""
Dear Team Member,

You are invited to add this course to your calendar:

📅 Course: {course_name}

Please find the calendar invitation attached to this email. You can:
1. Open the .ics file to add it to your calendar
2. Import it into Google Calendar, Outlook, or Apple Calendar
3. The event will be automatically added to your calendar

Best regards,
Course Reminder System
"""
    
    msg.attach(MIMEText(content, 'plain'))
    
    # Attach the ICS file
    with open(ics_file_path, 'rb') as f:
        attachment = MIMEText(f.read())
        attachment.add_header('Content-Disposition', 'attachment', filename=f"{course_name}.ics")
        msg.attach(attachment)
    
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
            server.sendmail(config['sender_email'], [recipient_email], text)
            
        logger.info(f"✅ Calendar invitation sent to {recipient_email}")
        return True
        
    except Exception as e:
        logger.error(f"❌ Failed to send calendar invitation: {e}")
        return False

def send_all_calendar_invitations():
    """Send calendar invitations for all courses to all recipients."""
    
    # Load configuration
    with open('email_config.json', 'r') as f:
        config = json.load(f)
    
    # Create calendar files
    calendar_files = create_calendar_invitations()
    
    if not calendar_files:
        logger.error("❌ No calendar files created")
        return
    
    # Send to all recipients
    success_count = 0
    total_emails = len(calendar_files) * len(config['recipients'])
    
    for ics_file in calendar_files:
        course_name = os.path.basename(ics_file).replace('.ics', '').replace('_', ' ')
        
        for recipient in config['recipients']:
            if send_calendar_invitation_email(recipient, ics_file, course_name):
                success_count += 1
    
    logger.info(f"📧 Sent {success_count}/{total_emails} calendar invitations")

def main():
    """Main function to create and send calendar invitations."""
    
    logger.info("📅 Course Calendar Invitation System")
    logger.info("=" * 50)
    
    # Create calendar invitations
    calendar_files = create_calendar_invitations()
    
    if calendar_files:
        logger.info(f"✅ Created {len(calendar_files)} calendar files")
        logger.info("📁 Files saved in 'calendar_invites' directory")
        
        # Ask if user wants to send emails
        send_emails = input("\n📧 Send calendar invitations via email? (y/n): ").lower().strip()
        
        if send_emails == 'y':
            send_all_calendar_invitations()
        else:
            logger.info("📁 Calendar files created. You can manually share them.")
    
    else:
        logger.error("❌ No calendar files were created")

if __name__ == "__main__":
    main()

