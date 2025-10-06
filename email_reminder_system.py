#!/usr/bin/env python3
"""
Automated Email Reminder System
Sends email reminders for scheduled events based on Excel data.
"""

import pandas as pd
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta
import os
import json
import logging
from typing import List, Dict, Any
import schedule
import time

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('email_reminder.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class EmailReminderSystem:
    def __init__(self, config_file='email_config.json'):
        """Initialize the email reminder system."""
        self.config = self.load_config(config_file)
        self.schedule_data = None
        
    def load_config(self, config_file: str) -> Dict[str, Any]:
        """Load email configuration from JSON file."""
        if os.path.exists(config_file):
            with open(config_file, 'r') as f:
                return json.load(f)
        else:
            # Create default config template
            default_config = {
                "smtp_server": "smtp-mail.outlook.com",
                "smtp_port": 587,
                "sender_email": "your_email@outlook.com",
                "sender_password": "your_password",
                "recipients": ["recipient1@example.com", "recipient2@example.com"],
                "excel_file": "data/KCDS, ISPD, DASH, NLP schedules.xlsx",
                "date_columns": {
                    "start_date": "Start Date",
                    "end_date": "End Date"
                },
                "reminder_days": [14, 7, 2, 1, -14]  # 2 weeks, 1 week, 2 days, 1 day before, 2 weeks after
            }
            with open(config_file, 'w') as f:
                json.dump(default_config, f, indent=4)
            logger.info(f"Created default config file: {config_file}")
            return default_config
    
    def load_schedule_data(self) -> pd.DataFrame:
        """Load schedule data from Excel file."""
        try:
            excel_file = self.config['excel_file']
            if not os.path.exists(excel_file):
                raise FileNotFoundError(f"Excel file not found: {excel_file}")
            
            # Try to read all sheets
            excel_data = pd.read_excel(excel_file, sheet_name=None)
            
            # Combine all sheets into one dataframe
            all_data = []
            for sheet_name, df in excel_data.items():
                df['Sheet'] = sheet_name
                all_data.append(df)
            
            self.schedule_data = pd.concat(all_data, ignore_index=True)
            logger.info(f"Loaded {len(self.schedule_data)} records from Excel file")
            return self.schedule_data
            
        except Exception as e:
            logger.error(f"Error loading schedule data: {e}")
            raise
    
    def parse_date(self, date_str: str) -> datetime:
        """Parse date string into datetime object."""
        if pd.isna(date_str):
            return None
        
        # Try different date formats
        date_formats = [
            '%Y-%m-%d',
            '%m/%d/%Y',
            '%d/%m/%Y',
            '%Y-%m-%d %H:%M:%S',
            '%m/%d/%Y %H:%M:%S'
        ]
        
        for fmt in date_formats:
            try:
                return datetime.strptime(str(date_str), fmt)
            except ValueError:
                continue
        
        # If all formats fail, try pandas to_datetime
        try:
            return pd.to_datetime(date_str)
        except:
            logger.warning(f"Could not parse date: {date_str}")
            return None
    
    def get_reminder_dates(self, start_date: datetime, end_date: datetime = None) -> List[datetime]:
        """Calculate reminder dates based on start and end dates."""
        if not start_date:
            return []
        
        reminder_dates = []
        today = datetime.now().date()
        
        for days in self.config['reminder_days']:
            if days > 0:  # Before start date
                reminder_date = start_date.date() - timedelta(days=days)
                if reminder_date >= today:
                    reminder_dates.append(reminder_date)
            elif days < 0 and end_date:  # After end date
                reminder_date = end_date.date() - timedelta(days=days)
                if reminder_date >= today:
                    reminder_dates.append(reminder_date)
        
        return sorted(reminder_dates)
    
    def create_email_content(self, event_data: Dict[str, Any], reminder_type: str) -> str:
        """Create email content based on event data and reminder type."""
        start_date = event_data.get('start_date', 'N/A')
        end_date = event_data.get('end_date', 'N/A')
        event_name = event_data.get('event_name', 'Event')
        program = event_data.get('program', 'N/A')
        
        if reminder_type == "2_weeks_before":
            subject = f"Course Reminder: {event_name}"
            content = f"""
Dear Team,

This is a reminder that {event_name} is scheduled to start in 2 weeks.

Course Details:
- Course: {event_name}
- Program: {program}
- Start Date: {start_date}
- End Date: {end_date}
"""
        elif reminder_type == "1_week_before":
            subject = f"Course Reminder: {event_name}"
            content = f"""
Dear Team,

This is a reminder that {event_name} is scheduled to start in 1 week.

Course Details:
- Course: {event_name}
- Program: {program}
- Start Date: {start_date}
- End Date: {end_date}
"""
        elif reminder_type == "2_days_before":
            subject = f"Course Reminder: {event_name}"
            content = f"""
Dear Team,

This is a reminder that {event_name} is starting in 2 days.

Course Details:
- Course: {event_name}
- Program: {program}
- Start Date: {start_date}
- End Date: {end_date}
"""
        elif reminder_type == "1_day_before":
            subject = f"Course Reminder: {event_name}"
            content = f"""
Dear Team,

This is a reminder that {event_name} is starting tomorrow.

Course Details:
- Course: {event_name}
- Program: {program}
- Start Date: {start_date}
- End Date: {end_date}
"""
        elif reminder_type == "2_weeks_after":
            subject = f"Course Reminder: {event_name}"
            content = f"""
Dear Team,

This is a follow-up reminder for {event_name} which ended 2 weeks ago.

Course Details:
- Course: {event_name}
- Program: {program}
- Start Date: {start_date}
- End Date: {end_date}
"""
        else:
            subject = f"Course Reminder: {event_name}"
            content = f"""
Dear Team,

This is a reminder about {event_name}.

Course Details:
- Course: {event_name}
- Program: {program}
- Start Date: {start_date}
- End Date: {end_date}
"""
        
        return subject, content
    
    def send_email(self, subject: str, content: str, recipients: List[str] = None) -> bool:
        """Send email to recipients."""
        if not recipients:
            recipients = self.config['recipients']
        
        try:
            # Create message
            msg = MIMEMultipart()
            msg['From'] = self.config['sender_email']
            msg['To'] = ', '.join(recipients)
            msg['Subject'] = subject
            
            # Add body
            msg.attach(MIMEText(content, 'plain'))
            
            # Create secure connection with SSL context that handles certificate issues
            context = ssl.create_default_context()
            context.check_hostname = False
            context.verify_mode = ssl.CERT_NONE
            
            with smtplib.SMTP(self.config['smtp_server'], self.config['smtp_port']) as server:
                server.starttls(context=context)
                server.login(self.config['sender_email'], self.config['sender_password'])
                
                # Send email
                text = msg.as_string()
                server.sendmail(self.config['sender_email'], recipients, text)
                
            logger.info(f"Email sent successfully to {len(recipients)} recipients")
            return True
            
        except Exception as e:
            logger.error(f"Failed to send email: {e}")
            return False
    
    def check_and_send_reminders(self):
        """Check for events that need reminders and send emails."""
        if self.schedule_data is None:
            self.load_schedule_data()
        
        today = datetime.now().date()
        sent_count = 0
        
        for _, row in self.schedule_data.iterrows():
            try:
                # Parse dates
                start_date = self.parse_date(row.get(self.config['date_columns']['start_date']))
                end_date = self.parse_date(row.get(self.config['date_columns']['end_date']))
                
                if not start_date:
                    continue
                
                # Get reminder dates
                reminder_dates = self.get_reminder_dates(start_date, end_date)
                
                # Check if today is a reminder date
                if today in reminder_dates:
                    # Determine reminder type
                    days_diff = (start_date.date() - today).days
                    
                    if days_diff == 14:
                        reminder_type = "2_weeks_before"
                    elif days_diff == 7:
                        reminder_type = "1_week_before"
                    elif days_diff == 2:
                        reminder_type = "2_days_before"
                    elif days_diff == 1:
                        reminder_type = "1_day_before"
                    elif end_date and (today - end_date.date()).days == 14:
                        reminder_type = "2_weeks_after"
                    else:
                        continue
                    
                    # Create event data
                    event_data = {
                        'event_name': row.get('Course Name', row.get('Event Name', row.get('Name', 'Event'))),
                        'program': row.get('Program Name', 'N/A'),
                        'start_date': start_date.strftime('%Y-%m-%d'),
                        'end_date': end_date.strftime('%Y-%m-%d') if end_date else 'N/A'
                    }
                    
                    # Create and send email
                    subject, content = self.create_email_content(event_data, reminder_type)
                    
                    if self.send_email(subject, content):
                        sent_count += 1
                        logger.info(f"Sent {reminder_type} reminder for {event_data['event_name']}")
                
            except Exception as e:
                logger.error(f"Error processing row: {e}")
                continue
        
        logger.info(f"Sent {sent_count} reminders today")
        return sent_count
    
    def run_daily_check(self):
        """Run the daily reminder check."""
        logger.info("Starting daily reminder check...")
        try:
            self.load_schedule_data()
            sent_count = self.check_and_send_reminders()
            logger.info(f"Daily check completed. Sent {sent_count} reminders.")
        except Exception as e:
            logger.error(f"Error in daily check: {e}")

def main():
    """Main function to run the email reminder system."""
    # Create email reminder system
    reminder_system = EmailReminderSystem()
    
    # Run initial check
    reminder_system.run_daily_check()
    
    # Schedule daily checks
    schedule.every().day.at("09:00").do(reminder_system.run_daily_check)
    
    logger.info("Email reminder system started. Checking daily at 9:00 AM.")
    logger.info("Press Ctrl+C to stop.")
    
    # Keep the script running
    try:
        while True:
            schedule.run_pending()
            time.sleep(60)  # Check every minute
    except KeyboardInterrupt:
        logger.info("Email reminder system stopped.")

if __name__ == "__main__":
    main()
