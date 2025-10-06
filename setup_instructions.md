# Email Reminder System Setup Instructions

## Overview
This automated email reminder system sends email notifications for scheduled events based on your Excel data. It sends reminders at:
- 2 weeks before start date
- 1 week before start date  
- 2 days before start date
- 1 day before start date
- 2 weeks after end date

## Prerequisites
- Python 3.7 or higher
- Outlook account (or other SMTP email service)
- Excel file with schedule data

## Installation

1. **Install Python dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Configure email settings:**
   - The system will create a `email_config.json` file on first run
   - Edit this file with your email credentials:
   ```json
   {
       "smtp_server": "smtp.office365.com",
       "smtp_port": 587,
       "sender_email": "your_email@outlook.com",
       "sender_password": "your_password",
       "recipients": ["recipient1@example.com", "recipient2@example.com"],
       "excel_file": "data/KCDS, ISPD, DASH, NLP schedules.xlsx",
       "date_columns": {
           "start_date": "Start Date",
           "end_date": "End Date"
       },
       "reminder_days": [14, 7, 2, 1, -14]
   }
   ```

## Outlook Setup

1. **Use your regular Outlook password** (no app password needed for basic Outlook)
2. **If you have 2FA enabled**, you may need to:
   - Use an app password, or
   - Enable "Less secure app access" in your Microsoft account settings
3. **Alternative SMTP settings for Outlook:**
   - Server: `smtp.office365.com` (recommended)
   - Port: `587` (TLS) or `465` (SSL)
   - Authentication: Required

## Excel File Format

Your Excel file should contain columns for:
- Event names (any column name)
- Start dates (configured in `date_columns.start_date`)
- End dates (configured in `date_columns.end_date`)

The system will automatically detect and read all sheets in your Excel file.

## Running the System

### Option 1: Run Once (Manual)
```bash
python email_reminder_system.py
```

### Option 2: Run as a Service (Recommended)

#### On macOS/Linux:
1. **Create a cron job:**
   ```bash
   crontab -e
   ```
   
2. **Add this line to run daily at 9 AM:**
   ```
   0 9 * * * cd /Users/chhavinayyar/email-automation && python email_reminder_system.py
   ```

#### On Windows:
1. **Use Task Scheduler:**
   - Open Task Scheduler
   - Create Basic Task
   - Set trigger to daily at 9:00 AM
   - Set action to start program: `python`
   - Add arguments: `email_reminder_system.py`
   - Set start in: `/Users/chhavinayyar/email-automation`

### Option 3: Run Continuously
The script can run continuously and check for reminders every day at 9:00 AM:
```bash
python email_reminder_system.py
```
Press Ctrl+C to stop.

## Configuration Options

### Email Settings
- `smtp_server`: Your email provider's SMTP server
- `smtp_port`: SMTP port (usually 587 for TLS)
- `sender_email`: Your email address
- `sender_password`: Your email password or app password
- `recipients`: List of email addresses to send reminders to

### Excel Settings
- `excel_file`: Path to your Excel file
- `date_columns`: Column names for start and end dates
- `reminder_days`: Days before/after to send reminders

### Customizing Reminder Days
You can modify the `reminder_days` array in the config:
- Positive numbers: Days before start date
- Negative numbers: Days after end date
- Example: `[14, 7, 2, 1, -14]` means 2 weeks before, 1 week before, 2 days before, 1 day before, and 2 weeks after

## Logging

The system creates a log file `email_reminder.log` that tracks:
- When reminders are sent
- Any errors that occur
- Daily check results

## Troubleshooting

### Common Issues:

1. **"Could not parse date" warnings:**
   - Check your Excel date format
   - The system supports common formats like YYYY-MM-DD, MM/DD/YYYY, etc.

2. **"Failed to send email" errors:**
   - Verify your Outlook email credentials
   - Check if 2FA is enabled and you're using an app password
   - Ensure SMTP settings are correct (smtp-mail.outlook.com, port 587)
   - Try enabling "Less secure app access" in Microsoft account settings

3. **"Excel file not found" error:**
   - Check the file path in `email_config.json`
   - Ensure the Excel file exists in the specified location

4. **No reminders sent:**
   - Check if today matches any reminder dates
   - Verify your Excel data has proper date columns
   - Check the log file for details

### Testing the System:
1. Modify a date in your Excel file to be 14 days from today
2. Run the script manually: `python email_reminder_system.py`
3. Check if you receive the reminder email

## Support

If you encounter issues:
1. Check the log file: `email_reminder.log`
2. Verify your configuration in `email_config.json`
3. Test with a simple Excel file first
4. Ensure your email credentials are correct
