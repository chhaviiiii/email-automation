# 📧 Email Automation System

An automated email reminder system that sends calendar invitations and course reminders based on Excel schedule data. Perfect for academic institutions, training programs, and event management.

## Features

- **Calendar Integration**: Generate and send professional calendar invites (.ics files)
- **Automated Reminders**: Send email reminders at multiple intervals (14 days, 7 days, 2 days, 1 day before, and 14 days after)
- **Excel Data Processing**: Automatically process course schedules from Excel files
- **Daily Automation**: Runs automatically via cron job or continuous monitoring
- **Professional Emails**: Well-formatted emails with detailed instructions
- **Secure Configuration**: Sensitive data excluded from version control

## Quick Start

### Prerequisites

- Python 3.7 or higher
- SMTP email account (Outlook, Gmail, etc.)
- Excel file with course schedule data

### Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/chhaviiiii/email-automation.git
   cd email-automation
   ```

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Configure email settings:**
   - Copy `email_config.json.example` to `email_config.json`
   - Update with your email credentials and recipients

4. **Add your Excel schedule file:**
   - Place your course schedule Excel file in the `data/` directory
   - Update the `excel_file` path in `email_config.json`

## 📋 Configuration

### Email Configuration (`email_config.json`)

```json
{
    "smtp_server": "smtp.mail.ubc.ca",
    "smtp_port": 587,
    "sender_email": "your_email@domain.com",
    "sender_password": "your_password",
    "recipients": [
        "recipient1@domain.com",
        "recipient2@domain.com"
    ],
    "excel_file": "data/your_schedule.xlsx",
    "date_columns": {
        "start_date": "Start Date",
        "end_date": "End Date"
    },
    "reminder_days": [14, 7, 2, 1, -14]
}
```

### Excel File Format

Your Excel file should contain columns:
- **Course Name**: Name of the course/event
- **Start Date**: Course start date (YYYY-MM-DD format)
- **End Date**: Course end date (YYYY-MM-DD format)
- **Program Name**: Program or department name

## 🛠️ Usage

### Send Calendar Invitations

```bash
python3 send_professional_calendar_invite.py
```

This will:
- Generate a comprehensive calendar file (.ics) from your Excel data
- Send professional emails with detailed instructions
- Include instructions for Google Calendar, Outlook, Apple Calendar, etc.

### Run Reminder System

**One-time check:**
```bash
python3 run_reminder.py
```

**Continuous monitoring:**
```bash
python3 email_reminder_system.py
```

### Set Up Daily Automation

```bash
python3 setup_automation.py
```

This will:
- Set up daily cron job (runs at 9:00 AM)
- Configure automatic reminder checking
- Set up logging

## 📅 Reminder Schedule

The system automatically sends reminders:

| Timing | Description |
|--------|-------------|
| 14 days before | Early notification for course preparation |
| 7 days before | Week-before reminder |
| 2 days before | Final preparation reminder |
| 1 day before | Last-minute reminder |
| 14 days after | Follow-up reminder |

## 📧 Email Features

### Calendar Invitations
- Professional email formatting
- Detailed instructions for all major calendar applications
- Complete course schedule with start/end dates
- Easy import process for recipients

### Reminder Emails
- Course-specific information
- Program details
- Start and end dates
- Professional formatting

## 🔧 File Structure

```
email-automation/
├── .gitignore                           # Excludes sensitive files
├── create_calendar_invite.py            # Calendar creation functionality
├── email_reminder_system.py             # Main reminder system
├── run_reminder.py                      # Simple reminder runner
├── send_professional_calendar_invite.py # Professional calendar emails
├── setup_automation.py                  # Automation setup
├── data/                                # Course schedule Excel files
│   └── your_schedule.xlsx
├── README.md                            # This file
└── requirements.txt                     # Python dependencies
```

## 🔒 Security

- **Configuration files** with sensitive data are excluded from version control
- **Log files** are excluded to prevent information leakage
- **Generated files** are excluded as they can be regenerated
- **Python cache files** are excluded

## 📝 Logging

All system activity is logged to `email_reminder.log`:
- Reminder sending status
- Error messages
- System operations
- Daily check results

## 🛠️ Troubleshooting

### Common Issues

1. **SMTP Authentication Error**
   - Check email credentials in `email_config.json`
   - Ensure 2-factor authentication is properly configured
   - Verify SMTP server settings

2. **Excel File Not Found**
   - Check file path in `email_config.json`
   - Ensure Excel file exists in `data/` directory
   - Verify file permissions

3. **No Reminders Sent**
   - Check if courses have valid start/end dates
   - Verify reminder timing (system only sends on specific days)
   - Check log file for error messages

### Manual Testing

Test the system manually:
```bash
python3 run_reminder.py
```

Check logs:
```bash
tail -f email_reminder.log
```

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## 📄 License

This project is open source and available under the MIT License.

## 📞 Support

For issues and questions:
- Check the troubleshooting section
- Review log files for error messages
- Create an issue in the GitHub repository

---

**Note**: This system is designed for educational and professional use. Ensure compliance with your organization's email policies and data protection regulations.
