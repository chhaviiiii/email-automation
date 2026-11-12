#!/usr/bin/env python3
"""
Simple script to run the email reminder system once.
Use this for testing or manual execution.
"""

from email_reminder_system import EmailReminderSystem
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def main():
    """Run the email reminder system once."""
    logger.info("Starting email reminder check...")
    
    try:
        # Create and run the reminder system
        reminder_system = EmailReminderSystem()
        reminder_system.run_daily_check()
        logger.info("Email reminder check completed successfully!")
        
    except Exception as e:
        logger.error(f"Error running reminder system: {e}")
        raise

if __name__ == "__main__":
    main()
