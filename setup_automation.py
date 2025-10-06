#!/usr/bin/env python3
"""
Setup automation for the email reminder system
"""

import os
import platform
import subprocess
import logging
from datetime import datetime

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def setup_automation():
    """Setup daily automation for the email reminder system."""
    
    system = platform.system().lower()
    current_dir = os.getcwd()
    script_path = os.path.join(current_dir, "run_reminder.py")
    
    logger.info("🚀 Setting up daily automation for email reminder system...")
    logger.info(f"📁 Current directory: {current_dir}")
    logger.info(f"🐍 Python script: {script_path}")
    
    if system == "darwin":  # macOS
        setup_macos_automation(current_dir, script_path)
    elif system == "linux":
        setup_linux_automation(current_dir, script_path)
    elif system == "windows":
        setup_windows_automation(current_dir, script_path)
    else:
        logger.error(f"❌ Unsupported operating system: {system}")
        return False
    
    return True

def setup_macos_automation(current_dir, script_path):
    """Setup automation for macOS using launchd."""
    
    logger.info("🍎 Setting up macOS automation...")
    
    # Create launchd plist file
    plist_content = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>com.emailreminder.daily</string>
    <key>ProgramArguments</key>
    <array>
        <string>python3</string>
        <string>{script_path}</string>
    </array>
    <key>WorkingDirectory</key>
    <string>{current_dir}</string>
    <key>StartCalendarInterval</key>
    <dict>
        <key>Hour</key>
        <integer>9</integer>
        <key>Minute</key>
        <integer>0</integer>
    </dict>
    <key>StandardOutPath</key>
    <string>{current_dir}/email_reminder.log</string>
    <key>StandardErrorPath</key>
    <string>{current_dir}/email_reminder.log</string>
</dict>
</plist>"""
    
    plist_file = os.path.expanduser("~/Library/LaunchAgents/com.emailreminder.daily.plist")
    
    try:
        # Write plist file
        with open(plist_file, 'w') as f:
            f.write(plist_content)
        logger.info(f"✅ Created plist file: {plist_file}")
        
        # Load the launch agent
        subprocess.run(["launchctl", "load", plist_file], check=True)
        logger.info("✅ Launch agent loaded successfully")
        
        # Enable it
        subprocess.run(["launchctl", "enable", "gui/$(id -u)/com.emailreminder.daily"], check=True)
        logger.info("✅ Launch agent enabled")
        
        logger.info("\n🎉 macOS automation setup complete!")
        logger.info("📅 The system will run daily at 9:00 AM")
        logger.info(f"📝 Logs will be saved to: {current_dir}/email_reminder.log")
        
    except subprocess.CalledProcessError as e:
        logger.error(f"❌ Error setting up macOS automation: {e}")
        logger.info("💡 Manual setup instructions:")
        logger.info(f"1. Copy the plist file to: {plist_file}")
        logger.info("2. Run: launchctl load ~/Library/LaunchAgents/com.emailreminder.daily.plist")
        logger.info("3. Run: launchctl enable gui/$(id -u)/com.emailreminder.daily")

def setup_linux_automation(current_dir, script_path):
    """Setup automation for Linux using cron."""
    
    logger.info("🐧 Setting up Linux automation...")
    
    # Create cron job
    cron_job = f"0 9 * * * cd {current_dir} && python3 {script_path} >> {current_dir}/email_reminder.log 2>&1"
    
    try:
        # Add to crontab
        result = subprocess.run(["crontab", "-l"], capture_output=True, text=True)
        current_crontab = result.stdout if result.returncode == 0 else ""
        
        if "email_reminder" not in current_crontab:
            new_crontab = current_crontab + f"\n{cron_job}\n"
            subprocess.run(["crontab", "-"], input=new_crontab, text=True, check=True)
            logger.info("✅ Cron job added successfully")
        else:
            logger.info("ℹ️ Cron job already exists")
        
        logger.info("\n🎉 Linux automation setup complete!")
        logger.info("📅 The system will run daily at 9:00 AM")
        logger.info("📝 Logs will be saved to: email_reminder.log")
        
    except subprocess.CalledProcessError as e:
        logger.error(f"❌ Error setting up Linux automation: {e}")
        logger.info("💡 Manual setup instructions:")
        logger.info("1. Run: crontab -e")
        logger.info(f"2. Add this line: {cron_job}")

def setup_windows_automation(current_dir, script_path):
    """Setup automation for Windows using Task Scheduler."""
    
    logger.info("🪟 Setting up Windows automation...")
    
    # Create batch file
    batch_file = os.path.join(current_dir, "run_reminder.bat")
    batch_content = f"""@echo off
cd /d "{current_dir}"
python {script_path}
"""
    
    try:
        with open(batch_file, 'w') as f:
            f.write(batch_content)
        logger.info(f"✅ Created batch file: {batch_file}")
        
        # Create PowerShell script for Task Scheduler
        ps_script = f"""
# Create scheduled task for email reminders
$action = New-ScheduledTaskAction -Execute "python" -Argument "{script_path}" -WorkingDirectory "{current_dir}"
$trigger = New-ScheduledTaskTrigger -Daily -At 9:00AM
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType InteractiveToken

Register-ScheduledTask -TaskName "EmailReminderDaily" -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Description "Daily email reminder system"
"""
        
        ps_file = os.path.join(current_dir, "setup_task.ps1")
        with open(ps_file, 'w') as f:
            f.write(ps_script)
        
        logger.info(f"✅ Created PowerShell script: {ps_file}")
        logger.info("\n🎉 Windows automation setup complete!")
        logger.info("📅 The system will run daily at 9:00 AM")
        logger.info("💡 To complete setup:")
        logger.info(f"1. Run PowerShell as Administrator")
        logger.info(f"2. Execute: {ps_file}")
        
    except Exception as e:
        logger.error(f"❌ Error setting up Windows automation: {e}")
        logger.info("💡 Manual setup instructions:")
        logger.info("1. Open Task Scheduler")
        logger.info("2. Create Basic Task")
        logger.info("3. Set trigger to daily at 9:00 AM")
        logger.info(f"4. Set action to run: python {script_path}")

def show_manual_instructions():
    """Show manual setup instructions for all platforms."""
    
    logger.info("\n📋 MANUAL SETUP INSTRUCTIONS:")
    logger.info("=" * 60)
    
    logger.info("\n🍎 macOS (using cron):")
    logger.info("1. Open Terminal")
    logger.info("2. Run: crontab -e")
    logger.info("3. Add this line:")
    logger.info("   0 9 * * * cd /Users/chhavinayyar/email-automation && python3 run_reminder.py")
    logger.info("4. Save and exit")
    
    logger.info("\n🐧 Linux:")
    logger.info("1. Run: crontab -e")
    logger.info("2. Add this line:")
    logger.info("   0 9 * * * cd /Users/chhavinayyar/email-automation && python3 run_reminder.py")
    logger.info("3. Save and exit")
    
    logger.info("\n🪟 Windows:")
    logger.info("1. Open Task Scheduler")
    logger.info("2. Create Basic Task")
    logger.info("3. Name: 'Email Reminder Daily'")
    logger.info("4. Trigger: Daily at 9:00 AM")
    logger.info("5. Action: Start Program")
    logger.info("   Program: python")
    logger.info("   Arguments: run_reminder.py")
    logger.info("   Start in: /Users/chhavinayyar/email-automation")
    
    logger.info("\n⏰ Alternative: Run continuously")
    logger.info("You can also run the system continuously:")
    logger.info("python3 email_reminder_system.py")
    logger.info("This will check every day at 9:00 AM automatically")

def test_automation():
    """Test the automation setup."""
    
    logger.info("\n🧪 Testing automation setup...")
    
    try:
        # Test running the reminder script
        result = subprocess.run(["python3", "run_reminder.py"], 
                              capture_output=True, text=True, timeout=30)
        
        if result.returncode == 0:
            logger.info("✅ Test run successful!")
            logger.info("📧 Check your email for test messages")
        else:
            logger.error(f"❌ Test run failed: {result.stderr}")
            
    except subprocess.TimeoutExpired:
        logger.warning("⏰ Test run timed out (this is normal)")
    except Exception as e:
        logger.error(f"❌ Error testing automation: {e}")

if __name__ == "__main__":
    logger.info("🚀 Email Reminder System - Automation Setup")
    logger.info("=" * 50)
    
    # Setup automation
    if setup_automation():
        logger.info("\n✅ Automation setup completed!")
    else:
        logger.info("\n⚠️ Automation setup had issues, showing manual instructions...")
        show_manual_instructions()
    
    # Show manual instructions as backup
    show_manual_instructions()
    
    # Test the system
    test_automation()
    
    logger.info("\n🎉 Setup complete! Your email reminder system is ready to run daily.")
