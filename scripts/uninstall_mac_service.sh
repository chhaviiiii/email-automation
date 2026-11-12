#!/bin/bash
set -euo pipefail

LABEL="com.chhavi.email-reminders"
launchctl bootout "gui/$(id -u)/$LABEL" 2>/dev/null || true
rm -f "$HOME/Library/LaunchAgents/$LABEL.plist"
echo "Removed background service. Your events and settings are still in reminders.db."
