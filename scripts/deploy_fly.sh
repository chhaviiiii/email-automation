#!/bin/bash
# Run tests, then deploy to Fly.io (Toronto).
set -euo pipefail
export PATH="$HOME/.fly/bin:$PATH"
cd "$(dirname "$0")/.."

.venv/bin/python -m pytest -q
fly deploy --app chhavi-reminders --ha=false --remote-only
curl --fail --silent --show-error https://chhavi-reminders.fly.dev/health
echo
echo "Deployed: https://chhavi-reminders.fly.dev"
