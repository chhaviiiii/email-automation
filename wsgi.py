"""Production entry point: `gunicorn wsgi:app --workers 1`. Must stay a single worker so only one sender runs."""

import scheduler
from app import app

scheduler.start_background_sender()
