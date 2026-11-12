"""Send email through UBC's SMTP server."""

import os
import smtplib
import ssl
from email.message import EmailMessage
from email.utils import formatdate, make_msgid


class SendError(Exception):
    pass


def get_password(settings):
    """SMTP password: SMTP_PASSWORD env var if set, otherwise the one saved on the website."""
    return os.environ.get("SMTP_PASSWORD") or settings.get("smtp_password") or None


def _ssl_context():
    try:
        import certifi
        return ssl.create_default_context(cafile=certifi.where())
    except ImportError:
        return ssl.create_default_context()


def send_email(settings, recipients, subject, body):
    if not recipients:
        raise SendError("No recipients")
    password = get_password(settings)
    if not password:
        raise SendError("No email password saved yet.")

    msg = EmailMessage()
    msg["From"] = settings["sender_email"]
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = subject
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain=settings["sender_email"].split("@")[-1])
    msg.set_content(body)

    try:
        with smtplib.SMTP(settings["smtp_server"], int(settings["smtp_port"]), timeout=30) as server:
            server.starttls(context=_ssl_context())
            server.login(settings["sender_email"], password)
            server.send_message(msg)
    except smtplib.SMTPAuthenticationError as exc:
        raise SendError("UBC rejected the email password. Check that it's your current CWL password.") from exc
    except (smtplib.SMTPException, OSError) as exc:
        raise SendError(f"Couldn't send through {settings['smtp_server']}: {exc}") from exc
    return "smtp"
