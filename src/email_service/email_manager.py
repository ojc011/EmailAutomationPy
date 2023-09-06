import smtplib
import os
from email.mime.text import MIMEText


def send_email(receiver_email, subject, body, smtp_config):
    """Send an email using the provided SMTP configuration."""
    msg = MIMEText(body, "html")
    msg["From"] = smtp_config["sender_email"]
    msg["To"] = receiver_email
    msg["Subject"] = subject

    try:
        with smtplib.SMTP_SSL(
            smtp_config["smtp_server"], smtp_config["smtp_port"]
        ) as server:
            server.login(smtp_config["sender_email"], smtp_config["password"])
            server.sendmail(
                smtp_config["sender_email"], receiver_email, msg.as_string()
            )
        return True
    except Exception as e:
        print(f"Error sending email: {e}")
        return False


import hashlib


def generate_unsubscribe_link(base_url, email):
    """Generate a unique unsubscribe link for the given email."""
    # Create a simple hash for the email
    email_hash = hashlib.md5(email.encode()).hexdigest()
    return f"{base_url}/unsubscribe/{email_hash}"


OUTLOOK_EMAIL = os.environ.get("OUTLOOK_EMAIL")
OUTLOOK_PASSWORD = os.environ.get("OUTLOOK_PASSWORD")

if __name__ == "__main__":
    SMTP_CONFIG = {
        "sender_email": OUTLOOK_EMAIL,
        "password": OUTLOOK_PASSWORD,
        "smtp_server": "smtp-mail.outlook.com",
        "smtp_port": 587,
    }

    receiver = "receiver@example.com"
    subject = "Test Email with Unsubscribe"

    # Generate the unsubscribe link
    unsubscribe_link = generate_unsubscribe_link("http://yourwebsite.com", receiver)
    body = f"""
        <p>Hello there!</p>
        <p>This is a test email.</p>
        <p>If you'd like to unsubscribe, <a href="{unsubscribe_link}">click here</a>.</p>
    """

    send_email(receiver, subject, body, SMTP_CONFIG)
