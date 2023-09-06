import smtplib
from email.mime.text import MIMEText
import hashlib


def send_email(receiver_email, subject, body, smtp_config):
    """Send an email using the provided SMTP configuration."""
    msg = MIMEText(body, "html")
    msg["From"] = smtp_config["sender_email"]
    msg["To"] = receiver_email
    msg["Subject"] = subject

    try:
        with smtplib.SMTP(
            smtp_config["smtp_server"], smtp_config["smtp_port"]
        ) as server:
            server.starttls()
            server.login(smtp_config["sender_email"], smtp_config["password"])
            server.sendmail(
                smtp_config["sender_email"], receiver_email, msg.as_string()
            )
        return True
    except Exception as e:
        print(f"Error sending email: {e}")
        return False


def generate_unsubscribe_link(base_url, email):
    """Generate a unique unsubscribe link for the given email."""
    email_hash = hashlib.md5(email.encode()).hexdigest()
    return f"{base_url}/unsubscribe/{email_hash}"
