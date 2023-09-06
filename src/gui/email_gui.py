import tkinter as tk
import os
from ..email_service.email_manager import send_email, generate_unsubscribe_link

GMAIL_EMAIL = os.environ.get("GMAIL_EMAIL")
GMAIL_PASSWORD = os.environ.get("GMAIL_PASSWORD")


def send_email_callback():
    recipient = recipient_entry.get()
    subject = subject_entry.get()
    message = message_entry.get("1.0", tk.END)

    unsubscribe_link = generate_unsubscribe_link("http://roofline.com", recipient)
    message_with_unsubscribe = (
        f"{message}\n\n<a href='{unsubscribe_link}'>Unsubscribe</a>"
    )

    SMTP_CONFIG = {
        "sender_email": GMAIL_EMAIL,
        "password": GMAIL_PASSWORD,
        "smtp_server": "smtp.gmail.com",
        "smtp_port": 587,
    }

    success = send_email(recipient, subject, message_with_unsubscribe, SMTP_CONFIG)

    if success:
        log_text.set("Email sent successfully!")
    else:
        log_text.set("Error: Could not send email.")


def main():
    global recipient_entry, subject_entry, message_entry, log_text

    window = tk.Tk()
    window.title("Email Automation")
    window.geometry("600x400")

    tk.Label(window, text="Recipient Email:").pack(pady=10)
    recipient_entry = tk.Entry(window, width=50)
    recipient_entry.pack()

    tk.Label(window, text="Subject:").pack(pady=10)
    subject_entry = tk.Entry(window, width=50)
    subject_entry.pack()

    tk.Label(window, text="Message:").pack(pady=10)
    message_entry = tk.Text(window, height=10, width=50)
    message_entry.pack()

    send_btn = tk.Button(window, text="Send Email", command=send_email_callback)
    send_btn.pack(pady=20)

    log_text = tk.StringVar()
    log_label = tk.Label(window, textvariable=log_text)
    log_label.pack(pady=10)

    window.mainloop()


if __name__ == "__main__":
    main()
