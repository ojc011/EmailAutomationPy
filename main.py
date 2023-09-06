import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import csv
import smtplib
from email.mime.text import MIMEText
import os
import time
import random
import threading

GMAIL_EMAIL = os.environ.get("GMAIL_EMAIL")
GMAIL_PASSWORD = os.environ.get("GMAIL_PASSWORD")


# Email Operations
def send_email(receiver_email, subject, body):
    email_template = """
    <div style="border: 2px solid black; padding: 20px; width: 500px; margin: 20px auto;">
        <h2 style="text-align: center;">Hayden Building Maintenance Corporation</h2>
        <hr>
        <div style="font-size: 18px; line-height: 2.0;">
            {}
        </div>
    </div>
    """.format(
        body
    )

    msg = MIMEText(email_template, "html")
    msg["From"] = GMAIL_EMAIL
    msg["To"] = receiver_email
    msg["Subject"] = subject

    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(GMAIL_EMAIL, GMAIL_PASSWORD)
            server.sendmail(GMAIL_EMAIL, receiver_email, msg.as_string())
        return True
    except Exception as e:
        print(f"Error sending email: {e}")
        return False


# GUI Operations
class EmailAutomationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Email Automation")
        self.root.geometry("800x650")

        self.csv_filepath = None
        self.stop_event = threading.Event()

        self.upload_btn = tk.Button(
            self.root, text="Upload CSV", command=self.upload_csv_callback
        )
        self.upload_btn.pack(pady=20)

        tk.Label(self.root, text="Subject:").pack(pady=10)
        self.subject_entry = tk.Entry(self.root, width=50)
        self.subject_entry.pack()

        tk.Label(self.root, text="Email Body:").pack(pady=10)
        self.message_entry = tk.Text(self.root, height=10, width=50)
        self.message_entry.pack()

        tk.Label(self.root, text="Status Log:").pack(pady=10)
        self.log_textbox = tk.Text(
            self.root,
            height=10,
            width=50,
            bg="light gray",
            fg="black",
            state=tk.DISABLED,
        )
        self.log_textbox.pack(pady=10)

        style = ttk.Style()
        style.configure("BW.TButton", foreground="white", background="black")

        self.send_btn = ttk.Button(
            self.root,
            text="Send Emails",
            command=self.send_emails_callback,
            state=tk.DISABLED,
            style="BW.TButton",
        )
        self.send_btn.pack(pady=20)

        self.stop_btn = ttk.Button(
            self.root,
            text="Stop Sending",
            command=self.stop_sending_callback,
            state=tk.DISABLED,
            style="BW.TButton",
        )
        self.stop_btn.pack(pady=10)

    def upload_csv_callback(self):
        filepath = filedialog.askopenfilename()
        if not filepath:
            return

        if not filepath.endswith(".csv"):
            messagebox.showerror("Error", "Please select a valid CSV file!")
            return

        self.csv_filepath = filepath
        self.send_btn["state"] = tk.NORMAL

    def update_log(self, message):
        self.log_textbox.config(state=tk.NORMAL)
        self.log_textbox.insert(tk.END, message + "\n")
        self.log_textbox.see(tk.END)
        self.log_textbox.config(state=tk.DISABLED)

    def send_emails(self):
        self.send_btn["state"] = tk.DISABLED
        self.stop_btn["state"] = tk.NORMAL

        if not self.csv_filepath:
            messagebox.showerror("Error", "Please upload a CSV file!")
            return

        subject = self.subject_entry.get()
        message = self.message_entry.get("1.0", tk.END)

        with open(self.csv_filepath, "r", encoding="utf-8-sig") as csv_file:
            reader = csv.reader(csv_file)
            next(reader)
            emails = [row[14] for row in reader]

        sent_count = 0
        for email in emails:
            if self.stop_event.is_set():
                self.update_log("Email sending process stopped!")
                break

            if not email:
                self.update_log("Encountered an empty email. Process stopped.")
                break

            if send_email(email, subject, message):
                sent_count += 1
                self.update_log(
                    f"Successfully sent to: {email} - {sent_count}/{len(emails)}"
                )
            else:
                self.update_log(f"Failed to send to: {email}")

            time.sleep(random.uniform(15, 20))

        self.update_log(f"Completed: Sent {sent_count}/{len(emails)} emails!")
        messagebox.showinfo("Completed", "Emails sent!")
        self.send_btn["state"] = tk.NORMAL
        self.stop_btn["state"] = tk.DISABLED
        self.stop_event.clear()

    def stop_sending_callback(self):
        self.stop_event.set()

    def send_emails_callback(self):
        if hasattr(self, "email_thread") and self.email_thread.is_alive():
            messagebox.showwarning("Warning", "An email process is already running!")
            return

        self.email_thread = threading.Thread(target=self.send_emails)
        self.email_thread.start()


if __name__ == "__main__":
    root = tk.Tk()
    app = EmailAutomationApp(root)
    root.mainloop()
