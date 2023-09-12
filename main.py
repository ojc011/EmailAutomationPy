import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import csv
import smtplib
from email.mime.text import MIMEText
import os
import time
import random
import threading
import shutil
import json
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import imaplib
import email
from email import policy
from bs4 import BeautifulSoup

from dotenv import load_dotenv

load_dotenv()

GMAIL_EMAIL = os.environ.get("GMAIL_EMAIL")
GMAIL_PASSWORD = os.environ.get("GMAIL_PASSWORD")


def send_email(receiver_email, subject, body):
    msg = MIMEMultipart("related")

    msg["From"] = GMAIL_EMAIL
    msg["To"] = receiver_email
    msg["Subject"] = subject

    email_template = f"""
    <div style="border: 2px solid black; padding: 20px; width: 600px; margin: 20px auto;">
    <div style="text-align: center; padding-bottom: 10px;">
        <a href="https://roofline.com">
            <img src="cid:logo" alt="Logo" style="max-width: 100px; max-height: 100px; margin-bottom: 10px;" />
        </a>
        <h1 style="margin: 0;">Hayden Building Maintenance Corporation</h1>
    </div>
        <hr>
        <div style="text-align: center; font-size: 20px; line-height: 2.0; padding: 10px 0;">
            {body}
        </div>
        <hr>
        <div style="text-align: center; font-size: 14px; line-height: 1.5;">
            <p>Oliver J. Cronk</p>
            <p>Marketing Coordinator</p>
            <p>Hayden Building Maintenance Corp.</p>
            <p>ROOFING * WATERPROOFING * RESTORATION</p>
            <p>169 Western Highway - PO Box G</p>
            <p>West Nyack, NY 10994</p>
            <p>Office: 845-353-3400</p>
            <p>Cell: 845-925-2357</p>
        </div>
        <hr>
        <div style="padding: 5px; text-align: center; margin-top: 5px;">
            <a href="https://roofline.com" style="color: blue; text-decoration: none; margin-bottom: 10px; display: block;">Visit Our Website</a>
            <a href="mailto:ocronk10@gmail.com?subject=Unsubscribe From HBM Emails&body=Please unsubscribe me from this mailing list." style="color: red; text-decoration: none;">Unsubscribe</a>
        </div>
    </div>
    """
    msg.attach(MIMEText(email_template, "html"))

    try:
        with open("assets/hbm.png", "rb") as img:
            mime_img = MIMEImage(img.read())
            mime_img.add_header("Content-ID", "<logo>")
            mime_img.add_header("Content-Disposition", 'inline; filename="HBMLogo.png"')
            msg.attach(mime_img)
    except Exception as e:
        app.update_log(f"Error loading image: {e}")
        return False

    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(GMAIL_EMAIL, GMAIL_PASSWORD)
            server.sendmail(GMAIL_EMAIL, receiver_email, msg.as_string())
        return True
    except Exception as e:
        app.update_log(f"Error sending email: {e}")
        return False

class EmailAutomationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Email Automation")
        self.root.geometry("1000x700")  # Adjust the window size as needed

        # Create a frame for buttons on the left
        self.left_frame = tk.Frame(self.root)
        self.left_frame.pack(side=tk.LEFT, padx=10, fill=tk.Y)

        # Create a frame for the log on the right
        self.right_frame = tk.Frame(self.root)
        self.right_frame.pack(side=tk.RIGHT, padx=10, fill=tk.BOTH, expand=True)

        self.current_index = 0
        self.total_count = 0
        self.sent_count = 0
        self.scan_and_update_in_progress = False

        self.csv_filepath = None

        self.upload_btn = tk.Button(
            self.left_frame, text="Upload CSV", command=self.upload_csv_callback
        )
        self.upload_btn.pack(pady=10, fill=tk.X)

        self.csv_label = tk.Label(self.left_frame, text="")
        self.csv_label.pack(pady=2.5, fill=tk.X)

        tk.Label(self.left_frame, text="Subject:").pack(pady=5)
        self.subject_entry = tk.Entry(self.left_frame, width=50)
        self.subject_entry.pack(pady=5, fill=tk.X)

        tk.Label(self.left_frame, text="Email Body:").pack(pady=5)
        self.message_entry = tk.Text(self.left_frame, height=10, width=50)
        self.message_entry.pack(pady=5, fill=tk.BOTH, expand=True)

        self.save_preset_btn = ttk.Button(
            self.left_frame,
            text="Save Preset",
            command=self.save_preset,
            style="BW.TButton",
        )
        self.save_preset_btn.pack(pady=10, fill=tk.X)

        self.load_preset_btn = ttk.Button(
            self.left_frame,
            text="Load Preset",
            command=self.load_preset,
            style="BW.TButton",
        )
        self.load_preset_btn.pack(pady=10, fill=tk.X)

        self.send_btn = ttk.Button(
            self.left_frame,
            text="Send Emails",
            command=self.send_emails_callback,
            state=tk.DISABLED,
            style="BW.TButton",
        )
        self.send_btn.pack(pady=10, fill=tk.X)

        self.stop_btn = ttk.Button(
            self.left_frame,
            text="Stop",
            command=self.stop_process,
            state=tk.DISABLED,
            style="BW.TButton",
        )
        self.stop_btn.pack(pady=10, fill=tk.X)

        self.scan_and_update_csv_btn = ttk.Button(
            self.left_frame,
            text="Scan and Update CSV",
            command=self.scan_and_update_csv,
            style="BW.TButton",
        )
        self.scan_and_update_csv_btn.pack(pady=10, fill=tk.X)

        self.download_btn = ttk.Button(
            self.left_frame,
            text="Download CSV",
            command=self.download_csv,
            state=tk.DISABLED,
            style="BW.TButton",
        )
        self.download_btn.pack(pady=10, fill=tk.X)

        # Create a text widget for the log on the right
        self.log_textbox = tk.Text(
            self.right_frame,
            height=10,
            width=50,
            bg="light gray",
            fg="black",
            state=tk.DISABLED,
        )
        self.log_textbox.pack(pady=5, fill=tk.BOTH, expand=True)

        self.style = ttk.Style()
        self.style.configure("BW.TButton", foreground="white", background="black")

        self.stop_flag = threading.Event()

    def save_preset(self):
        preset_data = {
            "subject": self.subject_entry.get(),
            "body": self.message_entry.get("1.0", tk.END),
        }
        try:
            with open("presets.json", "a") as file:
                json.dump(preset_data, file)
                file.write("\n")
        except Exception as e:
            self.update_log(f"Error saving preset: {e}")

    def load_preset(self):
        try:
            with open("presets.json", "r") as file:
                presets = [json.loads(line) for line in file]
        except FileNotFoundError:
            presets = []

        if presets:
            # Implement a GUI for the user to choose a preset
            chosen_preset = random.choice(presets)  # Replace with a GUI to choose
            self.subject_entry.delete(0, tk.END)
            self.subject_entry.insert(0, chosen_preset["subject"])
            self.message_entry.delete("1.0", tk.END)
            self.message_entry.insert("1.0", chosen_preset["body"])
        else:
            messagebox.showinfo("Info", "No presets found.")

    def upload_csv_callback(self):
        filepath = filedialog.askopenfilename()
        if filepath:
            self.csv_filepath = filepath
            self.csv_label.config(text=os.path.basename(filepath))
            self.send_btn["state"] = tk.NORMAL

    def update_log(self, message):
        self.log_textbox.config(state=tk.NORMAL)
        self.log_textbox.insert(tk.END, message + "\n")
        self.log_textbox.see(tk.END)
        self.log_textbox.config(state=tk.DISABLED)

    def send_emails(self):
        if not self.csv_filepath:
            messagebox.showerror("Error", "Please upload a CSV file!")
            return

        subject = self.subject_entry.get()
        message = self.message_entry.get("1.0", tk.END)

        try:
            with open(self.csv_filepath, "r", encoding="utf-8-sig") as csv_file:
                reader = csv.reader(csv_file)
                rows = list(reader)
        except Exception as e:
            self.update_log(f"Failed to open CSV file: {e}")
            return

        valid_email_count = sum(
            1
            for row in rows[1:]
            if row[0] and (not row[2:] or row[2].lower() not in ("skip", "1"))
        )

        self.total_count = valid_email_count
        self.sent_count = 0

        attempt_count = 0

        for index, row in enumerate(
            rows[self.current_index + 1 :], start=self.current_index + 1
        ):
            if not row[0]:
                self.update_log("Encountered an empty email. Process stopped.")
                break

            if row[2].lower() in ("1",) or row[1].lower() in ("skip",):
                reason = "Unsubscribed" if row[1].lower() == "skip" else "Already Sent"
                self.update_log(f"Skipping {row[0]}: {reason}")
                continue
            attempt_count += 1

            if self.stop_flag.is_set():
                self.update_log("Process stopped by user.")
                break

            if send_email(row[0], subject, message):
                self.sent_count += 1
                self.update_log(
                    f"Successfully sent to: {row[0]} - {self.sent_count}/{attempt_count}"
                )
                row[2] = "1"
            else:
                self.update_log(f"Failed to send to: {row[0]}")

            self.current_index = index

            # Check the stop_flag more frequently to stop the process quickly
            if self.stop_flag.is_set():
                self.update_log("Process stopped by user.")
                break

            time.sleep(random.uniform(60, 80))  # Adjust the sleep duration as needed

        with open(
            "updated_" + os.path.basename(self.csv_filepath),
            "w",
            newline="",
            encoding="utf-8-sig",
        ) as csv_file:
            writer = csv.writer(csv_file)
            writer.writerows(rows)

        self.update_log(f"Completed: Sent {self.sent_count}/{attempt_count} emails!")
        messagebox.showinfo("Completed", "Emails sent!")
        self.download_btn["state"] = tk.NORMAL
        self.send_btn["state"] = tk.NORMAL
        self.send_btn["text"] = "Send Emails"
        self.stop_btn["state"] = tk.DISABLED

    def stop_process(self):
        self.stop_flag.set()

        # If a scan-and-update operation is not already in progress, trigger it
        if not self.scan_and_update_in_progress:
            self.scan_and_update_in_progress = True

    def download_csv(self):
        filepath = "updated_" + os.path.basename(self.csv_filepath)
        if filepath:
            dest = filedialog.asksaveasfilename(
                defaultextension=".csv",
                initialfile=filepath,
                filetypes=[("CSV files", "*.csv")],
            )
            if dest:
                shutil.copy(filepath, dest)
                self.update_log(f"CSV downloaded as {dest}")

    def send_emails_callback(self):
        self.send_btn["state"] = tk.DISABLED
        self.stop_btn["state"] = tk.NORMAL

        self.stop_flag.clear()

        email_thread = threading.Thread(target=self.send_emails)
        email_thread.start()

    def scan_and_update_csv(self):
        # Set up the IMAP client and log in
        try:
            mail = imaplib.IMAP4_SSL("imap.gmail.com")
            mail.login(GMAIL_EMAIL, GMAIL_PASSWORD)
        except imaplib.IMAP4.error as e:
            self.update_log(f"Failed to login to Gmail: {e}")
            return

        # Define the search criteria
        criteria = '(OR FROM "mailer-daemon@googlemail.com" FROM "postmaster@")'

        # Initialize a list to store removed emails
        removed_emails = []

        # Search in both the inbox and sent folders
        for folder in ["Inbox", "Sent"]:
            result, _ = mail.select(folder)
            if result != "OK":
                self.update_log(f"Failed to select {folder} folder.")
                continue

            # Search for emails matching the criteria
            result, msg_nums = mail.search(None, criteria)
            if result != "OK":
                self.update_log("Failed to perform email search.")
                continue

            # Get the email ids
            email_ids = msg_nums[0].split()

            # Loop through all email ids
            for e_id in email_ids:
                # Fetch the email by id
                result, email_data = mail.fetch(e_id, "(BODY.PEEK[TEXT])")

                # If fetching was successful, parse the email
                if result == "OK":
                    msg = email.message_from_bytes(
                        email_data[0][1], policy=policy.default
                    )
                    # Extract the content and check if it contains an email address from the CSV
                    if msg.is_multipart():
                        for part in msg.iter_parts():
                            if part.get_content_type() == "text/plain":
                                content = part.get_payload(decode=True).decode("utf-8")
                            elif part.get_content_type() == "text/html":
                                content = part.get_payload(decode=True).decode("utf-8")
                                soup = BeautifulSoup(content, "html.parser")
                                content = soup.get_text()
                    else:
                        content = msg.get_payload()

                    # Check if any email from the CSV file is mentioned in the content
                    if self.csv_filepath:
                        with open(
                            self.csv_filepath, "r", encoding="utf-8-sig"
                        ) as csv_file:
                            reader = csv.reader(csv_file)
                            rows = list(reader)

                        for index, row in enumerate(rows[1:]):
                            if row[0] in content:
                                self.update_log(
                                    f"Found mention of {row[0]} in an email from mailer-daemon or postmaster. Removing from CSV."
                                )
                                removed_emails.append(row[0])
                                rows.pop(index + 1)
                                break

                        # Save the updated CSV
                        with open(
                            self.csv_filepath, "w", newline="", encoding="utf-8-sig"
                        ) as csv_file:
                            writer = csv.writer(csv_file)
                            writer.writerows(rows)

        # Log out from the mail account
        mail.logout()

        # Log the removed emails and success message
        if removed_emails:
            self.update_log(
                f"Removed {len(removed_emails)} emails from the CSV:\n"
                + "\n".join(removed_emails)
            )
        else:
            self.update_log("No emails from mailer-daemon or postmaster found.")
        self.update_log("Scan and Update CSV completed successfully.")


if __name__ == "__main__":
    root = tk.Tk()
    app = EmailAutomationApp(root)
    root.mainloop()
