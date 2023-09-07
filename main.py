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
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

GMAIL_EMAIL = os.environ.get("GMAIL_EMAIL")
GMAIL_PASSWORD = os.environ.get("GMAIL_PASSWORD")


def send_email(receiver_email, subject, body):
    msg = MIMEMultipart("related")

    msg["From"] = GMAIL_EMAIL
    msg["To"] = receiver_email
    msg["Subject"] = subject

    email_template = f"""
    <div style="border: 2px solid black; padding: 20px; width: 500px; margin: 20px auto;">
        <div style="display: flex; align-items: center;">
            <img src="cid:logo" alt="Logo" style="max-width: 100px; max-height: 100px; margin-right: 20px;" />
            <h2 style="flex-grow: 0; text-align: center; margin: 0 auto;">Hayden Building Maintenance Corporation</h2>
        </div>
        <hr>
        <div style="text-align: center; font-size: 18px; line-height: 2.0;">
            {body}
        </div>
        <hr>
        <div style="border: 1px solid black; padding: 10px; text-align: center; margin-top: 20px;">
            <a href="https://roofline.com" style="color: blue; text-decoration: none; margin-bottom: 10px; display: block;">Visit Our Website</a>
            <a href="mailto:olivercronk@roofline.com?subject=Unsubscribe&body=Please unsubscribe me from this mailing list." style="color: red; text-decoration: none;">Unsubscribe</a>
        </div>
    </div>
    """
    msg.attach(MIMEText(email_template, "html"))


    with open("assets/hbm.png", "rb") as img:
        mime_img = MIMEImage(img.read())
        mime_img.add_header("Content-ID", "<logo>")
        msg.attach(mime_img)

    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(GMAIL_EMAIL, GMAIL_PASSWORD)
            server.sendmail(GMAIL_EMAIL, receiver_email, msg.as_string())
        return True
    except Exception as e:
        print(f"Error sending email: {e}")
        return False


class EmailAutomationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Email Automation")
        self.root.geometry("1920x1080")

        self.csv_filepath = None

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

        self.stop_flag = threading.Event()

        self.stop_btn = ttk.Button(
            self.root,
            text="Stop",
            command=self.stop_process,
            state=tk.DISABLED,
            style="BW.TButton",
        )
        self.stop_btn.pack(pady=20)

        self.download_btn = ttk.Button(
            self.root,
            text="Download CSV",
            command=self.download_csv,
            state=tk.DISABLED,
            style="BW.TButton",
        )
        self.download_btn.pack(pady=20)

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
        if not self.csv_filepath:
            messagebox.showerror("Error", "Please upload a CSV file!")
            return

        subject = self.subject_entry.get()
        message = self.message_entry.get("1.0", tk.END)

        with open(self.csv_filepath, "r", encoding="utf-8-sig") as csv_file:
            reader = csv.reader(csv_file)
            rows = list(reader)

        sent_count = 0
        for index, row in enumerate(rows[1:], start=1):
            if not row[14]:
                self.update_log("Encountered an empty email. Process stopped.")
                break

            if self.stop_flag.is_set():
                self.update_log("Process stopped by user.")
                break

            if send_email(row[14], subject, message):
                sent_count += 1
                self.update_log(
                    f"Successfully sent to: {row[14]} - {sent_count}/{len(rows[1:])}"
                )
                row[1] = "1"
            else:
                self.update_log(f"Failed to send to: {row[14]}")

            time.sleep(random.uniform(15, 20))

        with open(
            "updated_" + os.path.basename(self.csv_filepath),
            "w",
            newline="",
            encoding="utf-8-sig",
        ) as csv_file:
            writer = csv.writer(csv_file)
            writer.writerows(rows)

        self.update_log(f"Completed: Sent {sent_count}/{len(rows[1:])} emails!")
        messagebox.showinfo("Completed", "Emails sent!")
        self.download_btn["state"] = tk.NORMAL

    def stop_process(self):
        self.stop_flag.set()

    def download_csv(self):
        filepath = "updated_" + os.path.basename(self.csv_filepath)
        if filepath:
            shutil.copy(
                filepath,
                filedialog.asksaveasfilename(
                    defaultextension=".csv", filetypes=[("CSV files", "*.csv")]
                ),
            )

    def send_emails_callback(self):
        self.stop_flag.clear()
        self.send_btn["state"] = tk.DISABLED
        self.stop_btn["state"] = tk.NORMAL
        email_thread = threading.Thread(target=self.send_emails)
        email_thread.start()


if __name__ == "__main__":
    root = tk.Tk()
    app = EmailAutomationApp(root)
    root.mainloop()
