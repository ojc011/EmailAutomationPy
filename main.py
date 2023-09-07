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
            <a href="mailto:olivercronk@roofline.com?subject=Unsubscribe&body=Please unsubscribe me from this mailing list." style="color: red; text-decoration: none;">Unsubscribe</a>
        </div>
    </div>
    """
    msg.attach(MIMEText(email_template, "html"))

    try:
        with open("assets/hbm.png", "rb") as img:
            mime_img = MIMEImage(img.read())
            mime_img.add_header("Content-ID", "<logo>")
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
        self.root.geometry("720x1080")
        self.current_index = 0
        self.total_count = 0
        self.sent_count = 0

        self.csv_filepath = None

        self.upload_btn = tk.Button(
            self.root, text="Upload CSV", command=self.upload_csv_callback
        )
        self.upload_btn.pack(pady=10)

        self.csv_label = tk.Label(self.root, text="")
        self.csv_label.pack(pady=2.5)

        tk.Label(self.root, text="Subject:").pack(pady=5)
        self.subject_entry = tk.Entry(self.root, width=50)
        self.subject_entry.pack()

        tk.Label(self.root, text="Email Body:").pack(pady=5)
        self.message_entry = tk.Text(self.root, height=10, width=50)
        self.message_entry.pack()

        self.save_preset_btn = ttk.Button(
            self.root, text="Save Preset", command=self.save_preset, style="BW.TButton"
        )
        self.save_preset_btn.pack(pady=10)

        self.load_preset_btn = ttk.Button(
            self.root, text="Load Preset", command=self.load_preset, style="BW.TButton"
        )
        self.load_preset_btn.pack(pady=10)

        tk.Label(self.root, text="Status Log:").pack(pady=5)
        self.log_textbox = tk.Text(
            self.root,
            height=10,
            width=50,
            bg="light gray",
            fg="black",
            state=tk.DISABLED,
        )
        self.log_textbox.pack(pady=5)

        style = ttk.Style()
        style.configure("BW.TButton", foreground="white", background="black")

        self.send_btn = ttk.Button(
            self.root,
            text="Send Emails",
            command=self.send_emails_callback,
            state=tk.DISABLED,
            style="BW.TButton",
        )
        self.send_btn.pack(pady=10)

        self.stop_flag = threading.Event()

        self.stop_btn = ttk.Button(
            self.root,
            text="Stop",
            command=self.stop_process,
            state=tk.DISABLED,
            style="BW.TButton",
        )
        self.stop_btn.pack(pady=10)

        self.download_btn = ttk.Button(
            self.root,
            text="Download CSV",
            command=self.download_csv,
            state=tk.DISABLED,
            style="BW.TButton",
        )
        self.download_btn.pack(pady=10)

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
            time.sleep(random.uniform(60, 70))

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


if __name__ == "__main__":
    root = tk.Tk()
    app = EmailAutomationApp(root)
    root.mainloop()
