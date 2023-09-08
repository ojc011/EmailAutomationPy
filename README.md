# Email Automation Tool

The Email Automation Tool is a powerful Python application designed to simplify and streamline the process of conducting email campaigns. Built with a user-friendly Tkinter-based graphical interface, this tool empowers users to effortlessly manage email distributions with efficiency and precision.

With features like customizable email composition, preset management, CSV integration for contact lists, real-time logging, and the ability to pause and resume email sending at any point, it's a comprehensive solution for email marketers and business professionals. Additionally, the tool offers an automatic email scanning and updating feature to maintain the integrity of your contact list. Whether you're sending newsletters, promotions, or important announcements, the Email Automation Tool offers a seamless and efficient approach to managing email campaigns.

## Table of Contents

- [Introduction](#introduction)
- [Features](#features)
- [Getting Started](#getting-started)
  - [Prerequisites](#prerequisites)
  - [Installation](#installation)
- [Usage](#usage)
- [Presets](#presets)
- [CSV Operations](#csv-operations)
- [Contributing](#contributing)
- [License](#license)

## Introduction

The Email Automation Tool is a Python application designed to simplify the process of sending mass emails using a CSV contact list. It provides a user-friendly interface for composing email campaigns, managing email presets, and tracking the status of sent emails.

Key features include:

- **Email Composition:** Easily compose email campaigns with customizable subject lines and message bodies.
- **Preset Management:** Save and load email presets for reuse in future campaigns.
- **CSV Integration:** Upload a CSV contact list with recipient email addresses and track email delivery status.
- **Stop and Resume:** Pause and resume email sending at any time.
- **Email Scan and Update:** Automatically remove unsubscribed or problematic email addresses from your contact list.
- **Logging:** Real-time logging of email sending progress and errors.

## Features

- **GUI Interface:** The application uses Tkinter for its graphical user interface, making it user-friendly and accessible.
- **Email Templates:** You can create and save email templates for different campaign types.
- **CSV Management:** The tool supports CSV files for managing your email recipient list.
- **Email Sending:** Send emails using your Gmail account with SMTP.
- **Logging:** Real-time logging of email sending progress and any errors encountered.
- **Presets:** Save and load email presets to streamline your email campaign setup.
- **Stop and Resume:** You can stop and resume email sending at any point.
- **Email Scan and Update:** Automatically scan and update your CSV contact list to remove problematic email addresses.

## Getting Started

### Prerequisites

Before using the Email Automation Tool, ensure you have the following:

- Python 3.x installed on your system.
- A Gmail account with the credentials (GMAIL_EMAIL and GMAIL_PASSWORD) configured in your environment.

### Installation

1. Clone the repository to your local machine:

```bash
git clone https://github.com/yourusername/email-automation-tool.git
```

**Navigate to the project directory:**

```bash
cd email-automation-tool
```

**Install the required Python packages:**
`pip install -r requirements.txt`

## Usage

1. **Run the application** by executing `main.py`.

2. Use the **GUI** to compose your email campaign, set up email presets, and upload a CSV contact list.

3. Click the **"Send Emails"** button to start sending emails. You can pause and resume the process as needed.

4. Monitor the real-time logging for progress updates and any errors encountered.

5. After completion, you can **download the updated CSV contact list** and review the sent emails.

## Presets

- **Email presets** allow you to save and load predefined email subjects and message bodies for reuse in different campaigns.
- To save a preset, enter the subject and message in the respective fields, then click the **"Save Preset"** button.
- To load a preset, click the **"Load Preset"** button, select the desired preset, and it will populate the subject and message fields.

## CSV Operations

- The tool supports CSV files for managing your email recipient list.
- You can **upload a CSV contact list** with columns for email addresses, unsubscribe status, and more.
- The tool will automatically skip unsubscribed addresses and mark sent emails in the CSV.
- You can also **scan and update the CSV** to remove email addresses found in emails from mailer-daemon or postmaster.

## Contributing

Contributions to this project are welcome. If you find a bug or have an idea for a new feature, please open an issue or submit a pull request.

## License
