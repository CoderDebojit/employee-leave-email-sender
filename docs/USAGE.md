# 📖 Usage Guide

This document explains how to configure and use **Employee Leave Email Sender**, a desktop automation application that reads employee leave information from an Excel workbook and automatically notifies reporting managers via email.

The project was developed as part of my automation initiatives while working as an **Automation Lead at Tata Consultancy Services (TCS)**. Its primary goal is to eliminate manual effort involved in sending leave notification emails and provide a consistent, repeatable workflow for daily team communication.

---

# 📋 Project Overview

The application automates the complete leave notification process by:

* Reading employee and leave information from a centralized Excel tracker.
* Identifying employees who are currently on leave.
* Preparing personalized leave notification emails.
* Sending notifications using either:

  * **Default Mail Client (Draft Opener)**, or
  * **SMTP (Background Send)**.

Instead of manually drafting multiple emails every day, managers can notify the appropriate stakeholders with just a few clicks.

---

# ✅ Prerequisites

Before running the application, ensure the following requirements are met.

| Requirement      | Details                                   |
| ---------------- | ----------------------------------------- |
| Operating System | Windows 10 / Windows 11                   |
| Python           | Python 3.8 or later                       |
| Dependencies     | Install using `requirements.txt`          |
| Excel            | Microsoft Excel (.xlsx)                   |
| SMTP Backend     | Google App Password (if using Gmail SMTP) |

Install the required packages:

```bash
pip install -r requirements.txt
```

---

# 📁 Excel Workbook Structure

The application expects an Excel workbook containing **two worksheets**.

---

## Sheet 1 — Employee Master

This sheet contains employee information along with reporting manager details.

| Column        | Description             |
| ------------- | ----------------------- |
| Name          | Employee Full Name      |
| Email         | Employee Email Address  |
| Team          | Team or Department      |
| Manager Name  | Reporting Manager Name  |
| Manager Email | Reporting Manager Email |
| CC1           | Optional CC Recipient   |
| CC2           | Optional CC Recipient   |
| CC3           | Optional CC Recipient   |

> **Note:** The employee name must exactly match the corresponding name in the **Leave Tracker** sheet.

---

## Sheet 2 — Leave Tracker

This sheet stores daily leave records.

| Column        | Description         |
| ------------- | ------------------- |
| Employee Name | Employee Name       |
| Leave Type    | Planned / Unplanned |
| Start Date    | Leave Start Date    |
| End Date      | Leave End Date      |
| Reason        | Leave Reason        |

### Date Format

All dates must follow the format:

```text
DD-MM-YYYY
```

Example:

```text
15-07-2026
```

---

# 🚀 Workflow

Using the application involves four simple steps.

---

## Step 1 — Generate the Excel Template

Launch the application and click:

```
Generate Template
```

A sample workbook named

```
Leave_Tracker_Template.xlsx
```

will be created automatically.

Use this workbook as your organization's leave tracker.

---

## Step 2 — Populate Employee Information

Fill the **Employee Master** worksheet with:

* Employee names
* Employee email addresses
* Team names
* Reporting manager information
* Optional CC recipients

Ensure employee names remain consistent across both worksheets.

---

## Step 3 — Update Daily Leave Records

Whenever an employee applies for leave, add a new row in the **Leave Tracker** sheet.

Include:

* Employee Name
* Leave Type
* Start Date
* End Date
* Reason

The application automatically detects employees who are currently on leave based on today's date.

---

## Step 4 — Configure Email Backend

Open the **Settings** tab and choose your preferred email delivery method.

### Option A — Default Mail Client (Draft Opener)

Recommended when you want to review emails before sending.

* No email password required
* Opens a pre-filled draft
* Uses your default mail application
* Manual review before sending

---

### Option B — SMTP (Background Send)

Recommended for fully automated email delivery.

* Sends emails directly
* Supports rich HTML email formatting
* No user interaction required
* Requires SMTP credentials

After configuring the settings, save them and return to the **Send Emails** tab.

Browse to your Excel workbook and click:

```
Send Emails Today
```

---

# 📧 Email Backend Comparison

| Feature                   | Default Mail Client (Draft Opener) | SMTP (Background Send)   |
| ------------------------- | ---------------------------------- | ------------------------ |
| Password Required         | ❌ No                               | ✅ Yes                    |
| Email Review              | ✅ Opens Draft                      | ❌ Sends Immediately      |
| HTML Email Support        | ❌ Plain Text                       | ✅ Rich HTML              |
| Cross Platform            | Depends on Default Mail App        | ✅ Yes                    |
| Silent Background Sending | ❌ No                               | ✅ Yes                    |
| Best Use Case             | Manual Verification                | Fully Automated Workflow |

---

# 🔐 Gmail SMTP Configuration

If you are using Gmail as your SMTP server:

1. Enable **Two-Factor Authentication**.
2. Open your Google Account.
3. Navigate to

```
Security
```

4. Open

```
App Passwords
```

5. Generate a new App Password.

Use the generated password instead of your Gmail account password.

Example configuration:

```text
SMTP Host : smtp.gmail.com
Port      : 587
TLS        : Enabled
```

---

# 🛠️ Build Executable

Create a standalone Windows executable using PyInstaller.

```bash
pyinstaller --onefile --windowed --name "LeaveMailer" leave_mailer.py
```

The generated executable will be available inside the `dist` directory.

---

# 🌿 Git Workflow

Clone the repository:

```bash
git clone https://github.com/CoderDebojit/employee-leave-email-sender.git

cd employee-leave-email-sender
```

Create a feature branch:

```bash
git checkout -b feature/improve-email-template
```

Check repository status:

```bash
git status
```

Stage changes:

```bash
git add .
```

Commit changes:

```bash
git commit -m "Improve email template layout"
```

Push changes:

```bash
git push origin feature/improve-email-template
```

---

# ⚠️ Troubleshooting

### Employee not found

Verify that the employee name matches exactly in both worksheets.

---

### Invalid Date

Ensure every date uses the following format:

```text
DD-MM-YYYY
```

---

### SMTP Authentication Failed

If using Gmail:

* Enable Two-Factor Authentication.
* Use an App Password instead of your account password.

---

### No Emails Generated

Verify that:

* The leave dates include today's date.
* Employee names match.
* Manager email addresses are available.

---

# 💡 Best Practices

* Keep a single centralized leave tracker.
* Always verify employee email addresses.
* Use SMTP mode for scheduled automation.
* Use Draft Opener mode when manual approval is required.
* Back up the Excel tracker regularly.

---

# 📬 Support

If you encounter any issues or have suggestions for improvements, feel free to open an issue in the GitHub repository.

Contributions, feedback, and feature requests are always welcome.
