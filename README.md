# 📧 Gmail Bulk Email Sender with Excel Auto-Logging

This Python script allows you to send **personalized HTML emails using Gmail SMTP** to a list of clients stored in an Excel file (`clients.xlsx`). It keeps track of which emails were sent, when, and to whom — and prevents duplicates or invalid addresses.

---

## ✅ Key Features

- ✅ Reads client data from an Excel file (`clients.xlsx`)
- 💌 Sends fully formatted **HTML emails** using Gmail SMTP
- 📬 Personalized greeting using the client’s name
- 🧠 Validates email format before sending
- 📝 Auto-updates Excel with:
  - ✅ Send status
  - 🕒 Timestamp of send
- ❌ Skips:
  - Already sent emails (`Status_Email = ✅ Sent`)
  - Invalid email addresses
- ⛔ **Prevents file corruption:** Will not run if Excel is open
- 🛑 Stops automatically after 5 consecutive failures
- 🧾 Displays real-time logs for every attempt
- 🧷 Supports fallback plain-text message for older email clients

---

## 🗃️ Excel File Format

Ensure your `clients.xlsx` file has the following structure:

| Name         | Email              | Status_Email     | Sent_At_Email_1     |
|--------------|--------------------|------------------|----------------------|
| John Doe     | john@gmail.com     |                  |                      |
| Jane Smith   | jane@example.com   | ✅ Sent           | 28-05-2025 12:30:45  |

> Columns are case-insensitive and extra columns will be ignored.

---

## ⚙️ Setup Instructions

### 1. 📦 Install Required Packages

```bash
pip install pandas openpyxl
