
# 🕒 Employee Weekly Timesheet Summary & Email Notification Tool

This is a **web-based tool** built using **Flask (Python)** and **JavaScript** to help HR/Admins automatically summarize employees' planned vs. actual work hours from a weekly Excel timesheet and prepare a per-employee summary (with optional email notifications).

---

## 📌 Features

* ✅ Upload weekly Excel timesheets in `.xls` or `.xlsx` format
* 🧠 Smart column mapping (auto-detects column names even if they vary)
* 📧 Generates employee email addresses if missing
* 📊 Provides a clean summary of planned vs. actual hours
* 🚀 Ready to integrate with an email service to send notifications
* 🔒 CORS enabled for frontend-backend integration

---

## 🚀 Getting Started

### 1. Clone the Repo

```bash
git clone https://github.com/your-username/timesheet-summary-tool.git
cd timesheet-summary-tool
```


## 📤 Sample Excel Format

| Employee Name | Date       | Planned Task | Planned Hours | Actual Task | Actual Hours | Employee Email                              |
| ------------- | ---------- | ------------ | ------------- | ----------- | ------------ | ------------------------------------------- |
| John Doe      | 2025-08-01 | Coding       | 8             | Coding      | 7            | [john.doe@xyz.com](mailto:john.doe@xyz.com) |

ℹ️ **Column names can vary**. The tool auto-detects based on similar labels like "Name", "Task Planned", etc.

---



## ⚙️ Environment Variables (Optional)

Create a `.env` file:

```env
SMTP_SERVER=smtp.gmail.com
SMTP_PORT=587
SENDER_EMAIL=your_email@gmail.com
SENDER_PASSWORD=your_app_password
```


## 📄 API Endpoint

### `POST /send-emails`

* **Body:** `multipart/form-data` with a key `file` (Excel file)
* **Returns:** JSON summary for each employee

---

## 📌 To Do / Future Work

* 📬 Enable actual email sending from backend
* 📁 Store past summaries in a database
* 🌐 Enhance frontend with status tracking
* 📅 Add date filters and visual charts

