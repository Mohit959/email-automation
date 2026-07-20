# Email Automation Tool

A Python-based automation tool that sends bulk emails to recipients listed in an Excel file. Perfect for job applications, newsletters, or mass communications. Features secure Gmail authentication, template-based emails with attachments, and a safety-first design with test mode and active filtering.

## 📋 Table of Contents

- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Configuration](#configuration)
- [Usage](#usage)
- [File Structure](#file-structure)
- [Excel File Format](#excel-file-format)
- [Security & Safety](#security--safety)
- [Troubleshooting](#troubleshooting)
- [Logs](#logs)

## ✨ Features

- **Dual Excel Mode**: Switch between development (test) and production recipient lists
- **Secure Gmail Integration**: Uses Gmail App Passwords with TLS encryption
- **Template-Based Emails**: Customize subject lines, body text, and attachments
- **Active Flag Filtering**: Only send to recipients marked as active (Active = 1)
- **Test Mode**: Preview emails before sending (safe dry-run)
- **File Attachments**: Automatically attach documents (e.g., resume, CV, proposal)
- **Auto-Detection**: Automatically finds email/name columns if naming differs
- **Comprehensive Logging**: Detailed logs for debugging and auditing
- **Error Handling**: Robust error handling with informative messages
- **Production-Ready**: Confirmation prompts and validation checks

## 📦 Prerequisites

- **Python 3.7+**
- **Gmail Account** with 2-Step Verification enabled
- **openpyxl** or **xlrd** (for reading Excel files)

### Python Dependencies

```bash
pandas>=1.0.0
openpyxl>=3.0.0
```

## 🚀 Installation

### 1. Clone the Repository

```bash
git clone https://github.com/Mohit959/email-automation.git
cd email-automation
```

### 2. Install Dependencies

```bash
pip install -r requirements.txt
```

Or manually install:

```bash
pip install pandas openpyxl
```

### 3. Set Up Gmail App Password

Gmail requires an **App Password** for automated email sending:

1. Go to [Google Account Security Settings](https://myaccount.google.com/security)
2. Enable **2-Step Verification** (if not already enabled)
3. Return to Security settings and select **App passwords**
4. Choose "Mail" and "Windows Computer" (or your device)
5. Google generates a 16-character password
6. **Save this password** - you'll enter it when running the script

**Note**: Never use your actual Gmail password. Use only the App Password provided by Google.

## ⚙️ Configuration

### Configuration File: `email_config.json`

Edit the `email_config.json` file to customize email settings:

```json
{
    "smtp_server": "smtp.gmail.com",
    "smtp_port": 587,
    "excel_file": "Dev.xlsx",
    "email_column": "email",
    "name_column": "name",
    "active_column": "active",
    "subject_template": "Application for Data Engineer role",
    "body_template_file": "email_template.txt",
    "attachment_file": "Resume_Mohit.pdf"
}
```

**Configuration Options:**

| Option | Description | Example |
|--------|-------------|---------|
| `smtp_server` | Gmail SMTP server address | `smtp.gmail.com` |
| `smtp_port` | SMTP port (TLS) | `587` |
| `excel_file` | Path to recipient Excel file | `Dev.xlsx` or `Prod.xlsx` |
| `email_column` | Column name for email addresses | `email` or `Email` |
| `name_column` | Column name for recipient names | `name` or `Name` |
| `active_column` | Column to filter active recipients | `active` or `Active` |
| `subject_template` | Email subject line | `Application for Data Engineer role` |
| `body_template_file` | Path to email body template | `email_template.txt` |
| `attachment_file` | Path to file attachment (optional) | `Resume_Mohit.pdf` |

### Email Template: `email_template.txt`

Create a plain text template for the email body:

```text
Hello,

This is the email content

Best regards,
Mohit Saini
```

You can include multiple lines and formatting. The template is sent as plain text.

## 📖 Usage

### Running the Script

```bash
python gmail_automation.py
```

### Interactive Workflow

The script guides you through a series of prompts:

#### Step 1: Select Excel File
```
Gmail Email Automation
==============================

Select Excel file:
1. Dev.xlsx (Development/Test data)
2. Prod.xlsx (Production data)

Enter your choice (1 or 2): 1
```

**Dev.xlsx**: Use for testing with a small sample recipient list  
**Prod.xlsx**: Use for actual production email sends

#### Step 2: Select Mode
```
Select mode:
1. Test mode (preview emails without sending)
2. Send emails

Enter your choice (1 or 2): 1
```

**Test Mode**: 
- Preview emails before sending
- No emails are actually sent
- Safe for dry-runs and validation
- Shows which recipients would receive emails

**Send Mode**:
- Sends real emails
- Requires additional confirmation
- Requires Gmail credentials

#### Step 3: Gmail Authentication (Send Mode Only)
```
==================================================
GMAIL AUTHENTICATION
==================================================
For Gmail authentication, you need an App Password.
To create an App Password:
1. Go to your Google Account settings
2. Security > 2-Step Verification (must be enabled)
3. App passwords > Generate a new app password
4. Select 'Mail' and your device
5. Use the generated 16-character password below
==================================================

Enter your Gmail address: your.email@gmail.com
Enter your Gmail App Password: ________________
```

Enter your Gmail address and the 16-character App Password.

#### Step 4: Final Confirmation
```
Are you sure you want to send emails? (yes/no): yes
```

Type `yes` to confirm and send emails to all active recipients.

### Example Usage Scenarios

**Scenario 1: Safe Testing**
```bash
$ python gmail_automation.py
# Select: 1 (Dev.xlsx)
# Select: 1 (Test mode)
# Preview first 5 emails
# How many emails to preview? (Enter number or press Enter for all): 5
```

**Scenario 2: Production Send**
```bash
$ python gmail_automation.py
# Select: 2 (Prod.xlsx)
# Select: 2 (Send emails)
# Enter Gmail credentials
# Confirm: yes
# ✓ Emails sent to all active recipients
```

## 📁 File Structure

```
email-automation/
├── gmail_automation.py          # Main automation script
├── email_config.json            # Configuration settings
├── email_template.txt           # Email body template
├── Dev.xlsx                     # Development/test recipient list
├── Prod.xlsx                    # Production recipient list
├── Resume_Mohit.pdf             # Attachment file
├── email_automation.log         # Log file (auto-generated)
├── README.md                    # This file
└── requirements.txt             # Python dependencies
```

## 📊 Excel File Format

### Column Requirements

Your Excel files (`Dev.xlsx` and `Prod.xlsx`) should contain the following columns:

| Column | Description | Required | Example |
|--------|-------------|----------|---------|
| `email` or `Email` | Recipient email address | ✅ Yes | `john@example.com` |
| `name` or `Name` | Recipient name (for greeting) | ✅ Yes | `John Doe` |
| `active` or `Active` | Flag to filter recipients (1=send, 0=skip) | ✅ Yes | `1` or `0` |

### Example Excel Structure

| email | name | active |
|-------|------|--------|
| john@example.com | John Doe | 1 |
| jane@example.com | Jane Smith | 1 |
| bob@example.com | Bob Johnson | 0 |
| alice@example.com | Alice Brown | 1 |

In this example, emails are sent to John, Jane, and Alice (active=1), but Bob is skipped (active=0).

### Auto-Detection

If column names don't exactly match the configuration, the script attempts to auto-detect:
- Looks for columns containing "email" or "mail"
- Looks for columns containing "name"
- Looks for columns containing "active"

Example: If your column is `Email Address` instead of `email`, the script will auto-detect it.

## 🔒 Security & Safety

### Safety Features

1. **Test Mode**: Always preview emails before sending
   ```bash
   # Dry-run to check email list and content
   Select mode: 1 (Test mode)
   ```

2. **Active Flag Filtering**: Only marked recipients receive emails
   - Set `active = 1` for recipients who should receive emails
   - Set `active = 0` to skip recipients

3. **Confirmation Prompts**: Production sends require explicit confirmation
   ```
   Are you sure you want to send emails? (yes/no): yes
   ```

4. **File Validation**: Script checks for attachment files before sending
   ```
   Attachment file not found: Resume.pdf. Continue? (yes/no):
   ```

5. **Separate Excel Files**: Development and production data are kept separate
   - Dev.xlsx for testing
   - Prod.xlsx for production

### Best Practices

✅ **Always use test mode first** to verify email content and recipient list  
✅ **Use Dev.xlsx for initial testing** before moving to Prod.xlsx  
✅ **Review the log file** after sending to confirm delivery  
✅ **Keep App Passwords secure** - never commit them to version control  
✅ **Test with a small sample** before large-scale sends  
✅ **Monitor logs** for bounce or delivery errors

### Security Notes

- **Gmail App Passwords**: Only the App Password is required, not your Gmail password
- **TLS Encryption**: All connections use TLS encryption (port 587)
- **No Credentials Stored**: Passwords are entered at runtime, never stored
- **Log Sensitive Data**: Logs may contain email addresses; keep logs secure

## 🐛 Troubleshooting

### Common Issues

#### 1. "Gmail authentication failed"

**Problem**: `(b'5.7.8', b'Username and password not accepted')`

**Solutions**:
- ✅ Verify you're using an **App Password**, not your Gmail password
- ✅ Ensure **2-Step Verification** is enabled on your Google Account
- ✅ Generate a **new App Password** from Google Account settings
- ✅ Check that Gmail address is correct

#### 2. "Excel file not found"

**Problem**: `FileNotFoundError: Excel file 'Dev.xlsx' not found`

**Solutions**:
- ✅ Ensure the Excel file is in the same directory as `gmail_automation.py`
- ✅ Check spelling of filename in `email_config.json`
- ✅ Verify file extension (must be `.xlsx`)

#### 3. "Email column not found"

**Problem**: `ValueError: No email column found`

**Solutions**:
- ✅ Check Excel column name matches config (`email_column`)
- ✅ Script auto-detects columns containing "email" or "mail"
- ✅ Update `email_config.json` with correct column names
- ✅ Ensure column is not empty

#### 4. "No active records found"

**Problem**: `Warning: No active records found! All records have Active = 0`

**Solutions**:
- ✅ Set `active = 1` for recipients who should receive emails
- ✅ Check the `active` column exists in Excel
- ✅ Use test mode to preview the recipient list

#### 5. "Attachment file not found"

**Problem**: `Warning: Attachment file not found: Resume.pdf`

**Solutions**:
- ✅ Ensure attachment file is in the same directory as the script
- ✅ Check filename and extension in `email_config.json`
- ✅ Choose to continue without attachment when prompted

#### 6. "SMTP connection failed"

**Problem**: `SMTPAuthenticationError` or connection timeout

**Solutions**:
- ✅ Check internet connection
- ✅ Verify `smtp_server` is `smtp.gmail.com`
- ✅ Verify `smtp_port` is `587` (for TLS)
- ✅ Check Gmail account isn't locked/suspicious activity detected
- ✅ Wait a few minutes and retry

### Debug Mode

For more detailed debugging:

1. **Check the log file** (`email_automation.log`) for detailed error messages
2. **Run in test mode** first to identify issues before sending
3. **Verify Excel data** - ensure no missing values or invalid formats
4. **Test with a single recipient** to isolate issues

## 📋 Logs

### Log File: `email_automation.log`

The script automatically creates a log file with detailed information:

```
2024-07-20 14:32:15,123 - INFO - Email template loaded from email_template.txt
2024-07-20 14:32:16,456 - INFO - Excel file 'Dev.xlsx' loaded successfully!
2024-07-20 14:32:16,457 - INFO - Columns found: ['email', 'name', 'active']
2024-07-20 14:32:16,458 - INFO - Number of rows: 150
2024-07-20 14:32:16,459 - INFO - Found 150 recipients after filtering
2024-07-20 14:32:17,890 - INFO - Successfully connected to Gmail!
2024-07-20 14:32:18,901 - INFO - TEST MODE - Would send to: john@example.com (John Doe)
2024-07-20 14:32:18,902 - INFO - Email sending completed!
2024-07-20 14:32:18,903 - INFO - Successfully sent: 150
2024-07-20 14:32:18,904 - INFO - Failed: 0
```

### Viewing Logs

```bash
# View entire log
cat email_automation.log

# View last 20 lines
tail -20 email_automation.log

# View errors only
grep ERROR email_automation.log

# Real-time log monitoring
tail -f email_automation.log
```

### Log Information

- ✅ File loading and validation
- ✅ Column detection and mapping
- ✅ Recipient count and filtering
- ✅ Gmail connection status
- ✅ Email sending results
- ✅ Errors and warnings
- ✅ Performance metrics

## 📝 Email Content Customization

### Subject Line

Edit `email_config.json`:
```json
"subject_template": "Your Custom Subject Line"
```

### Email Body

Edit `email_template.txt`:
```text
Dear Recipient,

Your custom email body here.

Best regards,
Your Name
```

### Attachment

Update the `attachment_file` in `email_config.json`:
```json
"attachment_file": "path/to/your/file.pdf"
```

Or leave empty to send without attachment:
```json
"attachment_file": ""
```

## 🔄 Workflow Example

Here's a typical workflow for using this tool:

### Step 1: Prepare Data
```
1. Create Dev.xlsx with test recipients
2. Mark test recipients with active=1
3. Save file in project directory
```

### Step 2: Customize Email
```
1. Edit email_template.txt with your message
2. Update subject in email_config.json
3. Add resume/attachment if needed
```

### Step 3: Test
```bash
python gmail_automation.py
# Select: 1 (Dev.xlsx)
# Select: 1 (Test mode)
# Preview email content and recipient list
```

### Step 4: Production
```
1. Create/update Prod.xlsx with all recipients
2. Mark recipients to receive emails (active=1)
3. Run: python gmail_automation.py
4. Select: 2 (Prod.xlsx)
5. Select: 2 (Send emails)
6. Confirm when ready
```

### Step 5: Verify
```
1. Check email_automation.log for results
2. Verify recipient count
3. Check for any errors or warnings
```

## 📧 Example Use Cases

- **Job Applications**: Send customized applications with resume
- **Newsletters**: Send bulk newsletters to subscriber list
- **Outreach**: Mass email campaigns to contacts
- **Notifications**: Send bulk notifications to recipients
- **Surveys**: Send survey invitations

## 🤝 Contributing

Contributions are welcome! Feel free to:
- Report bugs
- Suggest improvements
- Add new features
- Improve documentation

## 📄 License

This project is open source and available under the MIT License.

## 👤 Author

**Mohit Saini**  

## 📞 Support

For issues or questions:
1. Check the [Troubleshooting](#troubleshooting) section
2. Review the log file (`email_automation.log`)
3. Check existing GitHub issues
4. Create a new issue with detailed information

---

**Last Updated**: July 2024  
**Version**: 1.0.0
