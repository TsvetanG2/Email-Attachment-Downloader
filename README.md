# Email Attachment Downloader

A powerful desktop application for bulk downloading email attachments from Gmail and Outlook with advanced filtering, auto-renaming, and a modern GUI.

![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey.svg)
[![Download](https://img.shields.io/badge/Download_Installer-darkgreen.svg)](https://github.com/TsvetanG2/Email-Attachment-Downloader/raw/refs/heads/main/installer_output/EmailAttachmentDownloader_Setup_1.0.0.exe)


---

## Features

- **Multi-Provider Support** - Connect to Gmail or Outlook/Hotmail accounts
- **Advanced Filtering** - Filter emails by sender, subject, and date range
- **File Type Selection** - Choose which attachment types to download (PDF, images, documents, spreadsheets, etc.)
- **Calendar Date Picker** - Easy date selection with built-in calendar widget
- **Auto-Rename Files** - Multiple renaming patterns (date prefix, sender prefix, etc.)
- **Preview Before Download** - Review and select specific emails before downloading
- **Progress Tracking** - Real-time progress bar and detailed activity log
- **Threaded Downloads** - Fast parallel downloads without freezing the UI
- **Modern Dark UI** - Clean, professional interface built with CustomTkinter

---

## Screenshots

```
+------------------------------------------+
|  Email Attachment Downloader             |
+------------------------------------------+
|  Provider:  [Gmail v]  [?]               |
|  Email:     [your.email@gmail.com    ]   |
|  Password:  [****************        ]   |
|                                          |
|  [Connect]  Status: Connected            |
+------------------------------------------+
|  Email Filters                           |
|  From:    [sender@example.com        ]   |
|  Subject: [invoice                   ]   |
|  Date:    [Start date] to [End date]     |
+------------------------------------------+
|  File Types                              |
|  [x] PDF  [x] IMAGES  [x] DOCUMENTS      |
|  [x] SPREADSHEETS  [x] PRESENTATIONS     |
+------------------------------------------+
|           [Search Emails]                |
+------------------------------------------+
|  Results: 45 emails, 127 attachments     |
|                       [Preview Results]  |
+------------------------------------------+
|  Rename Pattern: [Date + Filename    v]  |
|  [ ] Convert to lowercase                |
+------------------------------------------+
|  Save to: [C:/Downloads/Attachments]     |
|        [Download All Attachments]        |
+------------------------------------------+
|  Progress: [==============    ] 70%      |
|  Downloading: invoice_march.pdf          |
+------------------------------------------+
```

---

## Installation

### Prerequisites

- Python 3.10 or higher
- pip (Python package manager)

### Quick Start

1. **Clone or download the repository**
   ```bash
   git clone https://github.com/yourusername/email-attachment-downloader.git
   cd email-attachment-downloader
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application**
   ```bash
   python main.py
   ```

---

## Email Provider Setup

### Gmail Setup

Gmail requires an **App Password** for IMAP access (not your regular password).

1. **Enable 2-Step Verification**
   - Go to [Google Account Security](https://myaccount.google.com/security)
   - Click on "2-Step Verification" and follow the setup

2. **Generate an App Password**
   - Go to [App Passwords](https://myaccount.google.com/apppasswords)
   - Select "Mail" as the app
   - Select your device
   - Click "Generate"
   - Copy the 16-character password

3. **Enable IMAP** (if not already enabled)
   - Go to [Gmail Settings](https://mail.google.com/mail/u/0/#settings/fwdandpop)
   - Under "IMAP access", select "Enable IMAP"

4. **Use in the app**
   - Select "Gmail" as provider
   - Enter your Gmail address
   - Paste the 16-character App Password (no spaces)

### Outlook / Hotmail Setup

1. **Using Regular Password** (if 2FA is disabled)
   - Simply use your regular Outlook password

2. **Using App Password** (if 2FA is enabled)
   - Go to [Microsoft Security](https://account.microsoft.com/security)
   - Under "Additional security options", find "App passwords"
   - Create a new app password
   - Use this password in the app

3. **Enable IMAP**
   - Go to Outlook.com Settings > Mail > Sync email
   - Enable "Let devices and apps use IMAP"

---

## Usage Guide

### 1. Connect to Your Email

1. Select your email provider (Gmail or Outlook)
2. Enter your email address
3. Enter your App Password
4. Click **Connect**

### 2. Set Up Filters

- **From**: Filter by sender email (e.g., `invoices@company.com`)
- **Subject**: Filter by keywords in subject (e.g., `invoice`)
- **Date Range**: Click the date buttons to open calendar picker

### 3. Select File Types

Check/uncheck the file types you want to download:
- PDF
- Images (PNG, JPG, GIF, etc.)
- Documents (DOC, DOCX, TXT, etc.)
- Spreadsheets (XLS, XLSX, CSV)
- Presentations (PPT, PPTX)
- Archives (ZIP, RAR, 7Z)

### 4. Search Emails

Click **Search Emails** to find matching emails. The results will show:
- Number of emails found
- Total attachment count

### 5. Preview Results (Optional)

Click **Preview Results** to:
- See a list of all matching emails
- Select/deselect specific emails
- View attachment names for each email

### 6. Configure Renaming

Choose a rename pattern:
| Pattern | Example Output |
|---------|----------------|
| Keep Original | `invoice.pdf` |
| Date + Filename | `2024-01-15_invoice.pdf` |
| Sender + Date + Filename | `john_2024-01-15_invoice.pdf` |
| Sender + Filename | `john_invoice.pdf` |
| Subject + Filename | `Monthly_Report_data.xlsx` |

### 7. Download

1. Set the download location (or use default)
2. Click **Download All Attachments**
3. Watch the progress bar and log

---

## Project Structure

```
email_attachment_downloader/
├── main.py                    # Application entry point
├── requirements.txt           # Python dependencies
├── README.md                  # This file
├── downloads/                 # Default download folder
└── src/
    ├── __init__.py
    ├── email_client.py        # IMAP connection & email fetching
    ├── downloader.py          # Attachment extraction & saving
    ├── renamer.py             # File renaming logic
    ├── date_picker.py         # Calendar date picker widget
    ├── preview_window.py      # Email preview/selection window
    └── gui.py                 # Main GUI application
```

---

## API Reference

### EmailClient

```python
from src.email_client import EmailClient, EmailProvider

# Using context manager
with EmailClient(EmailProvider.GMAIL) as client:
    client.connect('user@gmail.com', 'app_password')

    emails = client.search_emails(
        sender='invoices@company.com',
        subject='invoice',
        date_from=datetime(2024, 1, 1),
        date_to=datetime(2024, 12, 31)
    )

    for email in emails:
        print(f"{email.subject} - {email.attachment_count} attachments")
```

### AttachmentDownloader

```python
from src.downloader import AttachmentDownloader
from src.renamer import create_renamer_from_template

# Create renamer with date prefix
renamer = create_renamer_from_template('date_filename')

# Create downloader
downloader = AttachmentDownloader(
    download_dir='./downloads',
    renamer=renamer,
    allowed_extensions=['.pdf', '.xlsx']
)

# Download attachments
results = downloader.download_from_emails(emails)

for result in results:
    if result.success:
        print(f"Saved: {result.saved_filename}")
    else:
        print(f"Failed: {result.error}")
```

---

## Troubleshooting

### "Authentication Failed" Error

**Gmail:**
- Make sure you're using an App Password, not your regular password
- Verify 2-Step Verification is enabled
- Check that IMAP is enabled in Gmail settings

**Outlook:**
- If 2FA is enabled, you must use an App Password
- Ensure IMAP is enabled in Outlook settings

### "Connection Failed" Error

- Check your internet connection
- Verify the email address is correct
- Try again in a few minutes (rate limiting)

### No Emails Found

- Check your filter settings (sender, subject, dates)
- Try with fewer filters first
- Verify the emails exist in your INBOX folder

### Attachments Not Downloading

- Check file type filters are enabled
- Ensure the download folder is writable
- Check the log for specific errors

---

## Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| customtkinter | >= 5.2.0 | Modern GUI framework |
| Pillow | >= 10.0.0 | Image handling |
| python-dateutil | >= 2.8.2 | Date parsing |
| tkcalendar | >= 1.6.1 | Calendar date picker |

---

## Security Notes

- Your password is never stored - it's only kept in memory during the session
- Use App Passwords instead of your main password for better security
- The app only reads emails; it never modifies or deletes them
- All connections use SSL/TLS encryption (port 993)

---

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

---

## License

This project is licensed under the MIT License - see the [LICENSE](https://github.com/TsvetanG2/Email-Attachment-Downloader/blob/main/LICENSE) file for details.

---

## Acknowledgments

- [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) for the modern GUI framework
- [tkcalendar](https://github.com/j4321/tkcalendar) for the calendar widget

---

## Support

If you encounter any issues or have questions:

1. Check the [Troubleshooting](#troubleshooting) section
2. Search existing [Issues](https://github.com/yourusername/email-attachment-downloader/issues)
3. Open a new issue with:
   - Your Python version
   - Your operating system
   - Steps to reproduce the problem
   - Error messages from the log
