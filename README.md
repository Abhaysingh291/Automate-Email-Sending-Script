# Email Sender Script with Excel Integration

This Node.js script sends emails to recipients from an Excel file, with attachments and HTML-formatted content. It processes emails in batches to avoid SMTP rate limits.

## Prerequisites
- **Node.js**: Install from [here](https://nodejs.org/).
- **Gmail or SMTP account** for sending emails.
- **Excel file (`EmailTset.xlsx`)** containing a column "Email".
- **PDF files** for attachments .

## Installation and Usage

```bash
git clone <repo-url> && cd <project-directory> && npm install xlsx nodemailer dotenv && \
echo "SMTP_HOST=smtp.example.com\nSMTP_PORT=587\nSMTP_USER=your_email@example.com\nSMTP_PASS=your_email_password" > .env && \
node sendEmails.js
