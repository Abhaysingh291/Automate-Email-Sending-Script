# Email Sender Script with Excel Integration

This Node.js script sends emails to recipients from an Excel file, with attachments and HTML-formatted content. It processes emails in batches to avoid SMTP rate limits.

## Prerequisites
- **Node.js**: Install from [here](https://nodejs.org/).
- **Gmail or SMTP account** for sending emails.
- **Excel file (`EmailTset.xlsx`)** containing a column "Email".
- **PDF files** for attachments .

## Installation and Usage

- git clone <repo-url> .
- npm install xlsx nodemailer dotenv .
- SMTP_HOST=smtp.gmail.com .
- SMTP_PORT=587 .
- SMTP_USER=your_email@example.com .
- SMTP_PASS="your_email_password" .
- node AutoMailSendScript.js

