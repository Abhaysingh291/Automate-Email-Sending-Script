const xlsx = require("xlsx");
const nodemailer = require("nodemailer");
require("dotenv").config();
const path = require("path");

// Load Excel File
const workbook = xlsx.readFile("EmailTset.xlsx");
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const jsonData = xlsx.utils.sheet_to_json(sheet, { header: 1 });

// Check if Excel is Empty
if (!jsonData.length) {
  console.error("❌ Error: The Excel file is empty or not properly formatted.");
  process.exit(1);
}

const headerRow = jsonData[0];

// Find Email Column
const emailColumnIndex = headerRow.findIndex((col) => col && col.toLowerCase().trim() === "email");
if (emailColumnIndex === -1) {
  console.error("❌ Error: No 'Email' column found in the Excel file.");
  process.exit(1);
}

// Extract Emails (Skip Header)
const emailList = jsonData.slice(1).map((row) => row[emailColumnIndex]).filter(Boolean);

if (emailList.length === 0) {
  console.error("❌ Error: No valid email addresses found.");
  process.exit(1);
}

// Email Subject
const subject = "NJCHB jhGYK JHGHVBIH";

// Email Body (HTML Format for Bold Text)
const body = `
  <p>Dear Team,</p>
  <p>I hope this message finds you well. My name is </p>
  <p><b>Sincerely,</b><br>
 XYZ VTX<br>
  <b>Phone:</b> +91 0987654321</p>
`;

// Configure SMTP Transport
const transporter = nodemailer.createTransport({
  host: process.env.SMTP_HOST,
  port: process.env.SMTP_PORT,
  secure: false,
  auth: {
    user: process.env.SMTP_USER,
    pass: process.env.SMTP_PASS,
  },
});

// Function to Send Email with Attachments
const sendEmail = async (to, index) => {
  try {
    const filePath1 = path.join(__dirname, "XYZ.pdf");
    const filePath2 = path.join(__dirname, "YXS.pdf");

    await transporter.sendMail({
      from: process.env.SMTP_USER,
      to,
      subject,
      html: body,  // Use HTML instead of plain text
      attachments: [
        { filename: "XYZ.pdf", path: filePath1 },
        { filename: "YZX.pdf", path: filePath2 },
      ],
    });
    
    console.log(`✅ Email #${index + 1} sent to ${to}`);
  } catch (error) {
    console.error(`❌ Failed to send email #${index + 1} to ${to}:`, error.message);
  }
};

// Delay Function
const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

// Function to Send Emails in Batches
const sendEmails = async () => {
  console.log(`📧 Starting email sending process...`);
  
  for (let i = 0; i < emailList.length; i++) {
    await sendEmail(emailList[i], i);
    console.log(`⏳ Waiting 5 seconds before sending next email...`);
    await delay(5000);  // 5-second delay to prevent SMTP blocking
  }

  console.log(`✅ All emails have been sent successfully.`);
};

// Start Sending Emails
sendEmails();
