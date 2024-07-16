# Email Sender with Excel Integration

This project enables sending personalized emails to recipients listed in an Excel file, with the ability to dynamically style the Excel sheet.

## Features

- Read email addresses and usernames from an Excel file.
- Send emails using Nodemailer with attachments.
- Change the background color of rows in the Excel sheet after sending emails.

## Requirements

- Node.js
- npm
- A Gmail account for sending emails.

## Setup

1. Clone the repository:
   ```bash
   git clone https://github.com/TREX4096/Email_Spam_Node
   cd Email_Spam_Node

2. Install dependencies:

```bash
#Copy code
npm install
#Create a .env file in the root directory with your Gmail credentials:
cp .env.example .env

```
3. Make Public Dir:\
 Place your Excel file (Sample.xlsx) containing email addresses and usernames in the public directory.\

4. Usage:\
Ensure your image file (e.g., 1.jpg) is in the public directory.

Run the script:

```bash
tsc -b
node dist/index.js
