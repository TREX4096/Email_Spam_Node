import nodemailer from "nodemailer";
import fs from "fs";
import dotenv from "dotenv";
import xlsx from "xlsx-js-style";

dotenv.config();

var transporter = nodemailer.createTransport({
  service: "gmail",
  auth: {
    user: process.env.EMAIL_USER,
    pass: process.env.EMAIL_PASS,
  },
});

// Read the image file
var img = fs.readFileSync("/path/file").toString("base64");

// Function to read email addresses from Excel file
function readEmailAddresses(filePath: any) {
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

  return { data, workbook, worksheet };
}

const blueStyle = {
  font: {
    name: "Arial", // Font name
    sz: 24, // Font size (in points)
    color: { rgb: "FF0000" }, // Font color (Red)
    bold: true, // Bold text
    italic: true, // Italic text
  },
  fill: {
    fgColor: { rgb: "0000FF" },
  },
};

// Function to change row color
function changeRowColor(worksheet: any, rowIndex: number) {
  for (let col = 0; col < 2; col++) {
    const cellRef = xlsx.utils.encode_cell({ r: rowIndex - 1, c: col });
    console.log(worksheet[cellRef]);
    if (worksheet[cellRef]) {
      worksheet[cellRef].s = blueStyle;
    }
  }
}

// Email addresses from Excel file
const { data, workbook, worksheet } = readEmailAddresses("/path/file");

const emailAddresses: string[] = [];
const userName: string[] = [];
const length = data.length;

for (let i = 1; i < length; i++) {
  //@ts-ignore
  const email = data[i][1];
  if (typeof email === "string") {
    emailAddresses.push(email);
  }
  //@ts-ignore
  const username = data[i][0];
  if (typeof username === "string") {
    userName.push(username);
  }
}

console.log(emailAddresses);
console.log(userName);

const noOfRecipients = emailAddresses.length;

for (let i = 0; i < noOfRecipients; i++) {
  const mailOptions = {
    from: process.env.EMAIL_USER,
    to: emailAddresses[i],
    subject: `Sending Email using Node.js to ${userName[i]}`,
    html: `
      <p>That was easy!</p>
      <p>Your Designation: <strong>Your Designation Here</strong></p>
      <img src="cid:unique@nodemailer.com" style="width: 200px; height: auto;"/>
    `,
    attachments: [
      {
        filename: "1.jpg",
        content: img,
        encoding: "base64",
        cid: "unique@nodemailer.com",
      },
    ],
  };

  transporter.sendMail(mailOptions, function (error, info) {
    if (error) {
      console.log(`Error sending email to ${emailAddresses[i]}:`, error);
    } else {
      console.log(`Email sent to ${emailAddresses[i]}: ${info.response}`);
      changeRowColor(worksheet, i + 2); // Change row color for the sent email (i + 2 because of the header)
    }
  });
}

// Save the modified workbook after sending all emails
const outputFilePath =
  "/home/trex4096/Desktop/WEb3/Projects/EmailSender/public/Sample.xlsx";

setTimeout(() => {
  xlsx.writeFile(workbook, outputFilePath);
  console.log("Workbook saved.");
}, 5000); // Adjust timeout based on expected email send duration
