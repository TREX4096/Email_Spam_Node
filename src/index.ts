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
var img = fs
  .readFileSync("/home/trex4096/Desktop/WEb3/Projects/EmailSender/public/1.jpg")
  .toString("base64");

// Function to read email addresses from Excel file
function readEmailAddresses(filePath: any) {
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
  console.log(data);
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
  for (let col = 0; col < 4; col++) {
    const cellRef = xlsx.utils.encode_cell({ r: rowIndex - 1, c: col });
    if (worksheet[cellRef]) {
      worksheet[cellRef].s = blueStyle;
    }
  }
}

// Email addresses from Excel file
const { data, workbook, worksheet } = readEmailAddresses(
  "/home/trex4096/Desktop/WEb3/Projects/EmailSender/public/Sample.xlsx",
);

const emailAddresses: string[] = [];
const userName: string[] = [];
const company_Name: string[] = [];
const company_Size: string[] = [];
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

  //@ts-ignore
  const company = data[i][2];
  if (typeof company === "string") {
    company_Name.push(company);
  }

  //@ts-ignore
  const size = data[i][3];
  if (typeof size === "string") {
    company_Size.push(size);
  }
}

console.log(emailAddresses);
console.log(userName);
console.log(company_Name);
console.log(company_Size);

function scripts(user: string, type: string) {
  if (type == "Big") {
    return `Hey ${user},<br/><br/>
    We know that your company has an esteemed reputation in the market. However, we understand that branding is a continuous process, and to stand out among other companies, you are always looking for unique ways to brand and provide a personalized experience to your target audience. We can help you achieve this.<br/><br/>
    We are students of IIT Delhi and are excited to introduce you to the 49th edition of Rendezvous, Asia's most significant cultural festival, organized annually at IIT Delhi. Every year, Rendezvous draws over 160,000+ impressions from 1600+ institutions. In 4 days, we conducted 300+ events & pronates, which previously have witnessed 50+ artists like Sonu Nigam, Guru Randhawa, Salim Suleman and comedians such as Zakir Khan and Anubhav Bassi.<br/><br/>
    We partner with companies like yours each year to help them connect with their target audience and create a personalized brand experience. Collaborations have been highly successful, with companies seeing significant engagement and brand recognition at our events.<br/><br/>
    Here's how we helped companies like yours in the past:<br/>
    • Engage a large audience over four days.<br/>
    • Stage Time during the Events & premium artist Pronites.<br/>
    • Customise brand experiences with Tailored Events & Dedicated Zones.<br/>
    • Gorilla online & Offline Marketing strategy.<br/><br/>
    I'm eager to hear your thoughts on this exciting opportunity. If you're interested in reaching potential customers in a unique and engaging way, I'd love to discuss this further with you over a call.<br/><br/>
    Regards,<br/>
    `;
  } else {
    return `Hey ${user},<br/><br/>
    I stumbled upon your company recently and was genuinely impressed with the offerings. There's significant untapped potential, especially in gaining popularity among younger audiences.<br/><br/>
    I have an excellent way for you guys to push your company to reach and engage with its target audience. I'm excited to introduce you to the 49th edition of Rendezvous, Asia's most significant cultural festival, organized annually at IIT Delhi. Every year, Rendezvous draws over 160,000+ impressions from 1600+ institutions. In 4 days, we conducted 300+ events & pronate, which previously witnessed 50+ artists like Sonu Nigam, Guru Randhawa, Salim Suleman and comedians such as Zakir Khan and Anubhav Bassi.<br/><br/>
    We partner with companies like yours each year to help them connect with their target audience and create a personalized brand experience. Our past collaborations have been highly successful, with companies seeing significant engagement and brand recognition at our events.<br/><br/>
    Here's how we helped companies like yours in the past:<br/>
    • Engage a large audience over four days.<br/>
    • Stage Time during the Events & premium artist Pronites.<br/>
    • Customise brand experiences with Tailored Events & Dedicated Zones.<br/>
    • Online & Offline Promotion.<br/><br/>
    I'm eager to hear your thoughts on this exciting opportunity. If you're interested in reaching potential customers in a unique and engaging way, I'd love to discuss this further with you over a call.<br/><br/>
    Regards,<br/>
    `;
  }
}

const noOfRecipients = emailAddresses.length;

for (let i = 0; i < noOfRecipients; i++) {
  const mail_Script = scripts(userName[i], company_Size[i]);
  const mailOptions = {
    from: process.env.EMAIL_USER,
    to: emailAddresses[i],
    subject: `Association of ${company_Name[i]} x IIT Delhi`,
    html:
      mail_Script +
      `
      <img src="cid:unique@nodemailer.com" style="width: 20%; height: auto;"/>
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
