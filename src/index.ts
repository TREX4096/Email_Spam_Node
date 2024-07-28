import nodemailer from "nodemailer";
import fs from "fs";
import dotenv from "dotenv";
import xlsx from "xlsx-js-style";

dotenv.config();

const transporter = nodemailer.createTransport({
  service: "gmail",
  auth: {
    user: process.env.EMAIL_USER,
    pass: process.env.EMAIL_PASS,
  },
});

// Read the image file
const img = fs.readFileSync("test.jpg").toString("base64");

// Read the PDF file (ensure you provide the correct path)
const pdf = fs.readFileSync("Brochure.pdf").toString("base64");

// Function to read email addresses from Excel file
function readEmailAddresses(filePath: string) {
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
  return { data, workbook, worksheet };
}

const blueStyle = {
  font: {
    name: "Arial",
    sz: 24,
    color: { rgb: "FF0000" },
    bold: true,
    italic: true,
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
const { data, workbook, worksheet } = readEmailAddresses("Sample.xlsx");

const emailAddresses: string[] = [];
const userName: string[] = [];
const companyName: string[] = [];
const companySize: string[] = [];
const sector: string[] = [];

for (let i = 1; i < data.length; i++) {
  //@ts-ignore
  emailAddresses.push(data[i][1]);
  //@ts-ignore
  userName.push(data[i][0]);
  //@ts-ignore
  companyName.push(data[i][2]);
  //@ts-ignore
  companySize.push(data[i][3]);
  //@ts-ignore
  sector.push(data[i][4]);
}

console.log(emailAddresses);
console.log(userName);
console.log(companyName);
console.log(companySize);
console.log(sector);
// Function to generate email script
function generateScript(user: string, size: string, sector: string): string {
  if (size === "Big" && sector === "Web3") {
    return `Hey ${user},<br/><br/>
      We recognize your company's prestigious position in the Web3 sector. Branding is an ongoing process, and we can help you stand out by providing personalized experiences to your target audience.
      <br/><br/>
      I'm excited to introduce the 49th edition of Rendezvous, Asia’s largest cultural festival held annually at IIT Delhi. This festival attracts over 160,000+ impressions from 1600+ institutions. Over four days, we host 300+ events and pronites featuring artists like Sonu Nigam, Guru Randhawa, Salim Suleman, and comedians like Zakir Khan and Anubhav Bassi.
      <br/><br/>
      By collaborating with Rendezvous, you can access a vast audience and associate your brand with the prestige of IIT Delhi. Our past collaborators include 8fold AI, Binance, Logitech, and many more.
      <br/><br/>
      Here’s how we can help your company:
      <ul>
        <li>Engage a large audience over four days.</li>
        <li>Stage time during events and premium artist pronites.</li>
        <li>Customized brand experiences with tailored events and dedicated zones.</li>
        <li>Online and offline promotion.</li>
        <li>Hackathons, workshops, and speaker sessions.</li>
        <li>Access to data on campus students and startups.</li>
        <li>A pool of tech students for talent recruitment.</li>
      </ul>
      Let me know if you're interested. We can connect over a call to discuss more.
      <br/><br/>
      Best regards,<br>`;
  } else if (size === "Big" && sector === "Bank") {
    return `Hey ${user},<br/><br/>
      We understand that your bank holds a significant position in the market. To maintain a competitive edge, innovative branding is crucial, and we can assist you in achieving this.
      <br/><br/>
      We're excited to introduce the 49th edition of Rendezvous, Asia’s largest cultural festival held annually at IIT Delhi. The festival consistently attracts over 160,000+ impressions from 1600+ institutions, featuring over 50 renowned artists and comedians.
      <br/><br/>
      By partnering with Rendezvous, you gain unparalleled access to a diverse audience and can enhance your brand's prestige. Our past collaborators include esteemed companies such as 8fold AI, Binance, Logitech, and many more.
      <br/><br/>
      Here are some partnership opportunities:
      <ul>
        <li>Title sponsorship for select events.</li>
        <li>Access to valuable attendee data.</li>
        <li>Promotion of student credit cards and education loans.</li>
        <li>Dedicated stalls for direct interaction with attendees.</li>
        <li>Promotion of UPI services among students.</li>
        <li>Customized events and activities for launching new services.</li>
      </ul>
      We believe this partnership will offer your brand a unique opportunity to connect with a large and diverse audience. Please feel free to contact me to arrange a meeting.
      <br/><br/>
      Warm regards,<br>`;
  } else if (size === "Small" && sector === "Web3") {
    return `Hey ${user},<br/><br/>
      I recently came across your company and was impressed by your excellent product/service. It has tremendous potential to reach peak popularity, especially among the youth demographic.
      <br/><br/>
      I'm excited to introduce the 49th edition of Rendezvous, Asia’s largest cultural festival held annually at IIT Delhi. This festival attracts over 160,000+ impressions from 1600+ institutions. Over four days, we host 300+ events and pronites featuring artists like Sonu Nigam, Guru Randhawa, Salim Suleman, and comedians like Zakir Khan and Anubhav Bassi.
      <br/><br/>
      By collaborating with Rendezvous, you can access a vast audience and associate your brand with the prestige of IIT Delhi. Our past collaborators include 8fold AI, Binance, Logitech, and many more.
      <br/><br/>
      Here’s how we can help your company:
      <ul>
        <li>Engage a large audience over four days.</li>
        <li>Stage time during events and premium artist pronites.</li>
        <li>Customized brand experiences with tailored events and dedicated zones.</li>
        <li>Online and offline promotion.</li>
        <li>Hackathons, workshops, and speaker sessions.</li>
        <li>Access to data on campus students and startups.</li>
        <li>A pool of tech students for talent recruitment.</li>
      </ul>
      Let me know if you're interested. We can connect over a call to discuss more.
      <br/><br/>
      Best regards,<br>`;
  } else {
    return `Hey ${user},<br/><br/>
      We're excited to introduce the 49th edition of Rendezvous, Asia’s largest cultural festival held annually at IIT Delhi. This festival attracts over 160,000+ impressions from 1600+ institutions. Over four days, we host 300+ events and pronites featuring artists like Sonu Nigam, Guru Randhawa, Salim Suleman, and comedians like Zakir Khan and Anubhav Bassi.
      <br/><br/>
      By collaborating with Rendezvous, you can access a vast audience and associate your brand with the prestige of IIT Delhi. Our past collaborators include 8fold AI, Binance, Logitech, and many more.
      <br/><br/>
      Here’s how we can help your company:
      <ul>
        <li>Engage a large audience over four days.</li>
        <li>Stage time during events and premium artist pronites.</li>
        <li>Customized brand experiences with tailored events and dedicated zones.</li>
        <li>Online and offline promotion.</li>
        <li>Hackathons, workshops, and speaker sessions.</li>
        <li>Access to data on campus students and startups.</li>
        <li>A pool of tech students for talent recruitment.</li>
      </ul>
      Let me know if you're interested. We can connect over a call to discuss more.
      <br/><br/>
      Best regards,<br>`;
  }
}

const noOfRecipients = emailAddresses.length;

for (let i = 0; i < noOfRecipients; i++) {
  const mailScript = generateScript(userName[i], companySize[i], sector[i]);
  const mailOptions = {
    from: process.env.EMAIL_USER,
    to: emailAddresses[i],
    subject: `Association of ${companyName[i]} x IIT Delhi`,
    html: `${mailScript}<img src="cid:unique@nodemailer.com" style="width: 20%; height: auto;"/>`,
    attachments: [
      {
        filename: "1.jpg",
        content: img,
        encoding: "base64",
        cid: "unique@nodemailer.com",
      },
      {
        filename: "Brochure.pdf",
        content: pdf,
        encoding: "base64",
      },
    ],
  };

  transporter.sendMail(mailOptions, (error, info) => {
    if (error) {
      console.log(`Error sending email to ${emailAddresses[i]}:`, error);
    } else {
      console.log(`Email sent to ${emailAddresses[i]}: ${info.response}`);
      changeRowColor(worksheet, i + 2);
    }
  });
}

// Save the modified workbook after sending all emails
const outputFilePath = "Sample.xlsx";

setTimeout(() => {
  xlsx.writeFile(workbook, outputFilePath);
  console.log("Workbook saved.");
}, 5000); // Adjust timeout based on expected email send duration
