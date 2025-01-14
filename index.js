const nodemailer = require('nodemailer');
const xlsx = require('xlsx');
const { google } = require('googleapis');

// OAuth 2.0 client setup
const oAuth2Client = new google.auth.OAuth2(
  '',  // Replace with your OAuth 2.0 Client ID
  '',  // Replace with your OAuth 2.0 Client Secret
  'https://developers.google.com/oauthplayground'  // Redirect URI for local development
);

// Set the refresh token (you'll need to get this from the OAuth flow)
oAuth2Client.setCredentials({
  refresh_token: '',  // Replace with your refresh token
});

async function sendEmail() {
  try {
    // Get the access token using the refresh token
    const accessToken = await oAuth2Client.getAccessToken();

    // Set up the transporter using OAuth2
    const transporter = nodemailer.createTransport({
      service: 'gmail',
      auth: {
        type: 'OAuth2',
        user: '',  // Replace with your Gmail address
        clientId: '',
        clientSecret: '',
        refreshToken: '',
        accessToken: accessToken.token,
      },
    });

    // Read the Excel file and extract data
    const workbook = xlsx.readFile('emails.xlsx');  // Path to your Excel file
    const sheet_name_list = workbook.SheetNames;
    const contactSheet = workbook.Sheets[sheet_name_list[0]];  // Assuming contact data is in the first sheet

    // Parse the sheet into JSON
    const contacts = xlsx.utils.sheet_to_json(contactSheet);

    // Loop through each contact and send a personalized email
    contacts.forEach(contact => {
      const companyName = contact['COMPANY NAME'];
      const contactPerson = contact['CONTACT PERSON'];
      const email = contact['MAIL ID'];

      // Prepare the subject and body
      const subject = `SDE Internship/Full-time Application for ${companyName}`;
      const text = `Hi ${contactPerson},

I’m Himanshu, a final-year B.Tech student at Delhi Technological University (Computer Science with Applied Math), applying for the SDE role (Intern/Full-time) at ${companyName}.

Skills: C++, JavaScript, MERN, Python, React, FastAPI, Express, Django
LeetCode Knight (1864+), CodeChef 3-star (1673), ranked 929 globally in LeetCode Weekly Contest 347

Quick Highlights:

Interned at Rekin Pharma (Dec 2023 – Feb 2024) building APIs and payment gateways with FastAPI and Stripe.
Interned at Sav’star Pvt. Ltd. (June 2023 – July 2023) developing React websites and SEO optimization.

Resume: https://drive.google.com/file/d/1MED9bdsHdLW31X7cGYCXqYFRKVpUMU2Y/view?usp=sharing  

Best,
Himanshu`;

      const mailOptions = {
        from: '',  // Replace with your email
        to: email,  
        subject: subject,
        text: text,
      };

      // Send the email
      transporter.sendMail(mailOptions, function(error, info) {
        if (error) {
          console.log('Error: ' + error);
        } else {
          console.log(`Email sent to ${contactPerson} at ${companyName}: ` + info.response);
        }
      });
    });
  } catch (error) {
    console.log('Error:', error);
  }
}

sendEmail();
