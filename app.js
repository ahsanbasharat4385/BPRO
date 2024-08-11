const express = require("express"); // Official Express Server Libraray
const { google } = require("googleapis"); // Google sheet API Library for Google Sheet Integration
const { XeroClient } = require("xero-node"); // For Fetch Xero Data
const mongoose = require("mongoose"); // For Storing token in Mongoose DB
const dotenv = require("dotenv"); // for Environment Variable
const cron = require("node-cron"); // For Refreshing the Xero data on specific Interval

dotenv.config();

const app = express();
app.use(express.static("public"));
app.use(express.json());

// Mongoose DB Connection
mongoose.connect(process.env.MONGO_URI);
// DataBase Schema
const tokenSchema = new mongoose.Schema({
  access_token: String,
  refresh_token: String,
  expires_at: Date,
  id_token: String,
  scope: String,
  token_type: String,
  tenant_id: String,
});

const Token = mongoose.model("Token", tokenSchema);

// Xero Authentication credentials
const xero = new XeroClient({
  clientId: process.env.XERO_CLIENT_ID,
  clientSecret: process.env.XERO_CLIENT_SECRET,
  redirectUris: [ ( process.env.APP_URL || process.env.XERO_REDIRECT_URI )],
  scopes: process.env.XERO_SCOPE.split(" "),
  httpTimeout: 10000,
});

// Google sheet Authentication
const auth = new google.auth.GoogleAuth({
  keyFile: "credentials.json",
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});

const sheets = google.sheets({ version: "v4", auth });

// LoadToken When give organization Access
async function loadTokenSet() {
  const tokenDoc = await Token.findOne();
  if (tokenDoc) {
    xero.setTokenSet({
      access_token: tokenDoc.access_token,
      refresh_token: tokenDoc.refresh_token,
      expires_at: tokenDoc.expires_at.getTime(),
      id_token: tokenDoc.id_token,
      scope: tokenDoc.scope,
      token_type: tokenDoc.token_type,
    });
    await xero.updateTenants();
  }
}

// Save Token in DataBase
async function saveTokenSet(tokenSet) {
  await xero.updateTenants(); // Ensure tenants are updated before accessing
  const tenantId = xero.tenants[0]?.tenantId; // Use optional chaining to avoid undefined error
  if (!tenantId) {
    throw new Error("No tenant ID available.");
  }

  const expiresAt = new Date(Date.now() + tokenSet.expires_in * 1000);

  const tokenDoc = await Token.findOne();
  if (tokenDoc) {
    tokenDoc.access_token = tokenSet.access_token;
    tokenDoc.refresh_token = tokenSet.refresh_token;
    tokenDoc.expires_at = expiresAt;
    tokenDoc.id_token = tokenSet.id_token;
    tokenDoc.scope = tokenSet.scope;
    tokenDoc.token_type = tokenSet.token_type;
    tokenDoc.tenant_id = tenantId;
    await tokenDoc.save();
    console.log("Access token and refresh token updated in the database.");
  } else {
    await new Token({
      access_token: tokenSet.access_token,
      refresh_token: tokenSet.refresh_token,
      expires_at: expiresAt,
      id_token: tokenSet.id_token,
      scope: tokenSet.scope,
      token_type: tokenSet.token_type,
      tenant_id: tenantId,
    }).save();
    console.log("New token set stored in the database.");
  }
}

// Load tokens when starting the server
// loadTokenSet();

// Middleware to refresh token if needed
app.use(async (req, res, next) => {
  try {
    if (xero.readTokenSet().expires_at < Date.now()) {
      const tokenSet = await xero.refreshToken();
      await saveTokenSet(tokenSet);
    }
  } catch (error) {
    console.error("Error refreshing token:", error);
  }
  next();
});

// Xero Auth URL
app.get("/auth", async (req, res) => {
  const consentUrl = await xero.buildConsentUrl();
  res.redirect(consentUrl);
});

// Callback when give access or organization
app.get("/callback", async (req, res) => {
  try {
    const tokenSet = await xero.apiCallback(req.url);
    await saveTokenSet(tokenSet);
    res.redirect("/BalanceSheet");
  } catch (error) {
    console.error("Error during callback:", error);
    res.status(500).send("Authentication error");
  }
});


// Helper Function to Format Date
const formatDate = (date) => {
  return new Date(date).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
};

const getMonthsRange = () => {
  const today = new Date(); // Get Today Date
  const startDate = new Date(2000, 0, 1); // Starting date to fetch data from xero

  const dates = []; // Array to store the date Range
  let currentDate = new Date(startDate); // Initialize Current Date with the start date
  // Loop until currentDate is less than or equal to the today's date
  while (currentDate <= today) {
    // Calculate the start of Current Month
    const start = new Date(
      currentDate.getFullYear(),
      currentDate.getMonth(),
      2
    );
    // calculate the end of the current month(Start of the next month)
    const end = new Date(
      currentDate.getFullYear(),
      currentDate.getMonth() + 1,
      1
    );
    // Push the Start and end dates to ISO Strings into dates array
    dates.push({
      start: start.toISOString().split("T")[0],
      end: end.toISOString().split("T")[0],
    });
    // Move to the next month by incrementing the month of current date
    currentDate.setMonth(currentDate.getMonth() + 1);
  }
  // Reverse the date to show latest month on the top
  return dates.reverse();
};


const fetchAndUpdateInvoices = async () => {
  try {
    // Refresh token if expired
    if (xero.readTokenSet().expires_at < Date.now()) {
      const tokenSet = await xero.refreshToken();
      await saveTokenSet(tokenSet);
      console.log("Token has been refreshed....");
    }

    try {
      const result = await xero.accountingApi.getInvoices(
        xero.tenants[0].tenantId
      );
      const invoices = result.body.invoices;

      // Filter invoices where Type is ACCREC and status is AUTHORIZED or PAID
      const accrecInvoices = invoices.filter(
        (invoice) =>
          invoice.type === "ACCREC" &&
          (invoice.status === "AUTHORISED" || invoice.status === "PAID")
      );

      // Sort invoices by date in descending order
      accrecInvoices.sort((a, b) => new Date(b.date) - new Date(a.date));

      // Fetch these data from invoices
      const rows = accrecInvoices.map((invoice) => [
        invoice.type,
        invoice.invoiceNumber,
        // invoice.invoiceID,
        invoice.reference,
        invoice.amountDue,
        invoice.amountPaid,
        invoice.contact.name,
        formatDate(invoice.date), // Use the new formatDate function
        formatDate(invoice.dueDate), // Use the new formatDate function
        invoice.status,
        invoice.total,
        invoice.currencyCode,
      ]);

      // Update Google Sheets
      const sheetId = 817851991; // Update with your sheet ID

      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: process.env.GOOGLE_SHEET_ID,
        resource: {
          requests: [
            {
              updateCells: {
                range: {
                  sheetId,
                },
                fields: "userEnteredValue",
              },
            },
          ],
        },
      });

      const headerValues = [
        "Invoice Type",
        "Invoice Number",
        // "InvoiceID",
        "Invoice Refrence",
        "Amount Due",
        "Amount Paid",
        "ContactName",
        "Invoice Date",
        "Invoice DueDate",
        "Invoice Status",
        "Total",
        "Currency Code",
      ]; // Headers of Google Sheet

      await sheets.spreadsheets.values.update({
        spreadsheetId: process.env.GOOGLE_SHEET_ID,
        range: "KYB Invoices!A1",
        valueInputOption: "RAW",
        resource: {
          values: [headerValues, ...rows],
        },
      });

      // Styling the Header of google sheet
      const requests = [
        {
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: 0,
              endRowIndex: 1,
              startColumnIndex: 0,
              endColumnIndex: headerValues.length,
            },
            cell: {
              userEnteredFormat: {
                textFormat: {
                  bold: true,
                },
              },
            },
            fields: "userEnteredFormat.textFormat.bold",
          },
        },
      ];

      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: process.env.GOOGLE_SHEET_ID,
        resource: {
          requests,
        },
      });

      console.log(
        "All available Balance Sheet reports have been fetched and stored in Google Sheets."
      );
    } catch (error) {
      if (error.code === "ETIMEDOUT") {
        console.error("Request timed out:", error);
        res.status(504).send("Request timed out. Please try again later.");
      } else {
        console.error("Error fetching invoices:", error);
        res.status(500).send("Error fetching invoices.");
      }
    }
  } catch (error) {
    console.error("Error fetching Balance Sheet reports:", error);
  }
};

//EndPoint showing after successfully fetching BalanceSheet from Xero
app.get("/BalanceSheet", async (req, res) => {
  try {
    await fetchAndUpdateInvoices();
    const htmlContent = `
     <!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <link rel="icon" type="image/svg+xml" href="./images/bprologo.jfif" />
    <title>Bpro</title>
  </head>
  <body>
     <h3> Welcome, The all data have been fetched and stored in Google Sheets.</h3>
      <p>To prevent authentication errors in the future, please click on the 'Stop' button.</p>
      <button
      id="stopButton"
      style="
        width: 120px;
        color: black;
        font-family: Arial, Helvetica, sans-serif;
        font-weight: bold;
        font-size: 18x;
        padding: 5px;
        background-color: transparent;
        border: 2px solid black;
        cursor: pointer;
      "
      onmouseover="this.style.backgroundColor='red'; this.style.color='white';"
      onmouseout="this.style.backgroundColor='transparent'; this.style.color='black'"
    >
      Stop 
    </button>
      <script>
        document.getElementById('stopButton').addEventListener('click', function() {
          fetch('/stop', { method: 'POST' })
            .then(response => response.json())
            .then(data => {
              alert(data.message);
              window.location.href = '/';
            })
            .catch(error => console.error('Error:', error));
        });
      </script>
  </body>
</html>
`;
    res.send(htmlContent);
  } catch (error) {
    console.error("Error fetching Balance Sheet reports:", error);
    res.status(500).send("Error fetching Balance Sheet reports.");
  }
});

// Endpoint to stop the cron job and delete tokens
app.post("/stop", async (req, res) => {
  try {
    await Token.deleteMany();
    res.json({ message: "Token deleted from DataBase and process stopped." });
  } catch (error) {
    console.error("Error stopping the process:", error);
    res.status(500).json({ message: "Error stopping the process." });
  }
});

// Schedule the task to run every hour of 30 minutes
cron.schedule("30 * * * *", async () => {
  try {
    await fetchAndUpdateInvoices();
  } catch (error) {
    console.error("Error running cron job:", error);
  }
});
// Express Server Running Port
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});
