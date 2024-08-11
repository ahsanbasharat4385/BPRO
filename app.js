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

let APP_URL = process.env.APP_URL || false";
console.log({ APP_URL });

// Xero Authentication credentials
const xero = new XeroClient({
  clientId: process.env.XERO_CLIENT_ID,
  clientSecret: process.env.XERO_CLIENT_SECRET,
  redirectUris: [APP_URL || process.env.XERO_REDIRECT_URI],
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
  return new Date(date).toLocaleDateString("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric",
  });
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
      2,
    );
    // calculate the end of the current month(Start of the next month)
    const end = new Date(
      currentDate.getFullYear(),
      currentDate.getMonth() + 1,
      1,
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
        xero.tenants[0].tenantId,
      );
      const invoices = result.body.invoices;

      // Filter invoices where Type is ACCREC and status is AUTHORIZED or PAID
      const accrecInvoices = invoices.filter(
        (invoice) =>
          invoice.type === "ACCREC" &&
          (invoice.status === "AUTHORISED" || invoice.status === "PAID"),
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
        "All available Balance Sheet reports have been fetched and stored in Google Sheets.",
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

const fetchAndUpdateProfitAndLoss = async () => {
  try {
    // Refresh token if expired
    if (xero.readTokenSet().expires_at < Date.now()) {
      const tokenSet = await xero.refreshToken();
      await saveTokenSet(tokenSet);
      console.log("Token has been refreshed....");
    }

    const dates = getMonthsRange(); // Fetch all months from January 2023 to current month
    const formattedRows = [];
    const processedMonths = new Set(); // To keep track of processed months
    let zeroProfitMonthsCount = 0; // Counter for consecutive months with zero profits
    let stopProcessing = false; // Flag to stop processing further months if they find 3 consecutive months of P&L is Zero

    for (const { start, end } of dates) {
      if (stopProcessing) {
        console.log(
          "Terminating further processing. Ten consecutive months with zero profits detected.",
        );
        break;
      }

      try {
        await new Promise((resolve) => setTimeout(resolve, 4000)); // Introduce a delay of 5 seconds (20000 milliseconds)
        const report = await xero.accountingApi.getReportProfitAndLoss(
          xero.tenants[0].tenantId,
          start,
          end,
        );
        const reports = report.body.reports;

        if (!reports) {
          throw new Error("Reports not found in the profit and loss report.");
        }

        let includeMonth = false;

        reports.forEach((report) => {
          let currentMonth = "";
          report.rows.forEach((row) => {
            if (row.rowType === "Header" && row.cells.length > 1) {
              currentMonth = row.cells[1].value;
            }

            if (row.rowType === "Section" && row.title === "") {
              let hasGrossProfit = false;
              let hasNetProfit = false;

              if (row.rows && row.rows.length > 0) {
                row.rows.forEach((subRow) => {
                  if (
                    subRow.rowType === "Row" &&
                    subRow.cells &&
                    subRow.cells.length > 0
                  ) {
                    const cellValue = subRow.cells[0].value;
                    const value = subRow.cells[1].value
                      ? parseFloat(subRow.cells[1].value)
                      : NaN;

                    if (cellValue === "Gross Profit" && value !== 0.0) {
                      hasGrossProfit = true;
                    }
                    if (cellValue === "Net Profit" && value !== 0.0) {
                      hasNetProfit = true;
                    }
                  }
                });
              }

              if (hasGrossProfit || hasNetProfit) {
                includeMonth = true;
                zeroProfitMonthsCount = 0; // Reset the counter if profits are non-zero
                if (currentMonth && !processedMonths.has(currentMonth)) {
                  processedMonths.add(currentMonth);
                  report.rows.forEach((row) => {
                    if (row.rowType === "Section") {
                      if (row.title) {
                        formattedRows.push([currentMonth, row.title]);
                      }
                      row.rows.forEach((subRow) => {
                        if (
                          subRow.rowType === "Row" ||
                          subRow.rowType === "SummaryRow"
                        ) {
                          formattedRows.push([
                            currentMonth,
                            ...subRow.cells.map((cell) => cell.value || ""),
                          ]);
                        }
                      });
                    }
                  });
                  console.log(
                    `Adding month ${start} to ${end} as both profits are non-zero.`,
                  );
                }
              }
            }
          });
        });

        if (!includeMonth) {
          zeroProfitMonthsCount++;
          console.log(
            `Skipping month ${start} to ${end} as both profits are zero.`,
          );
        }

        // Add two empty rows after each month's data
        formattedRows.push([], []); // Two empty rows

        // Check if ten consecutive months have zero profits
        if (zeroProfitMonthsCount === 10) {
          stopProcessing = true;
        }
      } catch (error) {
        console.error(`Error fetching report for ${start} to ${end}:`, error);
      }
    }
    // Update Google Sheets
    const sheetId = 137930456; // Update with your sheet ID

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: process.env.GOOGLE_SHEET_ID,
      resource: {
        requests: [
          {
            updateCells: {
              range: {
                sheetId,
                // Used for clearing Sheet Column Range and Row Range
                // startRowIndex: 0, // Start from Row 1 (index 0)
                // endRowIndex: 5, // End on Row 6 (index 5)
                // startColumnIndex: 0, // Starting from Column A
                // endColumnIndex: 5, // End on Column D
              },
              fields: "userEnteredValue",
            },
          },
        ],
      },
    });

    const headerValues = ["Date", "Head Name", "Amount"]; // Headers of Google Sheet

    await sheets.spreadsheets.values.update({
      spreadsheetId: process.env.GOOGLE_SHEET_ID,
      range: "KYB P&L!A1",
      valueInputOption: "RAW",
      resource: {
        values: [headerValues, ...formattedRows],
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
    console.log(formattedRows);
    console.log(
      "All available profit and loss reports have been fetched and stored in Google Sheets.",
    );
  } catch (error) {
    console.error("Error fetching profit and loss reports:", error);
  }
};

const fetchAndUpdateBalanceSheet = async () => {
  try {
    // Refresh token if expired
    if (xero.readTokenSet().expires_at < Date.now()) {
      const tokenSet = await xero.refreshToken();
      await saveTokenSet(tokenSet);
      console.log("Token has been refreshed....");
    }

    const dates = getMonthsRange(); // Fetch all months from January 2023 to current month
    const processedMonths = new Set(); // To keep track of processed Month
    let zeroNetAssetsCount = 0; // Counter for consecutive months with zero Net Assets
    let stopProcessing = false; // Flag to stop processing further if 20 consecutive months with zero Net Assets

    const uniqueKeys = new Set(); // To store all unique cell0 values
    const sectionDataMap = new Map(); // Map to store details and their values for each month

    // Initialize the sectionDataMap with empty values for each section
    dates.forEach(({ start, end }) => {
      sectionDataMap.set(start, {}); // Initialize an empty object for each month
    });

    // Fetch and process reports
    for (const { start, end } of dates) {
      if (stopProcessing) {
        console.log(
          "Terminating further processing. Twenty consecutive months with zero Net Assets detected.",
        );
        break;
      }
      try {
        await new Promise((resolve) => setTimeout(resolve, 2000)); // Delay of 2 seconds between each month
        const report = await xero.accountingApi.getReportBalanceSheet(
          xero.tenants[0].tenantId,
          start,
          end,
        );
        const reports = report.body.reports;

        if (!reports) {
          throw new Error("Reports not found in the Balance Sheet");
        }

        const monthLabel = `${start} to ${end}`;

        let includeMonth = false;

        reports.forEach((report) => {
          let currentMonth = "";
          report.rows.forEach((row) => {
            if (row.rowType === "Header" && row.cells.length > 1) {
              currentMonth = row.cells[1].value;
            }
            if (row.rowType === "Section" && row.title === "") {
              let hasNetAssets = false;
              if (row.rows && row.rows.length > 0) {
                row.rows.forEach((subRow) => {
                  if (
                    subRow.rowType === "Row" &&
                    subRow.cells &&
                    subRow.cells.length > 0
                  ) {
                    const cellValue = subRow.cells[0].value;
                    const value = subRow.cells[1].value
                      ? parseFloat(subRow.cells[1].value)
                      : NaN;
                    if (cellValue === "Net Assets" && value !== 0.0) {
                      hasNetAssets = true;
                    }
                  }
                });
              }

              if (hasNetAssets) {
                includeMonth = true;
                zeroNetAssetsCount = 0; // Reset the counter if Net Assets is non-zero
                if (currentMonth && !processedMonths.has(currentMonth)) {
                  processedMonths.add(currentMonth);
                  report.rows.forEach((row) => {
                    if (row.rowType === "Section") {
                      if (row.title) {
                        // Add section title to unique keys
                        uniqueKeys.add(row.title);
                      }
                      row.rows.forEach((subRow) => {
                        if (
                          subRow.rowType === "Row" ||
                          subRow.rowType === "SummaryRow"
                        ) {
                          const cell0Value = subRow.cells[0]?.value || "";
                          const cell1Value = subRow.cells[1]?.value || "";

                          uniqueKeys.add(cell0Value); // Add cell0 to unique keys

                          // Store cell1 values in the sectionDataMap
                          if (!sectionDataMap.get(start)[cell0Value]) {
                            sectionDataMap.get(start)[cell0Value] = {};
                          }
                          sectionDataMap.get(start)[cell0Value][monthLabel] =
                            cell1Value;
                        }
                      });
                    }
                  });
                  console.log(
                    `Adding Month ${start} to ${end} as its Net Assets are non-zero.`,
                  );
                }
              }
            }
          });
        });
        if (!includeMonth) {
          zeroNetAssetsCount++;
          console.log(
            `Skipping month ${start} to ${end} as its Net Assets are zero.`,
          );
        }
        // Check if 10 consecutive months have zero Net Assets
        if (zeroNetAssetsCount === 10) {
          stopProcessing = true;
        }
      } catch (error) {
        console.error(`Error processing dates from ${start} to ${end}:`, error);
      }
    }

    // Prepare data for Google Sheets
    const headerValues = [
      "Heads",
      ...dates
        .filter(
          ({ start }) =>
            sectionDataMap.get(start) &&
            Object.keys(sectionDataMap.get(start)).length,
        )
        .map((date) => `${date.end}`),
    ];
    const formattedRows = [];

    uniqueKeys.forEach((key) => {
      const row = [key];
      dates
        .filter(
          ({ start }) =>
            sectionDataMap.get(start) &&
            Object.keys(sectionDataMap.get(start)).length,
        )
        .forEach(({ start, end }) => {
          const monthData = sectionDataMap.get(start);
          row.push(
            monthData && monthData[key]
              ? monthData[key][`${start} to ${end}`] || ""
              : "",
          );
        });
      formattedRows.push(row);
    });

    // Clear the previous content
    const sheetId = 206785303; // Update with your sheet ID

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: process.env.GOOGLE_SHEET_ID,
      resource: {
        requests: [
          {
            updateCells: {
              range: {
                sheetId,
                startRowIndex: 0, // Clear from Row 1 (index 0)
                endRowIndex: 1000, // End on Row 1000 (or a large enough value)
                startColumnIndex: 0, // Starting from Column A
                endColumnIndex: headerValues.length, // Dynamically set the end column index
              },
              fields: "userEnteredValue",
            },
          },
        ],
      },
    });

    // Update the sheet with new data
    await sheets.spreadsheets.values.update({
      spreadsheetId: process.env.GOOGLE_SHEET_ID,
      range: "KYB BS!A1",
      valueInputOption: "RAW",
      resource: {
        values: [headerValues, ...formattedRows],
      },
    });

    // Styling the Header of Google Sheet
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

    console.log(formattedRows);
    console.log(
      "All available Balance Sheet reports have been fetched and stored in Google Sheets.",
    );
  } catch (error) {
    console.error("Error fetching Balance Sheet reports:", error);
  }
};

//EndPoint showing after successfully fetching BalanceSheet from Xero
app.get("/BalanceSheet", async (req, res) => {
  try {
    await fetchAndUpdateInvoices();
    await fetchAndUpdateProfitAndLoss();
    await fetchAndUpdateBalanceSheet();
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
