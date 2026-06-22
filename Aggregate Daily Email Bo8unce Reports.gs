/**
 * CONFIGURATION
 * Use setInitialConfig() once to set these securely in Script Properties.
 */
const SHEET_NAME = "Postmark Log"; 

/**
 * Runs every X minutes via trigger. 
 * Fetches the latest 50 bounces and adds them to the sheet if they don't already exist.
 */
function pullRecentBounces() {
  const cfg = getConfig();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME) || SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  
  // 1. Get existing Message IDs to prevent duplicates
  const existingData = sheet.getDataRange().getValues();
  const existingIds = new Set();
  
  // Loop through existing sheet data (Column E / Index 4 is MessageID)
  for (let i = 1; i < existingData.length; i++) {
    if (existingData[i][4]) { 
      existingIds.add(existingData[i][4].toString());
    }
  }

  // 2. Fetch the latest 50 bounces from Postmark
  const url = `https://api.postmarkapp.com/bounces?count=50&offset=0&messagestream=${cfg.STREAM}`;
  const response = UrlFetchApp.fetch(url, {
    method: "get",
    headers: {
      "X-Postmark-Server-Token": cfg.API_TOKEN,
      "Accept": "application/json"
    },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    console.error("Failed to fetch bounces: " + response.getContentText());
    return;
  }

  const responseData = JSON.parse(response.getContentText());
  const bounces = responseData.Bounces || [];
  let addedCount = 0;

  // 3. Process new bounces (Looping backwards so oldest are added first, keeping sheet chronological)
  for (let i = bounces.length - 1; i >= 0; i--) {
    const bounce = bounces[i];
    
    // Skip if we already logged this exact bounce ID in the sheet
    if (existingIds.has(bounce.MessageID)) continue;
    
    // The Subject is automatically included in this API!
    let subjectLine = bounce.Subject || "Subject Unavailable";

    // Format the date properly for Google Sheets
    let formattedDate = bounce.BouncedAt ? new Date(bounce.BouncedAt) : new Date();

    // Append to sheet matching your 10-column layout
    sheet.appendRow([
        formattedDate,
        "Bounce",
        bounce.Email || "",
        bounce.Type || "",
        bounce.MessageID || "",
        bounce.Name || bounce.Tag || "", 
        bounce.Details || bounce.Description || "",
        subjectLine,      
        "", // DigestStatus (External)
        "", // DigestSentAt (External)
        "", // InternalStatus
        ""  // InternalSentAt
    ]);
    
    addedCount++;
  }
  
  console.log(`Successfully pulled and added ${addedCount} new bounces.`);
}

/**
 * Runs Daily. 
 * Sends a filtered report of EXTERNAL client bounces to the Customer Success Team.
 */
function sendDailyBounceDigest() {
  const cfg = getConfig();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME) || SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const data = sheet.getDataRange().getValues();
  
  const groupedBounces = {};
  const indicesToUpdate = [];
  let totalBounces = 0;

  // Iterate through rows (skipping header at index 0)
  for (let i = 1; i < data.length; i++) {
    const event = data[i][1];
    const email = data[i][2];
    const digestStatus = data[i][8]; // Column I

    if (digestStatus !== "SENT" && event === "Bounce") {
        
        // FILTER: Check if this is an internal FreightServices email or a typo
        if (isInternalAddress(email)) {
            // Mark it as handled so we don't check it again tomorrow, but SKIP adding it to the digest
            indicesToUpdate.push(i + 1);
            continue; 
        }

        const type = data[i][3];
        const details = data[i][6];
        const subject = data[i][7];
        const ts = data[i][0];

        // Run our categorization logic for valid external client emails
        const { category, recommendation } = categorizeBounce(details || "", type || "");

        // Initialize category grouping if it doesn't exist yet
        if (!groupedBounces[category]) {
            groupedBounces[category] = { recommendation: recommendation, rows: [] };
        }

        // Push data to the specific category
        groupedBounces[category].rows.push({
            ts: ts, email: email, details: details, subject: subject 
        });
        
        indicesToUpdate.push(i + 1);
        totalBounces++;
    }
  }

  if (totalBounces === 0) return; // Exit if nothing to send

  // Start building the HTML Body with FSI App Styling
  let htmlBody = `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        @import url('https://fonts.googleapis.com/css2?family=Bebas+Neue&family=Roboto:wght@300;400;500;700&display=swap');
      </style>
    </head>
    <body style="margin: 0; padding: 20px; background-color: #f4f6f8; font-family: 'Roboto', Helvetica, Arial, sans-serif; color: #212529; -webkit-font-smoothing: antialiased;">
      <div style="max-width: 800px; margin: 0 auto; background-color: #ffffff; padding: 35px; border-radius: 6px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
        
        <div style="border-bottom: 2px solid #212529; padding-bottom: 15px; margin-bottom: 30px;">
          <h2 style="font-family: 'Bebas Neue', Impact, sans-serif; font-size: 32px; font-weight: normal; letter-spacing: 0.5px; color: #212529; margin: 0;">FSI External Mail Bounce Report</h2>
          <p style="margin: 8px 0 0 0; font-size: 15px; color: #6c757d;">
            <strong>${totalBounces}</strong> emails to external recipients failed to deliver in the last 2 hours.
          </p>
        </div>
  `;

  for (const [category, groupData] of Object.entries(groupedBounces)) {
    let tableRows = groupData.rows.map(r => `
      <tr>
        <td style="padding: 12px 10px; border-bottom: 1px solid #dee2e6; vertical-align: top; font-size: 14px;">${Utilities.formatDate(new Date(r.ts), "MST", "MM/dd/yyyy HH:mm")}</td>
        <td style="padding: 12px 10px; border-bottom: 1px solid #dee2e6; vertical-align: top; font-size: 14px; color: #0d6efd;"><strong>${r.email}</strong></td>
        <td style="padding: 12px 10px; border-bottom: 1px solid #dee2e6; vertical-align: top; font-size: 14px; color: #495057;">${r.subject || 'N/A'}</td> 
        <td style="padding: 12px 10px; border-bottom: 1px solid #dee2e6; vertical-align: top; font-size: 13px; color: #6c757d;">${r.details}</td>
      </tr>
    `).join("");

    htmlBody += `
      <div style="margin-bottom: 45px;">
        <h3 style="font-family: 'Bebas Neue', Impact, sans-serif; font-size: 24px; font-weight: normal; letter-spacing: 0.5px; color: #dc3545; margin: 0 0 12px 0; display: flex; align-items: center;">
          <span style="margin-right: 8px;">⚠️</span> ${category} (${groupData.rows.length})
        </h3>
        
        <div style="background-color: #f8d7da; color: #842029; border-left: 4px solid #dc3545; padding: 14px 18px; margin-bottom: 20px; font-size: 14px; border-radius: 0 4px 4px 0;">
          <strong style="font-weight: 500;">Recommended Action:</strong> ${groupData.recommendation}
        </div>

        <table style="width: 100%; border-collapse: collapse; text-align: left; margin-bottom: 10px;">
          <thead>
            <tr>
              <th style="padding: 12px 10px; border-bottom: 2px solid #495057; font-weight: 500; font-size: 14px; color: #212529;">Date/Time</th>
              <th style="padding: 12px 10px; border-bottom: 2px solid #495057; font-weight: 500; font-size: 14px; color: #212529;">Recipient</th>
              <th style="padding: 12px 10px; border-bottom: 2px solid #495057; font-weight: 500; font-size: 14px; color: #212529;">Subject</th> 
              <th style="padding: 12px 10px; border-bottom: 2px solid #495057; font-weight: 500; font-size: 14px; color: #212529;">Server Response</th>
            </tr>
          </thead>
          <tbody>
            ${tableRows}
          </tbody>
        </table>
      </div>
    `;
  }

  htmlBody += `
        <div style="border-top: 1px solid #dee2e6; padding-top: 20px; margin-top: 30px; text-align: center;">
          <p style="font-size: 13px; color: #adb5bd; margin: 0;">Sent automatically via Postmark Integration</p>
        </div>
      </div>
    </body>
    </html>
  `;

  // Send the email to the EXTERNAL_DIGEST_RECIPIENT
  UrlFetchApp.fetch("https://api.postmarkapp.com/email", {
    method: "post",
    headers: {
      "X-Postmark-Server-Token": cfg.API_TOKEN,
      "Accept": "application/json",
      "Content-Type": "application/json"
    },
    payload: JSON.stringify({
      "From": cfg.FROM_EMAIL,
      "To": cfg.EXTERNAL_DIGEST,
      "Subject": `FSI External Mail Bounce Summary: ${totalBounces} Failures Detected`,
      "HtmlBody": htmlBody,
      "MessageStream": "outbound"
    })
  });

  // Mark rows as sent
  const now = new Date();
  indicesToUpdate.forEach(rowIdx => {
    sheet.getRange(rowIdx, 9).setValue("SENT");
    sheet.getRange(rowIdx, 10).setValue(now);
  });
}

/**
 * Runs every 30 minutes. 
 * Sends an UNFILTERED report of all recent bounces to the Internal Team.
 */
function sendInternalBounceReport() {
  const cfg = getConfig();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME) || SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const data = sheet.getDataRange().getValues();
  
  const groupedBounces = {};
  const indicesToUpdate = [];
  let totalBounces = 0;

  for (let i = 1; i < data.length; i++) {
    const event = data[i][1];
    const internalStatus = data[i][10]; // Look at Column K (Index 10)

    if (internalStatus !== "SENT" && event === "Bounce") {
        const ts = data[i][0];
        const email = data[i][2];
        const type = data[i][3];
        const details = data[i][6];
        const subject = data[i][7];

        const { category, recommendation } = categorizeBounce(details || "", type || "");

        if (!groupedBounces[category]) {
            groupedBounces[category] = { recommendation: recommendation, rows: [] };
        }

        groupedBounces[category].rows.push({
            ts: ts, email: email, details: details, subject: subject 
        });
        
        indicesToUpdate.push(i + 1);
        totalBounces++;
    }
  }

  if (totalBounces === 0) return; 

  let htmlBody = `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        @import url('https://fonts.googleapis.com/css2?family=Bebas+Neue&family=Roboto:wght@300;400;500;700&display=swap');
      </style>
    </head>
    <body style="margin: 0; padding: 20px; background-color: #f4f6f8; font-family: 'Roboto', Helvetica, Arial, sans-serif; color: #212529; -webkit-font-smoothing: antialiased;">
      <div style="max-width: 800px; margin: 0 auto; background-color: #ffffff; padding: 35px; border-radius: 6px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
        
        <div style="border-bottom: 2px solid #212529; padding-bottom: 15px; margin-bottom: 30px;">
          <h2 style="font-family: 'Bebas Neue', Impact, sans-serif; font-size: 32px; font-weight: normal; letter-spacing: 0.5px; color: #212529; margin: 0;">Internal IT Alert: Recent Email Bounces</h2>
          <p style="margin: 8px 0 0 0; font-size: 15px; color: #6c757d;">
            <strong>${totalBounces}</strong> emails failed to deliver in the last 30 minutes.
          </p>
        </div>
  `;

  for (const [category, groupData] of Object.entries(groupedBounces)) {
    let tableRows = groupData.rows.map(r => `
      <tr>
        <td style="padding: 12px 10px; border-bottom: 1px solid #dee2e6; vertical-align: top; font-size: 14px;">${Utilities.formatDate(new Date(r.ts), "MST", "MM/dd/yyyy HH:mm")}</td>
        <td style="padding: 12px 10px; border-bottom: 1px solid #dee2e6; vertical-align: top; font-size: 14px; color: #0d6efd;"><strong>${r.email}</strong></td>
        <td style="padding: 12px 10px; border-bottom: 1px solid #dee2e6; vertical-align: top; font-size: 14px; color: #495057;">${r.subject || 'N/A'}</td> 
        <td style="padding: 12px 10px; border-bottom: 1px solid #dee2e6; vertical-align: top; font-size: 13px; color: #6c757d;">${r.details}</td>
      </tr>
    `).join("");

    htmlBody += `
      <div style="margin-bottom: 45px;">
        <h3 style="font-family: 'Bebas Neue', Impact, sans-serif; font-size: 24px; font-weight: normal; letter-spacing: 0.5px; color: #dc3545; margin: 0 0 12px 0; display: flex; align-items: center;">
          <span style="margin-right: 8px;">⚠️</span> ${category} (${groupData.rows.length})
        </h3>
        
        <div style="background-color: #f8d7da; color: #842029; border-left: 4px solid #dc3545; padding: 14px 18px; margin-bottom: 20px; font-size: 14px; border-radius: 0 4px 4px 0;">
          <strong style="font-weight: 500;">Recommended Action:</strong> ${groupData.recommendation}
        </div>

        <table style="width: 100%; border-collapse: collapse; text-align: left; margin-bottom: 10px;">
          <thead>
            <tr>
              <th style="padding: 12px 10px; border-bottom: 2px solid #495057; font-weight: 500; font-size: 14px; color: #212529;">Date/Time</th>
              <th style="padding: 12px 10px; border-bottom: 2px solid #495057; font-weight: 500; font-size: 14px; color: #212529;">Recipient</th>
              <th style="padding: 12px 10px; border-bottom: 2px solid #495057; font-weight: 500; font-size: 14px; color: #212529;">Subject</th> 
              <th style="padding: 12px 10px; border-bottom: 2px solid #495057; font-weight: 500; font-size: 14px; color: #212529;">Server Response</th>
            </tr>
          </thead>
          <tbody>
            ${tableRows}
          </tbody>
        </table>
      </div>
    `;
  }

  htmlBody += `
        <div style="border-top: 1px solid #dee2e6; padding-top: 20px; margin-top: 30px; text-align: center;">
          <p style="font-size: 13px; color: #adb5bd; margin: 0;">Sent automatically via Postmark Integration</p>
        </div>
      </div>
    </body>
    </html>
  `;

  // Send the email to the INTERNAL_DIGEST_RECIPIENT
  UrlFetchApp.fetch("https://api.postmarkapp.com/email", {
    method: "post",
    headers: {
      "X-Postmark-Server-Token": cfg.API_TOKEN,
      "Accept": "application/json",
      "Content-Type": "application/json"
    },
    payload: JSON.stringify({
      "From": cfg.FROM_EMAIL,
      "To": cfg.INTERNAL_DIGEST,
      "Subject": `Internal Bounce Alert: ${totalBounces} Failures Detected`,
      "HtmlBody": htmlBody,
      "MessageStream": "outbound"
    })
  });

  const now = new Date();
  indicesToUpdate.forEach(rowIdx => {
    sheet.getRange(rowIdx, 11).setValue("SENT");
    sheet.getRange(rowIdx, 12).setValue(now);
  });
}

// -- Utility & Filtering Functions Below -- //

function categorizeBounce(detailsString, bounceType) {
  const details = detailsString.toLowerCase();
  
  if (details.includes("hop count exceeded") || details.includes("mail loop")) {
      return { category: "Mail Loop (Client IT Error)", recommendation: "The email address is likely correct, but the client's internal IT department has a misconfigured forwarding rule crashing the email. Contact them by phone to let them know." };
  } else if (details.includes("user unknown") || details.includes("no such user") || details.includes("does not exist") || details.includes("invalid recipient")) {
      return { category: "Invalid Address / Misspelling", recommendation: "Look closely at the email address for typos (e.g., 'gmial' instead of 'gmail' or a misspelled name). If it looks correct, the person may no longer work at the company." };
  } else if (details.includes("null mx") || details.includes("unresolvable") || details.includes("domain not found")) {
      return { category: "Domain Doesn't Exist", recommendation: "The part of the email address AFTER the '@' symbol is spelled wrong, or the company's website is completely offline." };
  } else if (details.includes("queue.expired") || bounceType === "Transient") {
      return { category: "Transient / Delayed / Full Inbox", recommendation: "The recipient's email server is temporarily offline or their inbox is full. Postmark tried to deliver this for 48 hours but eventually gave up. You can try sending again later." };
  } else {
      return { category: "General Block / Unknown Error", recommendation: "The email was blocked by a spam filter, a firewall, or an unspecified error. Verify the address is correct and the client expects our emails." };
  }
}

function isInternalAddress(email) {
  if (!email) return false;
  const domain = email.split('@')[1];
  if (!domain) return false;
  const lowerDomain = domain.toLowerCase();

  if (lowerDomain === 'freightservices.net') return true;

  const knownTypos = [
    'freightservces.net', 'freigthservices.net', 'freghtservicves.net',
    'freighservices.net', 'freighstervies.net', 'feightservices.net', 'freighstervices.net'
  ];
  if (knownTypos.includes(lowerDomain)) return true;

  const distance = getLevenshteinDistance(lowerDomain, 'freightservices.net');
  if (distance <= 4) return true;

  return false;
}

function getLevenshteinDistance(a, b) {
  if (!a || !b) return (a || b || "").length;
  if (a === b) return 0;
  
  const matrix = [];
  for (let i = 0; i <= b.length; i++) matrix[i] = [i];
  for (let j = 0; j <= a.length; j++) matrix[0][j] = j;

  for (let i = 1; i <= b.length; i++) {
    for (let j = 1; j <= a.length; j++) {
      if (b.charAt(i - 1) === a.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1, 
          Math.min(matrix[i][j - 1] + 1, matrix[i - 1][j] + 1) 
        );
      }
    }
  }
  return matrix[b.length][a.length];
}
