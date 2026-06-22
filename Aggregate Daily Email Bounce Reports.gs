/**
 * CONFIGURATION
 * Use setInitialConfig() once to set these securely in Script Properties.
 */
const SHEET_NAME = "Postmark Log";

/**
 * Runs every X minutes via trigger.
 * Fetches the latest 50 bounces from Titan and adds them to the sheet
 * if they don't already exist.
 */
function pullRecentBounces() {
  withScriptLock_(() => {
    const cfg = getConfig();
    const sheet = getLogSheet_();
    const existingIds = collectExistingMessageIds_(sheet);

    const bounces = fetchBounces_(cfg);
    if (bounces === null) return;

    // Build oldest-first so the sheet stays chronological
    const rowsToAppend = [];
    for (let i = bounces.length - 1; i >= 0; i--) {
      const bounce = bounces[i];
      if (existingIds.has(bounce.MessageID)) continue;

      rowsToAppend.push([
        bounce.BouncedAt ? new Date(bounce.BouncedAt) : new Date(),
        "Bounce",
        bounce.Email || "",
        bounce.Type || "",
        bounce.MessageID || "",
        bounce.Name || bounce.Tag || "",
        bounce.Details || bounce.Description || "",
        bounce.Subject || "Subject Unavailable",
        "", // DigestStatus (External)
        "", // DigestSentAt (External)
        "", // InternalStatus
        ""  // InternalSentAt
      ]);
    }

    if (rowsToAppend.length === 0) {
      console.log("No new bounces to add.");
      return;
    }

    const startRow = sheet.getLastRow() + 1;
    sheet
      .getRange(startRow, 1, rowsToAppend.length, rowsToAppend[0].length)
      .setValues(rowsToAppend);

    console.log(`Successfully pulled and added ${rowsToAppend.length} new bounces.`);
  });
}

/**
 * Runs Daily.
 * Sends a filtered report of EXTERNAL client bounces to the Customer Success Team.
 */
function sendDailyBounceDigest() {
  withScriptLock_(() => {
    const cfg = getConfig();
    sendDigestForColumn_({
      sheet: getLogSheet_(),
      statusColumn: 9,       // Column I
      timestampColumn: 10,   // Column J
      filterInternal: true,
      recipient: cfg.EXTERNAL_DIGEST,
      subjectFor: total => `FSI External Mail Bounce Summary: ${total} Failures Detected`,
      headerTitle: "FSI External Mail Bounce Report",
      headerSubtitle: total =>
        `<strong>${total}</strong> emails to external recipients failed to deliver in the last 2 hours.`
    });
  });
}

/**
 * Runs every 30 minutes.
 * Sends an UNFILTERED report of all recent bounces to the Internal Team.
 */
function sendInternalBounceReport() {
  withScriptLock_(() => {
    const cfg = getConfig();
    sendDigestForColumn_({
      sheet: getLogSheet_(),
      statusColumn: 11,      // Column K
      timestampColumn: 12,   // Column L
      filterInternal: false,
      recipient: cfg.INTERNAL_DIGEST,
      subjectFor: total => `Internal Bounce Alert: ${total} Failures Detected`,
      headerTitle: "Internal IT Alert: Recent Email Bounces",
      headerSubtitle: total =>
        `<strong>${total}</strong> emails failed to deliver in the last 30 minutes.`
    });
  });
}

// -- Shared digest pipeline -- //

function sendDigestForColumn_(opts) {
  const cfg = getConfig();
  const data = opts.sheet.getDataRange().getValues();

  const groupedBounces = {};
  const indicesToUpdate = [];
  let totalBounces = 0;

  for (let i = 1; i < data.length; i++) {
    const event = data[i][1];
    const email = data[i][2];
    const status = data[i][opts.statusColumn - 1];

    if (status === "SENT" || event !== "Bounce") continue;

    if (opts.filterInternal && isInternalAddress(email)) {
      // Mark as handled without including in the external digest
      indicesToUpdate.push(i + 1);
      continue;
    }

    const type = data[i][3];
    const details = data[i][6];
    const subject = data[i][7];
    const ts = data[i][0];

    const { category, recommendation } = categorizeBounce(details || "", type || "");

    if (!groupedBounces[category]) {
      groupedBounces[category] = { recommendation: recommendation, rows: [] };
    }
    groupedBounces[category].rows.push({ ts: ts, email: email, details: details, subject: subject });

    indicesToUpdate.push(i + 1);
    totalBounces++;
  }

  if (totalBounces === 0) {
    // Even with nothing to send, flag any internal-filtered rows so
    // they aren't re-scanned forever.
    if (indicesToUpdate.length > 0) {
      markRowsAsSent_(opts.sheet, indicesToUpdate, opts.statusColumn, opts.timestampColumn);
    }
    return;
  }

  const htmlBody = buildDigestHtml_(
    groupedBounces,
    totalBounces,
    opts.headerTitle,
    opts.headerSubtitle(totalBounces)
  );

  const ok = sendDigest_({
    token: cfg.INTERNAL_TOOLS_TOKEN,
    from: cfg.FROM_EMAIL,
    to: opts.recipient,
    subject: opts.subjectFor(totalBounces),
    html: htmlBody,
    stream: cfg.DIGEST_STREAM
  });

  if (!ok) {
    // Leave rows unflagged so the next run retries
    return;
  }

  markRowsAsSent_(opts.sheet, indicesToUpdate, opts.statusColumn, opts.timestampColumn);
}

// -- Postmark I/O -- //

function fetchBounces_(cfg) {
  const url = `https://api.postmarkapp.com/bounces?count=50&offset=0&messagestream=${encodeURIComponent(cfg.TITAN_STREAM)}`;
  const response = UrlFetchApp.fetch(url, {
    method: "get",
    headers: {
      "X-Postmark-Server-Token": cfg.TITAN_TOKEN,
      "Accept": "application/json"
    },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    console.error(`Failed to fetch bounces (${response.getResponseCode()}): ${response.getContentText()}`);
    return null;
  }

  const responseData = JSON.parse(response.getContentText());
  return responseData.Bounces || [];
}

function sendDigest_({ token, from, to, subject, html, stream }) {
  const response = UrlFetchApp.fetch("https://api.postmarkapp.com/email", {
    method: "post",
    headers: {
      "X-Postmark-Server-Token": token,
      "Accept": "application/json",
      "Content-Type": "application/json"
    },
    muteHttpExceptions: true,
    payload: JSON.stringify({
      "From": from,
      "To": to,
      "Subject": subject,
      "HtmlBody": html,
      "MessageStream": stream
    })
  });

  if (response.getResponseCode() === 200) return true;
  console.error(`Postmark send failed (${response.getResponseCode()}): ${response.getContentText()}`);
  return false;
}

// -- Sheet helpers -- //

function getLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
}

function collectExistingMessageIds_(sheet) {
  const data = sheet.getDataRange().getValues();
  const ids = new Set();
  for (let i = 1; i < data.length; i++) {
    if (data[i][4]) ids.add(data[i][4].toString());
  }
  return ids;
}

function markRowsAsSent_(sheet, rowIndices, statusCol, timestampCol) {
  if (!rowIndices.length) return;
  const statusLetter = columnLetter_(statusCol);
  const tsLetter = columnLetter_(timestampCol);
  const statusA1 = rowIndices.map(r => statusLetter + r);
  const tsA1 = rowIndices.map(r => tsLetter + r);
  const now = new Date();
  sheet.getRangeList(statusA1).setValue("SENT");
  sheet.getRangeList(tsA1).setValue(now);
}

function columnLetter_(col) {
  let letter = "";
  let n = col;
  while (n > 0) {
    const rem = (n - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    n = Math.floor((n - rem - 1) / 26);
  }
  return letter;
}

// -- Concurrency -- //

function withScriptLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) {
    console.warn("Could not acquire script lock; another run is in progress, skipping.");
    return;
  }
  try {
    fn();
  } finally {
    lock.releaseLock();
  }
}

// -- HTML rendering -- //

function buildDigestHtml_(groupedBounces, totalBounces, headerTitle, headerSubtitleHtml) {
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
          <h2 style="font-family: 'Bebas Neue', Impact, sans-serif; font-size: 32px; font-weight: normal; letter-spacing: 0.5px; color: #212529; margin: 0;">${escapeHtml_(headerTitle)}</h2>
          <p style="margin: 8px 0 0 0; font-size: 15px; color: #6c757d;">${headerSubtitleHtml}</p>
        </div>
  `;

  for (const [category, groupData] of Object.entries(groupedBounces)) {
    const tableRows = groupData.rows.map(r => `
      <tr>
        <td style="padding: 12px 10px; border-bottom: 1px solid #dee2e6; vertical-align: top; font-size: 14px;">${escapeHtml_(formatTimestamp_(r.ts))}</td>
        <td style="padding: 12px 10px; border-bottom: 1px solid #dee2e6; vertical-align: top; font-size: 14px; color: #0d6efd;"><strong>${escapeHtml_(r.email)}</strong></td>
        <td style="padding: 12px 10px; border-bottom: 1px solid #dee2e6; vertical-align: top; font-size: 14px; color: #495057;">${escapeHtml_(r.subject || 'N/A')}</td>
        <td style="padding: 12px 10px; border-bottom: 1px solid #dee2e6; vertical-align: top; font-size: 13px; color: #6c757d;">${escapeHtml_(r.details)}</td>
      </tr>
    `).join("");

    htmlBody += `
      <div style="margin-bottom: 45px;">
        <h3 style="font-family: 'Bebas Neue', Impact, sans-serif; font-size: 24px; font-weight: normal; letter-spacing: 0.5px; color: #dc3545; margin: 0 0 12px 0; display: flex; align-items: center;">
          <span style="margin-right: 8px;">⚠️</span> ${escapeHtml_(category)} (${groupData.rows.length})
        </h3>

        <div style="background-color: #f8d7da; color: #842029; border-left: 4px solid #dc3545; padding: 14px 18px; margin-bottom: 20px; font-size: 14px; border-radius: 0 4px 4px 0;">
          <strong style="font-weight: 500;">Recommended Action:</strong> ${escapeHtml_(groupData.recommendation)}
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

  return htmlBody;
}

function escapeHtml_(value) {
  if (value === null || value === undefined) return "";
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function formatTimestamp_(ts) {
  if (!ts) return "N/A";
  const date = new Date(ts);
  if (isNaN(date.getTime())) return "N/A";
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || "Etc/UTC";
  return Utilities.formatDate(date, tz, "MM/dd/yyyy HH:mm");
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

  const target = 'freightservices.net';
  // Skip Levenshtein when length differs by more than the threshold —
  // saves work and prevents false matches on short/unrelated domains.
  if (Math.abs(lowerDomain.length - target.length) > 2) return false;

  return getLevenshteinDistance(lowerDomain, target) <= 2;
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
