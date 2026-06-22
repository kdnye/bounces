/**
 * One-time setup: run setInitialConfig() from the Apps Script editor to
 * write these values into Script Properties, then clear the literals
 * below. After you've run it once, the live script reads everything via
 * getConfig() / PropertiesService.
 */
function setInitialConfig() {
  PropertiesService.getScriptProperties().setProperties({
    // --- Titan server (outbound mail) — used to FETCH bounces ---
    "TITAN_API_TOKEN": "f92c7446-aea6-4cbb-bc59-b48e65686451",
    "TITAN_MESSAGE_STREAM": "outbound",

    // --- Internal Tools server — used to SEND the bounce digest emails ---
    "INTERNAL_TOOLS_API_TOKEN": "4ba716ef-a15e-4b71-8132-2c91e59a918d",
    "DIGEST_MESSAGE_STREAM": "bounces",

    "FROM_EMAIL": "MailBounce@freightservices.net",

    // Routing for the two reports
    "EXTERNAL_DIGEST_RECIPIENT": "david.alexander@freightservices.net",
    "INTERNAL_DIGEST_RECIPIENT": "david.alexander@freightservices.net,michael.russell@freightservices.net",

    // Reserved for future alert routing
    "ALERT_RECIPIENTS": "david.alexander@freightservices.net"
  });
}

function getConfig() {
  const p = PropertiesService.getScriptProperties();
  return {
    TITAN_TOKEN: p.getProperty("TITAN_API_TOKEN"),
    TITAN_STREAM: p.getProperty("TITAN_MESSAGE_STREAM") || "outbound",
    INTERNAL_TOOLS_TOKEN: p.getProperty("INTERNAL_TOOLS_API_TOKEN"),
    DIGEST_STREAM: p.getProperty("DIGEST_MESSAGE_STREAM") || "bounces",
    FROM_EMAIL: p.getProperty("FROM_EMAIL"),
    EXTERNAL_DIGEST: p.getProperty("EXTERNAL_DIGEST_RECIPIENT"),
    INTERNAL_DIGEST: p.getProperty("INTERNAL_DIGEST_RECIPIENT"),
    ALERTS: p.getProperty("ALERT_RECIPIENTS")
  };
}
