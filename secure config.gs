
function setInitialConfig() {
  PropertiesService.getScriptProperties().setProperties({
    "POSTMARK_API_TOKEN": "4ba716ef-a15e-4b71-8132-2c91e59a918d", // KEEP YOUR LIVE TOKEN HERE
    "FROM_EMAIL": "MailBounce@freightservices.net",
    
    // The specific emails for routing the different reports
    "EXTERNAL_DIGEST_RECIPIENT": "david.alexander@freightservices.net", // E.g., The Customer Success team
    "INTERNAL_DIGEST_RECIPIENT": "david.alexander@freightservices.net,michael.russell@freightservices.net", // E.g., The IT / Internal team
    
    // Kept in the config block for future use, but not actively used for routing right now
    "ALERT_RECIPIENTS": "david.alexander@freightservices.net", 
    "MESSAGE_STREAM": "bounces" 
  });
}

function getConfig() {
  const p = PropertiesService.getScriptProperties();
  return {
    API_TOKEN: p.getProperty("POSTMARK_API_TOKEN"),
    FROM_EMAIL: p.getProperty("FROM_EMAIL"),
    EXTERNAL_DIGEST: p.getProperty("EXTERNAL_DIGEST_RECIPIENT"), 
    INTERNAL_DIGEST: p.getProperty("INTERNAL_DIGEST_RECIPIENT"), 
    ALERTS: p.getProperty("ALERT_RECIPIENTS"), // Pulled in, but resting safely
    STREAM: p.getProperty("MESSAGE_STREAM") || "outbound"
  };
}
