// Paste this into Extensions > Apps Script inside your Google Sheet.
// It receives each row from the web app and appends it to a tab.
// "response" rows -> STAI_Responses tab (one row per STAI, 3 per session)
// "summary" rows  -> Session_Summary tab (one row per session)

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const name = d.type === "summary" ? "Session_Summary" : "STAI_Responses";
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.appendRow(d.headers); // header row written once
    sh.appendRow(d.values);
    return ContentService.createTextOutput("ok");
  } finally {
    lock.releaseLock();
  }
}

// Optional: open the web-app URL in a browser to confirm it is live.
function doGet() {
  return ContentService.createTextOutput("STAI-6 collector is running");
}
