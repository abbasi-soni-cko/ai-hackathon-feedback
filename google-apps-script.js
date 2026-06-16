// Paste this entire script into Google Apps Script (script.google.com)
// Then deploy as a Web App (Execute as: Me, Who has access: Anyone)
// Copy the deployment URL and paste it into index.html as SHEET_URL

const SHEET_NAME = "Hackathon Feedback";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    let sheet = ss.getSheetByName(SHEET_NAME);

    // Create sheet + header row on first run
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        "Timestamp",
        "Event Planning (1-5)",
        "Impact & Use Case (1-5)",
        "Learning Impact (1-5)",
        "Overall Event Experience (1-5)",
        "Written Feedback",
        "Name (optional)",
        "Email (optional)"
      ]);
      // Bold + freeze the header row
      sheet.getRange(1, 1, 1, 9).setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      data.timestamp            || "",
      Number(data.planning)     || "",
      Number(data.impact)       || "",
      Number(data.learning)     || "",
      Number(data.overall)      || "",
      data.feedback        || "",
      data.name            || "",
      data.email           || ""
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Returns all feedback rows as JSON so the page can display them publicly
function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet || sheet.getLastRow() < 2) {
    return ContentService
      .createTextOutput(JSON.stringify({ responses: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  const responses = rows.map(r => ({
    timestamp:    r[0],
    planning:     r[1],
    impact:       r[2],
    learning:     r[3],
    overall:      r[4],
    feedback:     r[5],
    name:         r[6] || "Anonymous"
  }));

  return ContentService
    .createTextOutput(JSON.stringify({ responses }))
    .setMimeType(ContentService.MimeType.JSON);
}
