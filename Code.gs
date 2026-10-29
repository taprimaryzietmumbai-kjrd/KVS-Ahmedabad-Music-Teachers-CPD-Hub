/**
 * ZIET Mumbai – KV Music Teachers Participant Hub
 * Paste this file into Extensions > Apps Script in the supplied spreadsheet.
 * Submissions are recorded in a separate "Correction Submissions" tab so the
 * existing participant master list is not overwritten without review.
 */
const SPREADSHEET_ID = "1pksDgeV53waKMP1pmj0CWbpP3qyAT8nrN8RhU9Y-mKE";
const SUBMISSIONS_SHEET = "Correction Submissions";

function doGet() {
  return ContentService.createTextOutput("ZIET Participant Correction endpoint is ready.");
}

function doPost(e) {
  try {
    const payload = JSON.parse((e && e.parameter && e.parameter.payload) || "{}");
    const required = ["participantId", "name", "hi", "school"];
    required.forEach(function (key) {
      if (payload[key] === undefined || String(payload[key]).trim() === "") {
        throw new Error("Missing required field: " + key);
      }
    });
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sh = ss.getSheetByName(SUBMISSIONS_SHEET);
    if (!sh) {
      sh = ss.insertSheet(SUBMISSIONS_SHEET);
      sh.appendRow([
        "Submitted At", "Participant ID",
        "Previous English Salutation", "Previous English Name",
        "Previous Hindi Salutation", "Previous Hindi Name", "Previous Vidyalaya",
        "Requested English Salutation", "Requested English Name",
        "Requested Hindi Salutation", "Requested Hindi Name", "Requested Vidyalaya",
        "Status"
      ]);
      sh.setFrozenRows(1);
    }
    sh.appendRow([
      new Date(), payload.participantId,
      payload.previousSalutationEn || "", payload.previousName || "",
      payload.previousSalutationHi || "", payload.previousHi || "", payload.previousSchool || "",
      payload.salutationEn || "", payload.name, payload.salutationHi || "",
      payload.hi, payload.school, "Pending review"
    ]);
    return ContentService.createTextOutput(JSON.stringify({ ok: true, message: "Correction submitted." }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, message: String(err.message || err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
