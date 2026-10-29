/**
 * ZIET Mumbai – Participant Correction Web App
 * Writes approved participant corrections directly to the main roster.
 */
const SPREADSHEET_ID = "1pksDgeV53waKMP1pmj0CWbpP3qyAT8nrN8RhU9Y-mKE";
const MAIN_SHEET_NAME = "PRTM ADI";
const SUBMISSIONS_SHEET = "Correction Submissions";

function doGet() {
  return ContentService.createTextOutput("ZIET Participant Correction endpoint is ready.");
}

function doPost(e) {
  try {
    const payload = JSON.parse((e && e.parameter && e.parameter.payload) || "{}");
    ["participantId", "name", "hi", "school"].forEach(function (key) {
      if (payload[key] === undefined || String(payload[key]).trim() === "") {
        throw new Error("Missing required field: " + key);
      }
    });

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const main = ss.getSheetByName(MAIN_SHEET_NAME) || ss.getSheets()[0];
    const lastCol = main.getLastColumn();
    const headers = main.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
    const normalize = s => String(s || "").toLowerCase().replace(/[^a-z0-9\u0900-\u097f]/g, "");
    const headerKeys = headers.map(normalize);
    const findCol = aliases => {
      const aliasesN = aliases.map(normalize);
      return headerKeys.findIndex(h => aliasesN.includes(h)) + 1;
    };

    // Participant serial/ID is matched against column A; this prevents updating a different row.
    const serial = Number(payload.participantId);
    const ids = main.getRange(2, 1, Math.max(main.getLastRow() - 1, 1), 1).getValues().flat();
    const rowIndex = ids.findIndex(v => Number(v) === serial);
    if (rowIndex < 0) throw new Error("Participant ID " + serial + " was not found in column A of the main sheet.");
    const targetRow = rowIndex + 2;

    const englishNameCol = findCol(["Participant Name", "Name", "Teacher Name", "Name in English", "Teacher's Name"]);
    const hindiNameCol = findCol(["Name in Hindi", "Hindi Name", "प्रतिभागी का नाम", "नाम हिंदी में", "शिक्षक का नाम"]);
    const schoolCol = findCol(["Vidyalaya", "School", "School Name", "Vidyalaya Name", "KV Name"]);
    let schoolHiCol = findCol(["Vidyalaya Name Hindi", "Hindi Vidyalaya", "School Name Hindi", "विद्यालय का नाम", "विद्यालय का नाम हिंदी में"]);

    if (!englishNameCol || !hindiNameCol || !schoolCol) {
      throw new Error("Could not identify English name, Hindi name, and English Vidyalaya columns from row 1 headers. No changes were made.");
    }
    if (!schoolHiCol) {
      schoolHiCol = main.getLastColumn() + 1;
      main.getRange(1, schoolHiCol).setValue("Vidyalaya Name (Hindi)");
    }

    const englishFull = [payload.salutationEn, payload.name].filter(Boolean).join(" ").trim();
    const hindiFull = [payload.salutationHi, payload.hi].filter(Boolean).join(" ").trim();
    main.getRange(targetRow, englishNameCol).setValue(englishFull);
    main.getRange(targetRow, hindiNameCol).setValue(hindiFull);
    main.getRange(targetRow, schoolCol).setValue(payload.school);
    main.getRange(targetRow, schoolHiCol).setValue(payload.schoolHi || "");

    // Keep an audit trail of what was submitted and applied.
    let log = ss.getSheetByName(SUBMISSIONS_SHEET);
    if (!log) {
      log = ss.insertSheet(SUBMISSIONS_SHEET);
      log.appendRow(["Submitted At", "Participant ID", "Previous English Name", "Previous Hindi Name", "Previous Vidyalaya", "Previous Hindi Vidyalaya", "Updated English Name", "Updated Hindi Name", "Updated Vidyalaya", "Updated Hindi Vidyalaya", "Status"]);
      log.setFrozenRows(1);
    }
    log.appendRow([
      new Date(), serial, payload.previousName || "", payload.previousHi || "",
      payload.previousSchool || "", payload.previousSchoolHi || "",
      englishFull, hindiFull, payload.school, payload.schoolHi || "", "Applied to main sheet"
    ]);

    return ContentService.createTextOutput(JSON.stringify({ok:true, message:"Main participant record updated."})).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false, message:String(err.message || err)})).setMimeType(ContentService.MimeType.JSON);
  }
}
