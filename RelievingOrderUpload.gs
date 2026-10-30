/**
 * Relieving-order upload endpoint for the KVS Ahmedabad Music Teachers CPD Hub.
 * Deploy as a Google Apps Script Web App (Execute as: Me; access: Anyone).
 * Set DESTINATION_FOLDER_ID to the Drive folder where uploaded orders should be stored.
 */
const DESTINATION_FOLDER_ID = "PASTE_GOOGLE_DRIVE_FOLDER_ID_HERE";
const MAX_FILE_BYTES = 8 * 1024 * 1024;

function doGet() {
  return HtmlService.createHtmlOutput("Relieving-order upload service is running.");
}

function doPost(e) {
  var token = "";
  try {
    if (!e || !e.parameter || !e.parameter.payload) throw new Error("No upload payload was received.");
    var p = JSON.parse(e.parameter.payload);
    token = String(p.token || "");
    if (!token) throw new Error("Missing upload token.");
    if (DESTINATION_FOLDER_ID === "PASTE_GOOGLE_DRIVE_FOLDER_ID_HERE") throw new Error("Upload service setup is incomplete. Contact the workshop coordinator.");
    var mime = String(p.mimeType || "");
    var allowed = ["application/pdf", "image/jpeg", "image/png"];
    if (allowed.indexOf(mime) === -1) throw new Error("Only PDF, JPG, and PNG files are accepted.");
    var name = String(p.fileName || "relieving-order").replace(/[\\/:*?"<>|#%{}~&]/g, "_").slice(0, 150);
    var raw = Utilities.base64Decode(String(p.fileData || ""));
    if (!raw.length || raw.length > MAX_FILE_BYTES) throw new Error("The file is empty or exceeds the 8 MB limit.");
    var folder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
    var blob = Utilities.newBlob(raw, mime, name);
    var file = folder.createFile(blob);
    file.setDescription("Relieving order uploaded for " + String(p.training || "KVS Ahmedabad Music Teachers CPD Training") + ". Upload token: " + token);
    return resultPage_(true, token, "Upload successful.");
  } catch (err) {
    return resultPage_(false, token, err && err.message ? err.message : "Upload failed. Please contact the workshop coordinator.");
  }
}

function resultPage_(ok, token, message) {
  var data = JSON.stringify({type:"RELIEVING_ORDER_UPLOAD_RESULT", ok:ok, token:token, message:message})
    .replace(/</g, "\\u003c").replace(/>/g, "\\u003e").replace(/&/g, "\\u0026");
  var html = '<!doctype html><html><head><meta charset="utf-8"></head><body><script>' +
    'window.top.postMessage(' + data + ', "*");' +
    '</script>Upload response processed.</body></html>';
  return HtmlService.createHtmlOutput(html);
}