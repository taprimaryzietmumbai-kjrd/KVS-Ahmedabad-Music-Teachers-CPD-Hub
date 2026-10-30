/**
 * ZIET Mumbai – Participant Hub
 * OTP-protected personal profile portal. Deploy as a Google Apps Script Web App.
 * Identity is verified against Employee Code + already-registered email in the main sheet.
 */
const SPREADSHEET_ID = "1pksDgeV53waKMP1pmj0CWbpP3qyAT8nrN8RhU9Y-mKE";
const MAIN_SHEET_NAME = "PRTM ADI";
const SUBMISSIONS_SHEET = "Correction Submissions";
const OTP_TTL_SECONDS = 300;
const SESSION_TTL_SECONDS = 900;
const MAX_OTP_ATTEMPTS = 5;


/**
 * Receive public-directory correction requests from the Participants' Hub.
 * Requests are logged for coordinator verification; the master participant row is not changed automatically.
 */
function doPost(e) {
  try {
    const payloadText = e && e.parameter ? String(e.parameter.payload || "") : "";
    if (!payloadText) throw new Error("No correction details were received.");
    const p = JSON.parse(payloadText);
    const clean = v => String(v == null ? "" : v).trim().slice(0, 500);
    const {ss} = mainSheet_();
    const tabName = "Correction Submissions";
    let log = ss.getSheetByName(tabName);
    if (!log) {
      log = ss.insertSheet(tabName);
      log.appendRow(["Timestamp","Participant ID","Previous English Salutation","Previous English Name","Previous Hindi Salutation","Previous Hindi Name","Previous School","Previous School Hindi","Previous Year of Joining","Requested English Salutation","Requested English Name","Requested Hindi Salutation","Requested Hindi Name","Requested School","Requested School Hindi","Requested Year of Joining","Status"]);
      log.setFrozenRows(1);
    } else if (log.getLastRow() === 0) {
      log.appendRow(["Timestamp","Participant ID","Previous English Salutation","Previous English Name","Previous Hindi Salutation","Previous Hindi Name","Previous School","Previous School Hindi","Previous Year of Joining","Requested English Salutation","Requested English Name","Requested Hindi Salutation","Requested Hindi Name","Requested School","Requested School Hindi","Requested Year of Joining","Status"]);
      log.setFrozenRows(1);
    }
    log.appendRow([new Date(),clean(p.participantId),clean(p.previousSalutationEn),clean(p.previousName),clean(p.previousSalutationHi),clean(p.previousHi),clean(p.previousSchool),clean(p.previousSchoolHi),clean(p.previousJoiningYear),clean(p.salutationEn),clean(p.name),clean(p.salutationHi),clean(p.hi),clean(p.school),clean(p.schoolHi),clean(p.joiningYear),"Pending coordinator verification"]);
    return ContentService.createTextOutput("Correction request received. Pending coordinator verification.").setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService.createTextOutput("Correction request could not be recorded: " + String(err.message || err)).setMimeType(ContentService.MimeType.TEXT);
  }
}

function doGet(e) {
  const page = e && e.parameter ? String(e.parameter.page || "") : "";
  if (page === "hostel") {
    return HtmlService.createHtmlOutputFromFile("Hostel")
      .setTitle("Hostel Room Preference — ZIET Mumbai")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
  }
  return HtmlService.createHtmlOutputFromFile("Profile")
    .setTitle("ZIET Participant Personal Profile")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function normalized_(v) {
  return String(v == null ? "" : v).toLowerCase().replace(/[^a-z0-9\u0900-\u097f]/g, "");
}
function mainSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sh) throw new Error('Main sheet tab "' + MAIN_SHEET_NAME + '" was not found. Update MAIN_SHEET_NAME in Code.gs.');
  return {ss:ss, sh:sh};
}
function columns_(sh) {
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
  const find = aliases => {
    const wanted = aliases.map(normalized_);
    const i = headers.findIndex(h => wanted.includes(normalized_(h)));
    return i < 0 ? 0 : i + 1;
  };
  return {
    headers:headers,
    find:find,
    employee:find(["Employee Code","Employee ID","Emp Code","Employee No","कर्मचारी कोड"]),
    email:find(["Email","email","Email Address","E-mail","Official Email","ईमेल","ईमेल पता"]),
    name:find(["Participant Name","Participant Name English","Name","Teacher Name","Name in English","English Name","Name of Participant","Name of the Participant","Name of Teacher","Name of the Teacher","Teacher's Name","Name (English)","Full Name"]),
    hindiName:find(["Name in Hindi","Hindi Name","Hindi Name of Participant","Participant Name Hindi","Name of Participant in Hindi","Name (Hindi)","Hindi Full Name","प्रतिभागी का नाम","प्रतिभागी का नाम ","नाम हिंदी में","शिक्षक का नाम","हिंदी में नाम","नाम हिंदी"]),
    school:find(["Vidyalaya","School","School Name","Vidyalaya Name","Name of Vidyalaya","Name of the Vidyalaya","Name of KV","Kendriya Vidyalaya","KV Name","School/Vidyalaya","Name of School","Present KV"]),
    schoolHi:find(["Vidyalaya Name Hindi","Hindi Vidyalaya","School Name Hindi","Vidyalaya Hindi","विद्यालय","विद्यालय ","विद्यालय का नाम","विद्यालय का नाम हिंदी में","विद्यालय का हिंदी नाम","विद्यालय का नाम (हिंदी)"]),
    gender:find(["Gender","Sex","लिंग"]),
    category:find(["Social Category","Category","Reservation Category","सामाजिक वर्ग"]),
    disability:find(["Differently Abled","Disability","Divyang","दिव्यांग","दिव्यांगता"]),
    phone:find(["Phone Number","Mobile Number","Phone","Mobile","Contact Number","Whatsapp","WhatsApp","WhatsApp Number","मोबाइल नंबर","फोन नंबर"]),
    salutationEn:find(["English Salutation","Salutation"]),
    salutationHi:find(["Hindi Salutation","हिंदी संबोधन"])
  };
}
function findEmployeeRow_(sh, cols, employeeCode, email) {
  if (!cols.employee || !cols.email) throw new Error("The main sheet must have Employee Code and registered Email columns before OTP access can be used. No personal information was returned.");
  const values = sh.getDataRange().getDisplayValues();
  const ec = String(employeeCode).trim().toLowerCase();
  const em = String(email).trim().toLowerCase();
  for (let r=1;r<values.length;r++) {
    if (String(values[r][cols.employee-1]).trim().toLowerCase()===ec &&
        String(values[r][cols.email-1]).trim().toLowerCase()===em) return r+1;
  }
  return 0;
}
/**
 * Public request form for Employee Code / registered email corrections.
 * It never edits the master participant row; the coordinator verifies and updates it manually.
 */
function submitIdentityCorrectionRequest(data) {
  data = data || {};
  const clean = v => String(v == null ? "" : v).trim().slice(0, 500);
  const name = clean(data.name);
  const school = clean(data.school);
  const currentEmployeeCode = clean(data.currentEmployeeCode);
  const currentEmail = clean(data.currentEmail).toLowerCase();
  const newEmployeeCode = clean(data.newEmployeeCode);
  const newEmail = clean(data.newEmail).toLowerCase();
  const reason = clean(data.reason);
  // Quietly reject automated submissions using the hidden honeypot field.
  if (clean(data.website)) return {ok:true,message:"Your request has been received."};
  if (!name || !school || (!newEmployeeCode && !newEmail)) throw new Error("Please enter your name, Vidyalaya, and the corrected Employee Code or new email.");
  const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (currentEmail && !emailPattern.test(currentEmail)) throw new Error("Please enter a valid current email address.");
  if (newEmail && !emailPattern.test(newEmail)) throw new Error("Please enter a valid new email address.");
  const {ss} = mainSheet_();
  const tabName = "Identity Correction Requests";
  let log = ss.getSheetByName(tabName);
  if (!log) {
    log = ss.insertSheet(tabName);
    log.appendRow(["Timestamp","Full Name","Vidyalaya","Current Employee Code","Current Registered Email","Requested Employee Code","Requested New Email","Reason","Status"]);
    log.setFrozenRows(1);
  }
  log.appendRow([new Date(),name,school,currentEmployeeCode,currentEmail,newEmployeeCode,newEmail,reason,"Pending coordinator verification"]);
  MailApp.sendEmail({
    to:"taprimaryzietmumbai@gmail.com",
    subject:"Employee Code / Email correction request — ZIET Participant Hub",
    body:"A participant has submitted an identity correction request. No master-sheet record has been changed.\n\nName: "+name+
      "\nVidyalaya: "+school+
      "\nCurrent Employee Code: "+(currentEmployeeCode||"Not provided")+
      "\nCurrent registered email: "+(currentEmail||"Not provided")+
      "\nRequested Employee Code: "+(newEmployeeCode||"No change requested")+
      "\nRequested new email: "+(newEmail||"No change requested")+
      "\nReason: "+(reason||"Not provided")+
      "\n\nPlease verify the participant's identity before updating the master sheet. The request is logged in the '"+tabName+"' tab."
  });
  return {ok:true,message:"आपका अनुरोध भेज दिया गया है। समन्वयक सत्यापन के बाद रिकॉर्ड अपडेट करेंगे। / Request submitted. The coordinator will verify the details before updating the record."};
}

function sendProfileOtp(employeeCode, registeredEmail) {
  employeeCode = String(employeeCode || "").trim();
  registeredEmail = String(registeredEmail || "").trim().toLowerCase();
  if (!employeeCode || !registeredEmail || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(registeredEmail)) throw new Error("Enter your Employee Code and registered email address.");
  const cache = CacheService.getScriptCache();
  const {sh} = mainSheet_();
  const cols = columns_(sh);
  const row = findEmployeeRow_(sh, cols, employeeCode, registeredEmail);
  if (!row) throw new Error("The Employee Code and email did not match a registered record. Contact the workshop coordinator to verify your details.");
  const otp = String(Math.floor(100000 + Math.random()*900000));
  const key = "otp-" + Utilities.base64EncodeWebSafe(employeeCode.toLowerCase()+"|"+registeredEmail).slice(0,70);
  cache.put(key, JSON.stringify({hash:Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, otp).map(b=>(b+256).toString(16).slice(-2)).join(""), row:row, attempts:0}), OTP_TTL_SECONDS);
  MailApp.sendEmail({
    to:registeredEmail,
    subject:"Your ZIET Participant Hub verification code",
    body:"Your one-time verification code is: "+otp+"\n\nIt expires in 5 minutes and can be used only once. Do not share it with anyone. If you did not request this code, you can ignore this email."
  });
  return {ok:true,message:"OTP sent to your registered email. Check your inbox and spam folder. The code expires in 5 minutes."};
}
function verifyProfileOtp(employeeCode, registeredEmail, otp) {
  employeeCode=String(employeeCode||"").trim();
  registeredEmail=String(registeredEmail||"").trim().toLowerCase();
  otp=String(otp||"").trim();
  const cache=CacheService.getScriptCache();
  const key="otp-"+Utilities.base64EncodeWebSafe(employeeCode.toLowerCase()+"|"+registeredEmail).slice(0,70);
  const raw=cache.get(key);
  if(!raw) throw new Error("OTP expired or not requested. Request a new code.");
  const state=JSON.parse(raw);
  state.attempts++;
  const hash=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,otp).map(b=>(b+256).toString(16).slice(-2)).join("");
  if(hash!==state.hash) {
    if(state.attempts>=MAX_OTP_ATTEMPTS) cache.remove(key); else cache.put(key,JSON.stringify(state),OTP_TTL_SECONDS);
    throw new Error(state.attempts>=MAX_OTP_ATTEMPTS?"Too many incorrect attempts. Request a new OTP.":"Incorrect OTP. Please try again.");
  }
  cache.remove(key);
  const token=Utilities.getUuid()+Utilities.getUuid();
  cache.put("profile-session-"+token,JSON.stringify({row:state.row,employeeCode:employeeCode,email:registeredEmail}),SESSION_TTL_SECONDS);
  return {ok:true,token:token,message:"Verified. Your private profile is available for 15 minutes."};
}
function getPrivateProfile(token) {
  const session=CacheService.getScriptCache().get("profile-session-"+String(token||""));
  if(!session) throw new Error("Your session has expired. Please verify by OTP again.");
  const s=JSON.parse(session);
  const {sh}=mainSheet_(); const c=columns_(sh);
  const row=sh.getRange(s.row,1,1,sh.getLastColumn()).getDisplayValues()[0];
  const val=col=>col?row[col-1]:"";
  return {
    employeeCode:val(c.employee), email:val(c.email), name:val(c.name), hindiName:val(c.hindiName),
    school:val(c.school), schoolHi:val(c.schoolHi), gender:val(c.gender), socialCategory:val(c.category),
    differentlyAbled:val(c.disability), phone:val(c.phone), salutationEn:val(c.salutationEn), salutationHi:val(c.salutationHi)
  };
}
function savePrivateProfile(token, data) {
  const session=CacheService.getScriptCache().get("profile-session-"+String(token||""));
  if(!session) throw new Error("Your session has expired. Please verify by OTP again.");
  const s=JSON.parse(session);
  data=data||{};
  const {ss,sh}=mainSheet_(); const c=columns_(sh);
  const updates=[
    ["name",c.name],["hindiName",c.hindiName],["school",c.school],["schoolHi",c.schoolHi],
    ["gender",c.gender],["socialCategory",c.category],["differentlyAbled",c.disability],["phone",c.phone]
  ];
  if(!c.name||!c.hindiName||!c.school) throw new Error("The main sheet needs English Name, Hindi Name and Vidyalaya columns. Contact the coordinator.");
  // Email and Employee Code are identity keys and are not self-editable here.
  updates.forEach(([key,col])=>{if(col && data[key]!==undefined) sh.getRange(s.row,col).setValue(String(data[key]).trim());});
  if(c.schoolHi && data.schoolHi!==undefined) sh.getRange(s.row,c.schoolHi).setValue(String(data.schoolHi).trim());
  if(c.gender && data.gender!==undefined) sh.getRange(s.row,c.gender).setValue(String(data.gender).trim());
  if(c.category && data.socialCategory!==undefined) sh.getRange(s.row,c.category).setValue(String(data.socialCategory).trim());
  if(c.disability && data.differentlyAbled!==undefined) sh.getRange(s.row,c.disability).setValue(String(data.differentlyAbled).trim());
  if(c.phone && data.phone!==undefined) sh.getRange(s.row,c.phone).setValue(String(data.phone).trim());
  // If the master sheet lacks optional private-data columns, add them with headers.
  const addIfMissing=(key, header, value)=>{
    if(!value) return;
    let col=key==="gender"?c.gender:key==="socialCategory"?c.category:key==="differentlyAbled"?c.disability:key==="phone"?c.phone:key==="schoolHi"?c.schoolHi:0;
    if(!col){col=sh.getLastColumn()+1;sh.getRange(1,col).setValue(header);}
    sh.getRange(s.row,col).setValue(String(value).trim());
  };
  addIfMissing("schoolHi","Vidyalaya Name (Hindi)",data.schoolHi);
  addIfMissing("gender","Gender",data.gender);
  addIfMissing("socialCategory","Social Category",data.socialCategory);
  addIfMissing("differentlyAbled","Differently Abled",data.differentlyAbled);
  addIfMissing("phone","Phone Number",data.phone);
  const log=ss.getSheetByName(SUBMISSIONS_SHEET)||ss.insertSheet(SUBMISSIONS_SHEET);
  if(log.getLastRow()===0) log.appendRow(["Timestamp","Employee Code","Action","Updated Fields"]);
  log.appendRow([new Date(),s.employeeCode,"OTP verified profile update",Object.keys(data).join(", ")]);
  return {ok:true,message:"Your profile corrections have been saved to the main sheet."};
}

/**
 * Share verified contact details with the course coordinator.
 * This does not modify the participant's master record.
 * The participant must explicitly click the sharing button after OTP verification.
 */
function shareContactWithCoordinator(token) {
  const cached = CacheService.getScriptCache().get("profile-session-" + String(token || ""));
  if (!cached) throw new Error("Your session has expired. Please verify by OTP again.");
  const session = JSON.parse(cached);
  const {ss, sh} = mainSheet_();
  const c = columns_(sh);
  const row = sh.getRange(session.row, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  const val = col => col ? row[col - 1] : "";
  const employeeCode = val(c.employee);
  const email = val(c.email);
  const phone = val(c.phone);
  const name = val(c.name);
  const school = val(c.school);
  if (!employeeCode || !email) throw new Error("Employee Code or registered email is missing from the main sheet. Please contact the coordinator.");
  const logName = "Coordinator Contact Sharing";
  let log = ss.getSheetByName(logName);
  if (!log) {
    log = ss.insertSheet(logName);
    log.appendRow(["Shared At", "Participant Name", "Employee Code", "Registered Email", "Phone Number", "Vidyalaya", "Consent"]);
    log.setFrozenRows(1);
  }
  log.appendRow([new Date(), name, employeeCode, email, phone, school, "Participant explicitly consented after OTP verification"]);
  MailApp.sendEmail({
    to: "taprimaryzietmumbai@gmail.com",
    subject: "Participant contact details shared — ZIET Music Teachers CPD",
    body: "A participant has explicitly chosen to share their contact details with the course coordinator after OTP verification.\n\nParticipant: " + name +
      "\nEmployee Code: " + employeeCode +
      "\nRegistered Email: " + email +
      "\nPhone Number: " + (phone || "Not provided") +
      "\nVidyalaya: " + (school || "Not provided") +
      "\n\nThe participant's main record was not changed by this sharing action. The sharing event is recorded in the '" + logName + "' tab."
  });
  return {ok:true, message:"Your contact details have been shared with the course coordinator and recorded in the Google Sheet. Your master record was not changed."};
}


/** Save a hostel preference and upload ground-floor evidence to the designated Drive folder. */
function submitHostelPreference(data) {
  data = data || {};
  const clean = v => String(v == null ? "" : v).trim().slice(0, 500);
  const name=clean(data.name), employeeCode=clean(data.employeeCode), email=clean(data.email).toLowerCase();
  const gender=clean(data.gender), block=clean(data.block), floor=clean(data.floor), toilet=clean(data.toilet);
  const dob=clean(data.dob), disability=clean(data.disability), groundReason=clean(data.groundReason);
  if(!name || !employeeCode || !email || !gender || !block || !floor || !toilet || !dob) throw new Error("Please complete all required fields, including date of birth.");
  if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error("Enter a valid email address.");
  if(block==="A1" && gender!=="Female") throw new Error("Block A1 is reserved for ladies. Please choose A2 or A3.");
  if(!["A1","A2","A3"].includes(block) || !["Ground Floor","1st Floor","2nd Floor","3rd Floor"].includes(floor) || !["Indian","Western"].includes(toilet)) throw new Error("Please choose valid room preferences.");
  const birth=new Date(dob+"T00:00:00");
  if(isNaN(birth.getTime()) || birth > new Date()) throw new Error("Please enter a valid date of birth.");
  const age=Math.floor((Date.now()-birth.getTime())/31557600000);
  
  let fileUrl="";
  if(floor==="Ground Floor") {
    if(!data.document || !data.document.base64 || !data.document.name) throw new Error("Please upload genuine supporting documentation for a ground-floor allotment request.");
    const bytes=Utilities.base64Decode(String(data.document.base64));
    if(bytes.length > 8*1024*1024) throw new Error("The document must be 8 MB or smaller.");
    const mime=clean(data.document.mimeType) || "application/octet-stream";
    const safeName=clean(data.document.name).replace(/[\\/:*?"<>|]/g,"_").slice(0,120);
    const blob=Utilities.newBlob(bytes,mime,safeName);
    const folder=DriveApp.getFolderById("1lYibl8qAWdiRG-fD5Fe2D8WX_Szw14vd");
    const file=folder.createFile(blob);
    file.setDescription("Ground-floor hostel allotment supporting document. Submitted by "+name+" ("+employeeCode+").");
    fileUrl=file.getUrl();
  }
  const ss=SpreadsheetApp.openById("1pksDgeV53waKMP1pmj0CWbpP3qyAT8nrN8RhU9Y-mKE");
  const tab="Hostel Room Preferences";
  let sh=ss.getSheetByName(tab);
  if(!sh){sh=ss.insertSheet(tab);sh.appendRow(["Timestamp","Full Name","Employee Code","Email","Gender for A1 eligibility","Date of Birth","Age at submission","Differently Abled","Preferred Block","Preferred Floor","Preferred Toilet","Ground Floor Reason","Supporting Document URL","Review Status"]);sh.setFrozenRows(1);}
  sh.appendRow([new Date(),name,employeeCode,email,gender,birth,age,disability,block,floor,toilet,groundReason,fileUrl,"Pending coordinator review"]);
  MailApp.sendEmail({to:"taprimaryzietmumbai@gmail.com",subject:"Hostel Room Preference Submitted — KV Music Teachers CPD",body:"A hostel preference was submitted. Please verify eligibility and allotment before confirming.\n\nName: "+name+"\nEmployee Code: "+employeeCode+"\nEmail: "+email+"\nDOB: "+dob+" (age "+age+")\nDifferently Abled: "+(disability||"Not specified")+"\nBlock: "+block+"\nFloor: "+floor+"\nToilet: "+toilet+"\nGround-floor reason: "+(groundReason||"Not provided")+"\nSupporting document: "+(fileUrl||"Not uploaded")});
  return {ok:true,message:"Hostel preference submitted successfully. Allotment is subject to coordinator verification. "+(fileUrl ? "Supporting document uploaded to the designated Drive folder." : "")};
}


/** Populate the hostel form from the participant master sheet. */
function getHostelParticipantOptions() {
  const {sh}=mainSheet_();
  const c=columns_(sh);
  if(!c.employee || !c.name) throw new Error("The participant sheet needs Name and Employee Code columns.");
  const values=sh.getDataRange().getDisplayValues();
  const dobCol=c.find(["Date of Birth","DOB","Birth Date","जन्म तिथि","जन्म दिनांक"]);
  const out=[];
  for(let i=1;i<values.length;i++){
    const row=values[i];
    const employeeCode=String(row[c.employee-1]||"").trim();
    const name=String(row[c.name-1]||"").trim();
    if(!employeeCode || !name) continue;
    out.push({
      employeeCode:employeeCode,
      name:name,
      email:c.email?String(row[c.email-1]||"").trim():"",
      gender:c.gender?String(row[c.gender-1]||"").trim():"",
      disability:c.disability?String(row[c.disability-1]||"").trim():"",
      dob:dobCol?String(row[dobCol-1]||"").trim():"",
      rowNumber:i+1
    });
  }
  return out;
}
