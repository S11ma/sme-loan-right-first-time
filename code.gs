function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('SME Portal')
    .addItem('Open Chatbot', 'openChatbot')
    .addToUi();
}

// Open sidebar chatbot
function openChatbot() {
  var html = HtmlService.createHtmlOutputFromFile('Chatbot')
    .setTitle('SME Loan Pre‑Screen Bot');
  SpreadsheetApp.getUi().showSidebar(html);
}

// Called by chatbot when user clicks "Submit"
function submitChatbotApplication(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // CHANGE these names only if your tabs are different
  var dataSheet = ss.getSheetByName('Loan Application DataSet');

  if (!dataSheet) {
    throw new Error('Sheet "Loan Application DataSet" not found.');
  }

  // Read data from payload object
  var applicantId = payload.applicantId;
  var industry    = payload.industry;
  var amount      = payload.amount;
  var category    = payload.category;
  var incomeDoc   = payload.incomeDoc;
  var kyc         = payload.kyc;
  var business    = payload.business;

  if (!applicantId || !industry || !amount || !category) {
    throw new Error('Missing mandatory fields.');
  }

  // Last used row and next row
  var lastRow = dataSheet.getLastRow();
  var nextRow = lastRow + 1;

  // Columns:
  // B: S.No., C: Applicant ID, D: Industry, E: Loan Amount,
  // F: Loan Category, G: Applicant Category, H: Income, I: KYC, J: Business Proof
  dataSheet.getRange(nextRow, 2).setValue(nextRow - 1);      // S.No.
  dataSheet.getRange(nextRow, 3).setValue(applicantId);
  dataSheet.getRange(nextRow, 4).setValue(industry);
  dataSheet.getRange(nextRow, 5).setValue(amount);
  dataSheet.getRange(nextRow, 6).setValue('SME');
  dataSheet.getRange(nextRow, 7).setValue(category);
  dataSheet.getRange(nextRow, 8).setValue(incomeDoc);
  dataSheet.getRange(nextRow, 9).setValue(kyc);
  dataSheet.getRange(nextRow, 10).setValue(business);

  // Copy formulas K–O from previous row (rule engine)
  if (lastRow >= 2) {
    dataSheet.getRange(lastRow, 11, 1, 5).copyTo(
      dataSheet.getRange(nextRow, 11, 1, 5),
      {contentsOnly: false}
    );
  }

  // Return final status back to chatbot
  var status = dataSheet.getRange(nextRow, 15).getValue(); // Column O
  return status;
}



// // 1) Set this to the ID of your SME_Documents folder in Google Drive
// const FOLDER_ID = '1ftUNDNAVAAJpDxi_BgvQ-KoXExD7OEDe';

// // 2) Main data sheet name
// const DATA_SHEET_NAME = 'Loan Application DataSet';


// // =========================
// // MENU & SIDEBAR
// // =========================

// // Create custom menu when sheet opens
// function onOpen() {
//   SpreadsheetApp.getUi()
//     .createMenu('SME Portal')
//     .addItem('Open Chatbot', 'openChatbot')
//     .addToUi();
// }

// // Show the sidebar (HTML file name must be ChatbotUpload.html)
// function openChatbot() {
//   var html = HtmlService.createHtmlOutputFromFile('ChatbotUpload')
//     .setTitle('SME Loan Pre‑Screen Bot');
//   SpreadsheetApp.getUi().showSidebar(html);
// }


// // =========================
// // FILE STORAGE HELPER
// // =========================

// // Save one uploaded file object into Drive under FOLDER_ID/subFolderName
// // fileObj = { filename, mimeType, bytes(base64) }  OR null
// // Returns: file URL string or "" if no file
// function saveUploadedFile(fileObj, subFolderName) {
//   if (!fileObj || !fileObj.bytes) return "";

//   // Parent folder
//   var parent = DriveApp.getFolderById(FOLDER_ID);

//   // Subfolder (KYC / INCOME / BUSINESS)
//   var folders = parent.getFoldersByName(subFolderName);
//   var folder = folders.hasNext() ? folders.next() : parent.createFolder(subFolderName);

//   // Convert base64 to Blob and create file
//   var blob = Utilities.newBlob(
//     Utilities.base64Decode(fileObj.bytes),
//     fileObj.mimeType,
//     fileObj.filename
//   );
//   var file = folder.createFile(blob);

//   return file.getUrl();
// }


// // =========================
// // MAIN SUBMIT FUNCTION
// // =========================

// // Called from ChatbotUpload.html via google.script.run
// // formData contains text fields + file objects
// function submitChatbotApplicationWithFiles(formData) {
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
//   if (!dataSheet) {
//     throw new Error('Sheet "' + DATA_SHEET_NAME + '" not found.');
//   }

//   // 1) Read basic inputs
//   var applicantId = formData.applicantId;
//   var industry    = formData.industry;
//   var amount      = formData.amount;
//   var category    = formData.category;

//   if (!applicantId || !industry || !amount || !category) {
//     throw new Error('Please fill Applicant ID, Industry, Amount and Category.');
//   }

//   // 2) Save uploaded documents to Drive
//   // Front‑end already restricts extensions and size
//   var kycUrl      = saveUploadedFile(formData.kycFile, 'KYC');
//   var incomeUrl   = saveUploadedFile(formData.incomeFile, 'INCOME');
//   var businessUrl = saveUploadedFile(formData.businessFile, 'BUSINESS');

//   // 3) Translate presence to Yes / No flags for rule engine
//   var incomeDoc = incomeUrl   ? 'Yes' : 'No';
//   var kyc       = kycUrl      ? 'Yes' : 'No';
//   var business  = businessUrl ? 'Yes' : 'No';

//   // 4) Append a new row in Loan Application DataSet
//   var lastRow = dataSheet.getLastRow();   // last used row
//   var nextRow = lastRow + 1;              // new row index

//   // Columns:
//   // B: S.No.
//   // C: Applicant ID
//   // D: Applicant's Industry
//   // E: Loan Amount Requested
//   // F: Loan Category
//   // G: Applicant's Category
//   // H: Income Document Submitted
//   // I: KYC Submitted
//   // J: Business Proof Submitted

//   dataSheet.getRange(nextRow, 2).setValue(nextRow - 1);  // S.No.
//   dataSheet.getRange(nextRow, 3).setValue(applicantId);
//   dataSheet.getRange(nextRow, 4).setValue(industry);
//   dataSheet.getRange(nextRow, 5).setValue(amount);
//   dataSheet.getRange(nextRow, 6).setValue('SME');        // fixed
//   dataSheet.getRange(nextRow, 7).setValue(category);
//   dataSheet.getRange(nextRow, 8).setValue(incomeDoc);
//   dataSheet.getRange(nextRow, 9).setValue(kyc);
//   dataSheet.getRange(nextRow,10).setValue(business);

//   // Optional: store file URLs in P/Q/R (cols 16–18)
//   dataSheet.getRange(nextRow,16).setValue(kycUrl);
//   dataSheet.getRange(nextRow,17).setValue(incomeUrl);
//   dataSheet.getRange(nextRow,18).setValue(businessUrl);

//   // 5) Copy formulas for rule engine (K–O) from previous row
//   // Assumes row 2 already has working formulas in K2:O2
//   if (lastRow >= 2) {
//     dataSheet.getRange(lastRow, 11, 1, 5)   // K..O on lastRow
//       .copyTo(
//         dataSheet.getRange(nextRow, 11, 1, 5), // K..O on nextRow
//         { contentsOnly: false }
//       );
//   }

//   // 6) Read final status from column O and return to sidebar
//   var status = dataSheet.getRange(nextRow, 15).getValue();  // col 15 = O
//   return status;
// }
