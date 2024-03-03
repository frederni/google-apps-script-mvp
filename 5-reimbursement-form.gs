const NAME;
const DATE;
const PROJECT;
const ACCTNO;
const COMMENT;
const COL_DESCRIPTION; // 1-indexed
const COL_AMOUNT; // 1-indexed
const OFFSET_DESCRIPTIONS;
const ID_FOLDER;
const SHEET_NAME_EXPORT;
const SHEET_NAME_TEMPLATE;
const SHEET_NAME_RESPONSES;
const EMAIL_RECIPIENT;

function getSurveyColumnMap(){
  // Maps survey response from question name to 0-based column number (A=0, B=1 etc.)
  // This should be mapped differently based on your survey and questions. Timestamp is always column A
  return {
    "Timestamp": 0,
    "Name": 1,
    "Date": 2,
    "Project": 3,
    "Acctno": 4,
    "Description": 5,
    "Amount": 6,
    "Receipt": 7,
    "Comment": 8,
    "SentFlag": 9
  };
}

function clearTemplateSheet() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templateSheet = ss.getSheetByName(SHEET_NAME_TEMPLATE);
  // Clears existing data from the template.
  const rngClear = templateSheet.getRangeList([NAME, DATE, PROJECT, ACCTNO, COMMENT]).getRanges()
  rngClear.forEach(function (cell) {
    cell.clearContent();
  });
  // This sample only accounts for six rows of data 'B15:H20'. You can extend or make dynamic as necessary.
  templateSheet.getRange("B15:H20").clearContent();
}

function fillTemplate(row){
  Logger.log("Filling out template for row");

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ts = ss.getSheetByName(SHEET_NAME_TEMPLATE);
  clearTemplateSheet();
  const columnMap = getSurveyColumnMap();

  // Fill template with survey response
  ts.getRange(NAME).setValue(row[columnMap["Name"]]);
  ts.getRange(DATE).setValue(row[columnMap["Date"]]);
  ts.getRange(PROJECT).setValue(row[columnMap["Project"]]);
  ts.getRange(ACCTNO).setValue(row[columnMap["Acctno"]].toString(10)); // Cast as string to avoid loving prefixed 0
  ts.getRange(COMMENT).setValue(row[columnMap["Comment"]]);
  
  let costAmount = row[columnMap["Amount"]].toString(10).split("\n");
  let costDescription = row[columnMap["Description"]].split("\n");

  // Maxed at 6 rows, but can be extended with a modified template:
  let numCostRows = Math.min(6, Math.max(costDescription.length, costAmount.length));
  for(let i=0; i< numCostRows; i++){
    ts.getRange(OFFSET_DESCRIPTIONS + i, COL_DESCRIPTION).setValue(costDescription[i] ?? ""); 
    ts.getRange(OFFSET_DESCRIPTIONS + i, COL_AMOUNT).setValue(costAmount[i] ?? "");
  }
}

function exportSheetToPdf(sheet_name = SHEET_NAME_EXPORT, pdfName = null){
  /* Method heavily based on samples from Google, see
  https://developers.google.com/apps-script/samples/automations/generate-pdfs
  */
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ssId = ss.getId();
  let sheet = ss.getSheetByName(sheet_name);
  if(pdfName == null){
    let now = new Date();
    pdfName = "utlegg-" + now.toISOString().split('T')[0];
  }
  const fr = 0, fc = 0, lc = 9, lr = 27;
  const url = "https://docs.google.com/spreadsheets/d/" + ssId + "/export" +
    "?format=pdf&" +
    "size=7&" +
    "fzr=true&" +
    "portrait=true&" +
    "fitw=true&" +
    "gridlines=false&" +
    "printtitle=false&" +
    "top_margin=0.5&" +
    "bottom_margin=0.25&" +
    "left_margin=0.5&" +
    "right_margin=0.5&" +
    "sheetnames=false&" +
    "pagenum=UNDEFINED&" +
    "attachment=true&" +
    "gid=" + sheet.getSheetId() + '&' +
    "r1=" + fr + "&c1=" + fc + "&r2=" + lr + "&c2=" + lc;

  const params = { method: "GET", headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() } };
  const blob = UrlFetchApp.fetch(url, params).getBlob().setName(pdfName + '.pdf');

  // Gets the folder in Drive where the PDFs are stored.
  const folder = DriveApp.getFolderById(ID_FOLDER);
  const pdfFile = folder.createFile(blob);
  return pdfFile;
}

function sendEmail(pdfFile, row){
  const colMap = getSurveyColumnMap();
  let receiptIds = row[colMap["Receipt"]].split(", ").map(url => url.split("?id=")[1]).filter(id => id);
  let receiptFiles = receiptIds.map(id => DriveApp.getFileById(id)).filter(fileobj => fileobj);
  let attachments = [pdfFile, ...receiptFiles].map(fileobj => fileobj.getAs(fileobj.getMimeType())).filter(pdf => pdf);
  
  GmailApp.sendEmail(
    EMAIL_RECIPIENT, 
    'Nytt utlegg fra ' + row[colMap["Name"]],
    'Utlegg fra Google Form (innsendt ' + row[colMap["Timestamp"]] + ')',
    {
      attachments: attachments,
      name: 'Utleggssystem'
    });
}

function main(){
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_RESPONSES);
  const colMap = getSurveyColumnMap();
  let surveyResponses = sheet.getDataRange().getValues();
  for (let i = 1; i < surveyResponses.length; i++){
    let row = surveyResponses[i];
    if (row[colMap["SentFlag"]] != true && row[colMap["Timestamp"]] != '') {
      Logger.log("i = " + i);
      fillTemplate(row);
      Logger.log("Sucessfully filled out template. Exporting to PDF.");
      SpreadsheetApp.flush();
      Utilities.sleep(500);
      let pdfFile = exportSheetToPdf("Utleggsmal");
      Logger.log("Sucessfully exported PDF! Sending email with attachments.");
      sendEmail(pdfFile, row);
      sheet.getRange(i+1, colMap["SentFlag"]+1).setValue(true);
      Logger.log("Sucessfully sent email!");
    }
  }
}
