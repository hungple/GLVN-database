var RELEASE = "20240523"


var DECIMAL_COL_LEN      = 5; //IMPORTANT: any point column must has a name which its lenght = 5. For example: Part1 or Hwrk1
var EXTRA_CREDIT_COL_LEN = 6;
var POINT_COL_START      = 8;
var EMAIL_BODY           = "Mến chào quý Phụ Huynh,<br>Xin quý phụ huynh xem phiếu điểm đính kèm. Xin cám ơn.<br>Chương Trình GLVN.";

var idCol          = 1;
var fNameCol       = 3;
var lNameCol       = 4;
var pEMailCol      = 5;
var totalPointsCol = 6;
var actionCol      = 7;
  
// For debugging purpose
//var DEBUG_SPREADSHEET_ID = "1eyllhOnvlg7077oN7KFD8gWRG8uZlqSttJM3HsgOJno"; // MHT GL8B
var DEBUG_SPREADSHEET_ID = "1sSfaHAOZo0GMUDjOxH5czbSilGvn3ukm8v9rc77_2NM"; // MHT VN6A

/**
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Create HK1 report cards",
    functionName : "createReports_1"
  },
  {
    name : "Create HK2 report cards",
    functionName : "createReports_2"
  },
  {
    name : "Release: " + RELEASE,
    functionName : "showRelease"
  }  
  ];
  sheet.addMenu("GLVN", entries);
};


function createReports_1() {
  createDoc(false);
}

function createReports_2() {
  createDoc(true);
}

function showRelease() {
  var ui = SpreadsheetApp.getUi();
  
  var response = ui.alert(
      'Information!!!',
      'Release: ' + RELEASE,
      ui.ButtonSet.OK);
}


//================================================================================

function getReportCardTemplateId() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("admin");
  if (sheet == null) { // for debugging
    var spreadSheet = SpreadsheetApp.openById(DEBUG_SPREADSHEET_ID);
    sheet = spreadSheet.getSheetByName("admin");
  }
  return sheet.getRange("B2:B2").getCell(1, 1).getValue();
}


function getReportCardFolderId() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("admin");
  if (sheet == null) { // for debugging
    var spreadSheet = SpreadsheetApp.openById(DEBUG_SPREADSHEET_ID);
    sheet = spreadSheet.getSheetByName("admin");
  }
  return sheet.getRange("B3:B3").getCell(1, 1).getValue();
}


function getSignature() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("grades");
  if (sheet == null) { // for debugging
    var spreadSheet = SpreadsheetApp.openById(DEBUG_SPREADSHEET_ID);
    sheet = spreadSheet.getSheetByName("grades");
  }
  return sheet.getRange("F1:F1").getCell(1, 1).getValue();
}


function getLetterGrade(point, maxPoint) {
  if (point/maxPoint >= .9) {
    return "A";
  }
  else if (point/maxPoint >= .8) {
    return "B";
  }
  else if (point/maxPoint >= .7) {
    return "C";
  }
  else if (point/maxPoint >= .65) {
    return "D";
  }
  else {
    return "F";
  }
}

//================================================================================
function createDoc(isHK2) {
  var ui = SpreadsheetApp.getUi();

  var sendEmail = false;
  var colNames = [];
  var colPoints = [];
  var colMins = [];
  var colMaxs = [];
  
  var folerId     = getReportCardFolderId();

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("grades");
  if (sheet == null) { // for debugging
    var spreadSheet = SpreadsheetApp.openById(DEBUG_SPREADSHEET_ID);
    sheet = spreadSheet.getSheetByName("grades");
  }
  var range = sheet.getRange(1, 1, 70, 30); //row, col, numRows, numCols
  //////////////////////////////////////////////////////////////////////////////

  var cName = range.getCell(1, 1).getValue();
  var tNames = range.getCell(1, 2).getValue();

  var tmpName     = "HK1-Report-Card";
  if(isHK2) {
    tmpName     = "HK2-Report-Card";
  }
  
  // iterate through all range (min-max) column starting from the column right after the column `action`
  for (var cellCol = POINT_COL_START; ; cellCol++) {
    var colRange = range.getCell(1, cellCol).getValue();
    if(colRange == "") { break; }
    var minMaxArr = colRange.split("-");
    colMins[cellCol-POINT_COL_START] = minMaxArr[0];
    colMaxs[cellCol-POINT_COL_START] = minMaxArr[1];
  }

  // iterate through all column names starting from the column right after the column `action`
  for (var cellCol = POINT_COL_START; ; cellCol++) {
    var colName = range.getCell(2, cellCol).getValue();
    if(colName == "") { break; }
    colNames[cellCol-POINT_COL_START] = colName;
  }

  var halfCol = colNames.length/2;

  // iterate through all rows in the range
  for (var cellRow = 3; ; cellRow++) {
    
    var id = range.getCell(cellRow, idCol).getValue(); 
    if(id == "") { break; }
    
    var fName = range.getCell(cellRow, fNameCol).getValue(); 
    var lName = range.getCell(cellRow, lNameCol).getValue(); 
    var email = range.getCell(cellRow, pEMailCol).getValue().trim();
    var actionCell = range.getCell(cellRow, actionCol);
    var action = actionCell.getValue().trim();
    
    if(action != '') {
      if(action == 'e' && sendEmail == false) {
        var response = ui.alert(
          'Warning!!!',
          'Do you want to email to the report cards to the parents?',
          ui.ButtonSet.YES_NO
        );

        // Process the user's response.
        if (response == ui.Button.YES) {
          sendEmail = true;
        }
        else {
          return;
        }
      }

      var signature = getSignature();

      var docName;
      if(email == "" || email.length < 6) {
        docName = '--no-email-' + cName + '-' + fName + '-' + lName + '-' + id + '-' + tmpName;
      }
      else {
        docName = cName + '-' + fName + '-' + lName + '-' + id + '-' + tmpName;
      }
    
      // iterate throught all columns
      for (var cCol = 0; cCol<colNames.length; cCol++) {
        colPoints[cCol] = range.getCell(cellRow, cCol+POINT_COL_START).getValue();
      }      
      
      var formId      = getReportCardTemplateId();
      
      // Get document template, copy it as a new temp doc, and save the Doc’s id
      var copyId = DriveApp.getFileById(formId).makeCopy(docName).getId();

      // Open the temporary document
      var copyDoc = DocumentApp.openById(copyId);

      // Get the document’s body section
      var copyBody = copyDoc.getActiveSection();

      // Replace place holder keys,in our google doc template
      copyBody.replaceText('@cname@', cName);
      copyBody.replaceText('@tname@', tNames); //tNames.split(",")[0]
      copyBody.replaceText('@sname@', fName + ' ' + lName);
      

      // HK1 - fill in data for HK1
      var hk1Total = 0;
      for (var i = 0; i<halfCol-1; i++) {
        if (colNames[i].length == DECIMAL_COL_LEN) { //columns that hold points
          if(typeof(colPoints[i]) == 'number') {
            if (colPoints[i] >= colMins[i] && colPoints[i] <= colMaxs[i]) {
              copyBody.replaceText('@' + colNames[i] + '@', colPoints[i].toFixed(2));
              hk1Total = hk1Total + parseFloat(colPoints[i]);
            }
            else {
              ui.alert(
                'Error!!!',
                'Row ' + cellRow + ' has an out of range number: "' + colPoints[i] + '"',
                ui.ButtonSet.OK
              );
              return;
            }
          }
          else {
            ui.alert(
              'Error!!!',
              'Row ' + cellRow + ' has a blank character or a non-digit character: "' + colPoints[i] + '"',
              ui.ButtonSet.OK
            );
            return;
          }
        }
        else if(colNames[i].length == EXTRA_CREDIT_COL_LEN) { 
          if(typeof(colPoints[i]) == 'number') {
            copyBody.replaceText('@' + colNames[i] + '@', colPoints[i].toFixed(2));
            hk1Total = hk1Total + parseFloat(colPoints[i]);
          }
          else 
          {
            if(colPoints[i] != '') {
              ui.alert(
                'Error!!!',
                'Row ' + cellRow + ' has an invalid number "' + colPoints[i] + '"',
                ui.ButtonSet.OK
              );
              return;
            }
            else {
              copyBody.replaceText('@' + colNames[i] + '@', "");
            }
          }
        }
        else { //columns that hold text or attendance
          copyBody.replaceText('@' + colNames[i] + '@', colPoints[i]);
        }
      }
      copyBody.replaceText('@Total1@', getLetterGrade(hk1Total, 50));
      var c1 = colPoints[halfCol-1];
      if(c1.length < 70) {
        c1 = c1 + "\n\n";
      }
      else if(c1.length < 140) {
        c1 = c1 + "\n";
      }
      
      copyBody.replaceText('@Comment1@', c1);
      copyBody.replaceText('@Sign1@', signature);
      
      // HK2
      var hk2Total = 0;
      if(isHK2) { // fill in data for HK2
        for (var i = halfCol; i<colNames.length-1; i++) {
          if(colNames[i].length == DECIMAL_COL_LEN) { //columns that hold points
            if(typeof(colPoints[i]) == 'number') {
              if (colPoints[i] >= colMins[i] && colPoints[i] <= colMaxs[i]) {
                copyBody.replaceText('@' + colNames[i] + '@', colPoints[i].toFixed(2));
                hk2Total = hk2Total + parseFloat(colPoints[i]);
              }
              else {
                ui.alert(
                  'Error!!!',
                  'Row ' + cellRow + ' has an out of range number: "' + colPoints[i] + ' - HK2"',
                  ui.ButtonSet.OK
                );
                return;
              }
            }
            else {
              ui.alert(
                'Error!!!',
                'Row ' + cellRow + ' has a blank character or a non-digit character: "' + colPoints[i] + '"',
                ui.ButtonSet.OK
              );
              return;
            }
          }
          else if(colNames[i].length == EXTRA_CREDIT_COL_LEN) { 
            if(typeof(colPoints[i]) == 'number') {
              copyBody.replaceText('@' + colNames[i] + '@', colPoints[i].toFixed(2));
              hk2Total = hk2Total + parseFloat(colPoints[i]);
            }
            else 
            {
              if(colPoints[i] != '') {
                ui.alert(
                  'Error!!!',
                  'Row ' + cellRow + ' has an invalid number "' + colPoints[i] + '"',
                  ui.ButtonSet.OK
                );
                return;
              }
              else {
                copyBody.replaceText('@' + colNames[i] + '@', "");
              }
            }
          }
          else { //columns that hold text or attendance
            copyBody.replaceText('@' + colNames[i] + '@', colPoints[i]);
          }
        }
        copyBody.replaceText('@Total2@', getLetterGrade(hk2Total, 50));
        var c2 = colPoints[colNames.length-1];
        if(c2.length < 70) {
          c2 = c2 + "\n\n";
        }
        else if(c2.length < 140) {
          c2 = c2 + "\n";
        }

        copyBody.replaceText('@Comment2@', c2);
        copyBody.replaceText('@Sign2@', signature);
        
        // fill in data for Yearly Total
        for (var i = 0; i<halfCol-1; i++) {
          if(colNames[i].length == DECIMAL_COL_LEN || colNames[i].length == EXTRA_CREDIT_COL_LEN) {
            var tempTotal = 0;
            if(typeof(colPoints[i]) == 'number') {
              tempTotal = colPoints[i];
            }
            if(typeof(colPoints[i+halfCol]) == 'number') {
              tempTotal = tempTotal + colPoints[i+halfCol];
            }
            if(tempTotal > 0) {
              copyBody.replaceText('@' + colNames[i].substring(0,colNames[i].length-1) + '3@', tempTotal.toFixed(2));
            }
            else {
              copyBody.replaceText('@' + colNames[i].substring(0,colNames[i].length-1) + '3@', '');
            }
          }
          else {
            copyBody.replaceText('@' + colNames[i].substring(0,colNames[i].length-1) + '3@', colPoints[i]+colPoints[i+halfCol]);
          }
        }
        copyBody.replaceText('@Total3@', getLetterGrade(hk1Total + hk2Total, 100));

      }
      else { // fill in '-' for HK2 because this is processed for HK1
        for (var i = halfCol; i<colNames.length-1; i++) {
          copyBody.replaceText('@' + colNames[i] + '@', '-');
        }
        copyBody.replaceText('@Total2@', '-');
        copyBody.replaceText('@Comment2@', "\n");
        copyBody.replaceText('@Sign2@', '');
        
        // fill in '-' for Yearly Total
        for (var i = 0; i<halfCol-1; i++) {
          copyBody.replaceText('@' + colNames[i].substring(0,colNames[i].length-1) + '3@', '-');
        }
        copyBody.replaceText('@Total3@', '-');
      }
      
 
      // Save and close the temporary document
      copyDoc.saveAndClose();

      // Convert temporary document to PDF
      var pdf = DriveApp.getFileById(copyId).getAs("application/pdf");
 
      // Delete temp file
      DriveApp.getFileById(copyId).setTrashed(true);

      // Delete old file
      var files = DriveApp.getFolderById(folerId).getFilesByName(docName + ".pdf");
      while (files.hasNext()) {
        var file = files.next();
        if(file.getOwner().getEmail() == Session.getActiveUser()) {
          file.setTrashed(true); 
        }        
      }

      // Save pdf
      DriveApp.getFolderById(folerId).createFile(pdf);

      // Send email
      if(action == 'e' && email && email.length > 5) {
        // Attach PDF and send the email
        var subject = docName;
        MailApp.sendEmail(email, subject, EMAIL_BODY, {htmlBody: EMAIL_BODY, attachments: pdf});
      }

      actionCell.setValue('');
    }
  }
}


