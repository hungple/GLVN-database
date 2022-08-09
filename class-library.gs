var RELEASE = "20220809"
var DECIMAL_COL_LEN = 5; //IMPORTANT: any point column must has name with lenght = 5. For example: Part1 or Hwrk1
var POINT_COL_START = 8;
var EMAIL_BODY      = "Mến chào quí Phụ Huynh,<br>Xin phụ huynh xem phiếu điểm đính kèm. Xin cám ơn.<br>Chương Trình GLVN - MHT.";

var idCol          = 1;
var fNameCol       = 3;
var lNameCol       = 4;
var pEMailCol      = 5;
var totalPointsCol = 6;
var actionCol      = 7;
  

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
  return sheet.getRange("B2:B2").getCell(1, 1).getValue();
}


function getReportCardFolderId() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("admin");
  return sheet.getRange("B3:B3").getCell(1, 1).getValue();
}

function getStr(key) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("adminGrades");
  var range = sheet.getRange(2, 1, 20, 2); //row, col, numRows, numCols

  // iterate through all cells in the range
  for (var cellRow = 1; cellRow <= range.getHeight(); cellRow++) {
    var varName = range.getCell(cellRow, 1).getValue();
    if( varName == key){
      return range.getCell(cellRow, 2).getValue();
    }
  }
  return "";
}

function getGradePoint(key) {
  return parseInt(getStr(key).substring(1));
}

function log(message) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("log");
  if (sheet != null) {
    var textlogCell = sheet.getRange("A1:A1").getCell(1, 1);
    var text = textlogCell.getValue();
    textlogCell.setValue(text + "\n" + message);
  }
}

function logClear() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("log");
  if (sheet != null) {
    var textlogCell = sheet.getRange("A1:A1").getCell(1, 1);
    textlogCell.setValue("");
  }
}

//================================================================================
function getSignature(names) {
  var tNames = names.split(",");
  if(tNames.length > 1) {
    var tName1 = tNames[1];
    if(tName1.indexOf("sign:")>=0) {
      return tName1.split(":")[1];
    }
  }
  return "";
}

//================================================================================
function createDoc(isHK2) {
  logClear();

  var sendEmail = false;
  var colNames = [];
  var colPoints = [];

  var folerId     = getReportCardFolderId();

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("grades");
  var range = sheet.getRange(1, 1, 70, 30); //row, col, numRows, numCols
  //////////////////////////////////////////////////////////////////////////////

  var cName = range.getCell(1, 1).getValue();
  var tNames = range.getCell(1, 2).getValue();

  var tmpName     = "HK1-Report-Card";
  if(isHK2) {
    tmpName     = "HK2-Report-Card";
  }

  // iterate through all column names starting from the column right after the column `action`
  for (var cellCol = POINT_COL_START; ; cellCol++) {
    var colName = range.getCell(2, cellCol).getValue();
    if(colName == "") { break; }
    colNames[cellCol-POINT_COL_START] = colName;
  }

  //Logger.log(colNames);

  var halfCol = colNames.length/2;

  var PASSING_POINT = getGradePoint("GRADE_F");

  // iterate through all rows in the range
  for (var cellRow = 3; ; cellRow++) {

    var id = range.getCell(cellRow, idCol).getValue();
    if(id == "") { break; }

    var fName = range.getCell(cellRow, fNameCol).getValue();
    var lName = range.getCell(cellRow, lNameCol).getValue();
    var email = range.getCell(cellRow, pEMailCol).getValue().trim();
    var actionCell = range.getCell(cellRow, actionCol);
    var action = actionCell.getValue();

    log(fName + ' ' + lName);

    if(action != '' && action != 'd') {
      if(action == 'e' && sendEmail == false) {
        var ui = SpreadsheetApp.getUi();

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

      var signature = getSignature(tNames);
      if(action == 's') {
        signature = ''
      }

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
      log(colPoints);

      var formId      = getReportCardTemplateId();
      //Logger.log(formId);

      // Get document template, copy it as a new temp doc, and save the Doc’s id
      var copyId = DriveApp.getFileById(formId).makeCopy(docName).getId();

      // Open the temporary document
      var copyDoc = DocumentApp.openById(copyId);

      log(copyDoc.getName());

      // Get the document’s body section
      var copyBody = copyDoc.getActiveSection();

      // Replace place holder keys,in our google doc template
      copyBody.replaceText('@cname@', cName);
      copyBody.replaceText('@tname@', tNames.split(",")[0]);
      copyBody.replaceText('@sname@', fName + ' ' + lName);


      // HK1 - fill in data for HK1
      var hk1Total = 0;
      for (var i = 0; i<halfCol-1; i++) {
        if(colNames[i].length == DECIMAL_COL_LEN) { //columns that hold points
          //Logger.log(colNames[i]);
          if(typeof(colPoints[i]) == 'number') {
            copyBody.replaceText('@' + colNames[i] + '@', colPoints[i].toFixed(2));
            hk1Total = hk1Total + parseFloat(colPoints[i]);
          }
          else {
            //copyBody.replaceText('@' + colNames[i] + '@', '-');
            var ui = SpreadsheetApp.getUi();

            var response = ui.alert(
              'Error!!!',
              'Row ' + cellRow + ' has a blank character or an invalid number "' + colPoints[i] + '"',
              ui.ButtonSet.OK
            );
            return;
          }
        }
        else { //columns that hold text or attendance
          copyBody.replaceText('@' + colNames[i] + '@', colPoints[i]);
        }
      }
      copyBody.replaceText('@Total1@', hk1Total.toFixed(2));
      var c1 = colPoints[halfCol-1];
      if(c1.length < 70) {
        c1 = c1 + "\n\n";
      }
      else if(c1.length < 140) {
        c1 = c1 + "\n";
      }

      copyBody.replaceText('@Comment1@', c1);
      copyBody.replaceText('@sign1@', signature);

      // HK2
      var hk2Total = 0;
      if(isHK2) { // fill in data for HK2
        for (var i = halfCol; i<colNames.length-1; i++) {
          if(colNames[i].length == DECIMAL_COL_LEN) { //columns that hold points
            //Logger.log(colNames[i]);
            if(typeof(colPoints[i]) == 'number') {
              copyBody.replaceText('@' + colNames[i] + '@', colPoints[i].toFixed(2));
              hk2Total = hk2Total + parseFloat(colPoints[i]);
            }
            else {
              var ui = SpreadsheetApp.getUi();

              var response = ui.alert(
                'Error!!!',
                'Row ' + cellRow + ' has a blank character or an invalid number "' + colPoints[i] + '"',
                ui.ButtonSet.OK
              );
              return;
            }
          }
          else { //columns that hold text or attendance
            copyBody.replaceText('@' + colNames[i] + '@', colPoints[i]);
          }
        }
        copyBody.replaceText('@Total2@', hk2Total.toFixed(2));
        var c2 = colPoints[colNames.length-1];
        if(c2.length < 70) {
          c2 = c2 + "\n\n";
        }
        else if(c2.length < 140) {
          c2 = c2 + "\n";
        }

        copyBody.replaceText('@Comment2@', c2);
        copyBody.replaceText('@sign2@', signature);

        // fill in data for Yearly Total
        for (var i = 0; i<halfCol-1; i++) {
          if(colNames[i].length == DECIMAL_COL_LEN) {
            var tempTotal = colPoints[i]+colPoints[i+halfCol];
            if(typeof(tempTotal) == 'number') {
              copyBody.replaceText('@' + colNames[i].substring(0,colNames[i].length-1) + '3@', tempTotal.toFixed(2));
            }
            else {
              copyBody.replaceText('@' + colNames[i].substring(0,colNames[i].length-1) + '3@', '-');
            }
          }
          else {
            copyBody.replaceText('@' + colNames[i].substring(0,colNames[i].length-1) + '3@', colPoints[i]+colPoints[i+halfCol]);
          }
        }
        copyBody.replaceText('@Total3@', (hk1Total + hk2Total).toFixed(2));
        if (hk1Total + hk2Total >= PASSING_POINT) {
          copyBody.replaceText('@Pass@', 'Được lên lớp - Passed');
        }
        else {
          copyBody.replaceText('@Pass@', 'Ở lại lớp - Does Not Pass');
        }
      }
      else { // fill in '-' for HK2 because this is processed for HK1
        for (var i = halfCol; i<colNames.length-1; i++) {
          copyBody.replaceText('@' + colNames[i] + '@', '-');
        }
        copyBody.replaceText('@Total2@', '-');
        copyBody.replaceText('@Comment2@', "\n");
        copyBody.replaceText('@sign2@', '');

        // fill in '-' for Yearly Total
        for (var i = 0; i<halfCol-1; i++) {
          copyBody.replaceText('@' + colNames[i].substring(0,colNames[i].length-1) + '3@', '-');
        }
        copyBody.replaceText('@Total3@', '-');
        copyBody.replaceText('@Pass@', '');
      }


      // Save and close the temporary document
      copyDoc.saveAndClose();

      // Convert temporary document to PDF
      var pdf = DriveApp.getFileById(copyId).getAs("application/pdf");

      // Delete temp file
      DriveApp.getFileById(copyId).setTrashed(true);

      // Delete old file
      log("Delete old file");
      var files = DriveApp.getFolderById(folerId).getFilesByName(docName + ".pdf");
      while (files.hasNext()) {
        var file = files.next();
        if(file.getOwner().getEmail() == Session.getActiveUser()) {
          file.setTrashed(true);
        }
      }

      // Save pdf
      log("Create pdf");
      DriveApp.getFolderById(folerId).createFile(pdf);

      // Send email
      if(action == 'e' && email && email.length > 5) {
        // Attach PDF and send the email
        var subject = docName;
        // email = "hle007@yahoo.com";
        MailApp.sendEmail(email, subject, EMAIL_BODY, {htmlBody: EMAIL_BODY, attachments: pdf});
      }

      actionCell.setValue('');
    }
  }
}
