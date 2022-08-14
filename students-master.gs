var RELEASE = "20220814"

// Std_VGz6v3
var idCol               = 1;
var glGCol              = 7;
var glNCol              = 8;
var vnGCol              = 9;
var vnNCol              = 10;
var isRegCol            = 11;

var euchDateCol         = 26; // Z
var euchLocationCol     = 27; // AA
var confDateCol         = 28; // AB
var confLocationCol     = 29; // AC

var glFinalPointCol     = 33; // AG
var vnFinalPointCol     = 34; // AH


// gl-classes/vn-classes
var clsNameCol          = 1;
var gmailCol            = 6;
var actionCol           = 7;
var clsFolderIdCol      = 9;


// class: GL1A, VN1A
var class_idCol         = 2; // col B
var class_totalPointsCol= 15; // col O

var MAX_HONOR_ROLL = 20;


/////////////////////////////////////////////////////////////////////////////////////////////////////
//
// GLVN menu item
//
/////////////////////////////////////////////////////////////////////////////////////////////////////

/**
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [
    {
      name : "Share classes (gl-classes/vn-classes)",
      functionName : "shareClasses"
    },
    {
      name : "1 - Save student final points (Std)",
      functionName : "saveFinalPoints"
    },
    {
      name : "2 - Save First Communion date and location (Std)",
      functionName : "saveFCommunionInfo"
    },
    {
      name : "3 - Save Confirmation date and location (Std)",
      functionName : "saveConfirmationInfo"
    },
    {
      name : "4 - Un-share classes (gl-classes/vn-classes)",
      functionName : "unShareClasses"
    },
    {
      name : "5 - Save students into students-past folder",
      functionName : "saveStudentsPast"
    },
    {
      name : "6 - Increase student glG and vnG for new registration (Std)",
      functionName : "increaseGlGVnGForNewReg"
    },
    {
      name : "7 - Clear data in external classes (gl-classes/vn-classes)",
      functionName : "clearDataExternalClasses"
    },
    // DO NOT DELETE THESE TWO FUNCTIONS
    {
     name : "Update class sheets in this student-master (gl-classes/vn-classes : root only)",
     functionName : "updateSheetsInThisSpreadSheet"
    },
    {
     name : "Clone classes using GL1A or VN1A (gl-classes/vn-classes : root only)",
     functionName : "cloneClassesUsingGL1AorVN1A"
    },
    {
     name : "Update external class spreadSheets (gl-classes/vn-classes : root only)",
     functionName : "updateExternalClassSpreadSheets"
    },

    {
      name : "Release: " + RELEASE,
      functionName : "showRelease"
    }

    ];
  sheet.addMenu("GLVN", entries);
};

function showRelease() {
  var ui = SpreadsheetApp.getUi();

  var response = ui.alert(
      'Information!!!',
      'Release: ' + RELEASE,
      ui.ButtonSet.OK);
}

/////////////////////////////////////////////////////////////////////////////////////////////////////
//
// Utilities
//
/////////////////////////////////////////////////////////////////////////////////////////////////////
function getStr(key) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Admin");
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

function getPassingPoint() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("adminGrades");
  var fGrade = sheet.getRange("B6:B6").getCell(1, 1).getValue().slice(-2);

  return parseInt(fGrade);
}

// function test_getPassingPoint() {
//   var temp = getPassingPoint();
//   var t = "hello";
// }

function removeLastComma(strng){
  if(strng[strng.length-1] === ',')
    return strng.substring(0,strng.length-1);
  return strng;
}


function getExternalClassSpreadsheetId(clsName, clsFolder) {
  var files = clsFolder.getFilesByName(clsName);
  if (files.hasNext()) {
    var file = files.next();
    return file.getId();
  }
  return "";
}

function getReportCardsFolderId(clsName, clsFolder) {
  var folders = clsFolder.getFoldersByName(clsName + "-Report-Cards");
  if (folders.hasNext()) {
    var folder = folders.next();
    return folder.getId();
  }
  return "";
}





/////////////////////////////////////////////////////////////////////////////////////////////////////
//
// Share classes to the teachers
//
/////////////////////////////////////////////////////////////////////////////////////////////////////
function shareClasses() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var response = ui.alert(
      'Warning!!!',
      'Do you want to share classes to the teachers?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {
    var glReportCardTemplateId = getStr("GL_REPORT_CARD_TEMPLATE_ID");
    var classLibraryId         = getStr("CLASS_LIBRARY_ID");
    shareClassesImpl("gl-classes", glReportCardTemplateId, classLibraryId, true);
    var vnReportCardTemplateId = getStr("VN_REPORT_CARD_TEMPLATE_ID");
    var classLibraryId         = getStr("CLASS_LIBRARY_ID");
    shareClassesImpl("vn-classes", vnReportCardTemplateId, classLibraryId, true);
  }
}

function shareClassesImpl(sheetName, reportFormId, classLibraryId, isShared) {

  var admins = getStr("ADMIN_IDS").split(",");
  for (var i = 0; i < admins.length; i++) {
    admins[i] = admins[i].trim();
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var range = sheet.getRange(2, 1, 25, 15); //row, col, numRows, numCols

  var clsName, gmails, clsFolder, action;

  // iterate through all cells in the range
  for (var cellRow = 1; cellRow <= range.getHeight(); cellRow++) {
    clsName = range.getCell(cellRow, clsNameCol).getValue();
    gmails = range.getCell(cellRow, gmailCol).getValue().trim();

    if( clsName == "")
      break;

    var actionCell = range.getCell(cellRow, actionCol);
    action = actionCell.getValue().trim();

    if(action == "x" && gmails != "") {

      clsFolder = DriveApp.getFolderById(range.getCell(cellRow, clsFolderIdCol).getValue());

      var gmailArr = removeLastComma(gmails).split(",");
      for (i = 0; i < gmailArr.length; i++) {
        gmailArr[i] = gmailArr[i].trim();
      }

      // Share report card template and class library
      var doc = DocumentApp.openById(reportFormId);
      var libSpreadSheet = SpreadsheetApp.openById(classLibraryId);

      for (var i = 0; i < gmailArr.length; i++) {
        var gmail = gmailArr[i];
        try {
          //Only remove this gmail but don't remove other viewers
          doc.removeViewer(gmail);
          libSpreadSheet.removeViewer(gmail);

          if(isShared == true){
            doc.addViewer(gmail);
            libSpreadSheet.addEditor(gmail);
          }
        }
        catch(e) {
        } //ignore error
      }

      // Share the whole class folder
      try {
        var editors = clsFolder.getEditors();
        for (var j = 0; j < editors.length; j++) {
          if(isNotAdmin(admins, editors[j].getEmail())){
            clsFolder.removeEditor(editors[j].getEmail());
          }
        }

        // add new editor
        for (var i = 0; i < gmailArr.length; i++) {
          var gmail = gmailArr[i];
          if(isShared == true){
            clsFolder.addEditor(gmail);
          }
        }
      }
      catch(e) {
      } //ignore error

      actionCell.setValue('');
    }
  }
};

function isNotAdmin(admins, gmail) {
  for (var i = 0; i < admins.length; i++) {
    if(admins[i].toUpperCase() === gmail.toUpperCase()) {
      return false;
    }
  }
  return true;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////
//
// 1 - Save Final Points
//
/////////////////////////////////////////////////////////////////////////////////////////////////////
function saveFinalPoints() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var response = ui.alert(
      'Warning!!!',
      'Do you want to update pass/fail for students?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {
    saveFinalPointsImpl();
  }
}


function saveFinalPointsImpl() {

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Std_VGz6v3");
  var range = sheet.getRange(1, 1, 700, 34); //row, col, numRows, numCols
  var rowStartCell = sheet.getRange("AI1:AI1").getCell(1, 1);

  ////////////////////////////////////////////////////////////////////////////////////
  ////////////////////////////////////////////////////////////////////////////////////

  var rowStart = rowStartCell.getValue()+1;

  var id, isReg, glLevel, glName, vnLevel, vnName;
  // iterate through all cells in the range
  for (var cellRow = rowStart; ; cellRow++) {
    id = range.getCell(cellRow, idCol).getValue();

    if(id == "") break;

    isReg   = range.getCell(cellRow, isRegCol).getValue();
    glLevel = range.getCell(cellRow, glGCol).getValue();
    glName  = range.getCell(cellRow, glNCol).getValue();
    vnLevel = range.getCell(cellRow, vnGCol).getValue();
    vnName  = range.getCell(cellRow, vnNCol).getValue();

    if(isReg == "x") {
      if(glName != "") {
        range.getCell(cellRow, glFinalPointCol).setValue(getFinalGrade("GL" + glLevel + glName, id));
      }

      if(vnName != "") {
        range.getCell(cellRow, vnFinalPointCol).setValue(getFinalGrade("VN" + vnLevel + vnName, id));
      }
    }

    // update rowStart cell
    rowStartCell.setValue(cellRow);
  }

};


function getFinalGrade(className, id) { // sheet GL1A, GL1B, GL2A...

  //========================================================================

  var sheet =SpreadsheetApp.getActiveSpreadsheet().getSheetByName(className);
  if(sheet != null) {
    var range = sheet.getRange(2, 1, 60, 20); //row, col, numRows, numCols

    var idCell, totalPointsCell;

    // iterate through all cells in the range
    for (var cellRow = 1; cellRow <= range.getHeight(); cellRow++) {
      idCell = range.getCell(cellRow, class_idCol);
      totalPointsCell = range.getCell(cellRow, class_totalPointsCol);
      if(idCell.getValue() == id) {
        return totalPointsCell.getValue();
      }
    }
  }
  return 0;
};


/////////////////////////////////////////////////////////////////////////////////////////////////////
//
// 2 - Save first communion date and location
//
/////////////////////////////////////////////////////////////////////////////////////////////////////
function saveFCommunionInfo() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var response = ui.alert(
      'Warning!!!',
      'Do you want to update communion information for students?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {
    saveFCommunionInfoImpl();
  }
}

function saveFCommunionInfoImpl() {

  // Get data from the Calendar sheet
  var varSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calendar");
  var varRange = varSheet.getRange(1, 1, 10, 11); //row, col, numRows, numCols
  var commDate = varRange.getCell(3, 2).getValue(); //row, col
  var commLocation = getStr("CHURCH_INFO");
  var passing_point = getPassingPoint();

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Std_VGz6v3");
  var range = sheet.getRange(1, 1, 700, 34); //row, col, numRows, numCols
  var rowStartCell = sheet.getRange("AI1:AI1").getCell(1, 1); // <= Need to update column

  ////////////////////////////////////////////////////////////////////////////////////
  ////////////////////////////////////////////////////////////////////////////////////

  var rowStart = rowStartCell.getValue()+1;

  var id, isReg, glLevel, glName, glFinalPoint;
  // iterate through all cells in the range
  for (var cellRow = rowStart; ; cellRow++) {
    id = range.getCell(cellRow, idCol).getValue();

    if(id == "") break;

    isReg   = range.getCell(cellRow, isRegCol).getValue();
    glLevel = range.getCell(cellRow, glGCol).getValue();
    glName  = range.getCell(cellRow, glNCol).getValue();
    glFinalPoint  = range.getCell(cellRow, glFinalPointCol).getValue();

    if(isReg == "x" && glLevel == 3 && glName != "" && glFinalPoint >= passing_point) {
      range.getCell(cellRow, euchDateCol).setValue(commDate);
      range.getCell(cellRow, euchLocationCol).setValue(commLocation);
    }

    // update rowStart cell
    rowStartCell.setValue(cellRow);
  }

};



/////////////////////////////////////////////////////////////////////////////////////////////////////
//
// 3 - Save Confirmation date and location
//
/////////////////////////////////////////////////////////////////////////////////////////////////////
function saveConfirmationInfo() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var response = ui.alert(
      'Warning!!!',
      'Do you want to save Confirmation information for students?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {
    saveConfirmationInfoImpl();
  }
}

function saveConfirmationInfoImpl() {

  //////////////////////////// Get data from the Calendar sheet  //////////////////////////////
  var varSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calendar");
  var varRange = varSheet.getRange(1, 1, 10, 11); //row, col, numRows, numCols
  var confDate = varRange.getCell(4, 2).getValue(); //row, col
  var confLocation = getStr("CHURCH_INFO");

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Std_VGz6v3");
  var range = sheet.getRange(1, 1, 700, 34); //row, col, numRows, numCols
  var rowStartCell = sheet.getRange("AI1:AI1").getCell(1, 1); // <= Need to update column

  ////////////////////////////////////////////////////////////////////////////////////
  ////////////////////////////////////////////////////////////////////////////////////


  var rowStart = rowStartCell.getValue()+1;
  var id, isReg, glLevel, glName, glFinalPoint;
  // iterate through all cells in the range
  for (var cellRow = rowStart; ; cellRow++) {
    id = range.getCell(cellRow, idCol).getValue();

    if(id == "") break;

    isReg   = range.getCell(cellRow, isRegCol).getValue();
    glLevel = range.getCell(cellRow, glGCol).getValue();
    glName  = range.getCell(cellRow, glNCol).getValue();
    glFinalPoint  = range.getCell(cellRow, glFinalPointCol).getValue();

    if(isReg == "x" && glLevel == 8 && glName != "" && glFinalPoint >= 65) {
      range.getCell(cellRow, confDateCol).setValue(confDate);
      range.getCell(cellRow, confLocationCol).setValue(confLocation);
    }

    // update rowStart cell
    rowStartCell.setValue(cellRow);
  }

};





/////////////////////////////////////////////////////////////////////////////////////////////////////
//
// 4 - Un-share GL classes to the teachers
//
/////////////////////////////////////////////////////////////////////////////////////////////////////
function unShareClasses() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var response = ui.alert(
      'Warning!!!',
      'Do you want to stop sharing classes from the teachers?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {
    var glReportCardTemplateId = getStr("GL_REPORT_CARD_TEMPLATE_ID");
    var classLibraryId         = getStr("CLASS_LIBRARY_ID");
    shareClassesImpl("gl-classes", glReportCardTemplateId, classLibraryId, false);
    var vnReportCardTemplateId = getStr("VN_REPORT_CARD_TEMPLATE_ID");
    var classLibraryId         = getStr("CLASS_LIBRARY_ID");
    shareClassesImpl("vn-classes", vnReportCardTemplateId, classLibraryId, false);
  }
}


/////////////////////////////////////////////////////////////////////////////////////////////////////
//
// 5 - saveStudentsPast
//
/////////////////////////////////////////////////////////////////////////////////////////////////////
function saveStudentsPast() {

  var fcomStr = "=query(students!1:999, \"select A,B,C,D,E,F,L,M,N,O,P,Q,R,T,U,V,W,X,Y,Z,AA,AB,AC,AD,AE where G=3 and AG>=" + getPassingPoint() + " order by C,E\")";
  var confStr = "=query(students!1:999, \"select A,B,C,D,E,F,L,M,N,O,P,Q,R,T,U,V,W,X,Y,Z,AA,AB,AC,AD,AE where G=8 and AG>=" + getPassingPoint() + " order by C,E\")";


  ////////////////////////////////////////////////////////////////////////////////////
  ////////////////////////////////////////////////////////////////////////////////////

  var studentsPastFolderId = getStr("STUDENTS_PAST_FOLDER_ID");

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var file = DriveApp.getFileById(ss.getId());
  var clsFolder = DriveApp.getFolderById(studentsPastFolderId);

  // Make a copy
  var newFile = file.makeCopy("new-file", clsFolder);

  // Open the new spreadsheet
  var ss = SpreadsheetApp.openById(newFile.getId());

  // Remove links to all class spreadsheets
  ss.getSheetByName("GL1A").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("GL1B").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("GL2A").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("GL2B").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("GL3A").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("GL3B").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("GL4A").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("GL4B").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("GL5A").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("GL5B").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("GL6A").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("GL6B").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("GL7A").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("GL7B").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("GL8A").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("GL8B").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("VN1A").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("VN1B").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("VN2A").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("VN2B").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("VN3A").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("VN3B").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("VN4A").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("VN4B").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("VN5A").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("VN5B").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("VN6A").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("VN6B").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("VN7A").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("VN7B").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("VN8A").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("VN8B").getRange("O2:O2").getCell(1, 1).setValue("");
  //ss.getSheetByName("gl-classes").getRange("G2:H17").clear();
  //ss.getSheetByName("vn-classes").getRange("G2:H17").clear();

  // Update First Communion sheet
  ss.getSheetByName("Eucharist").getRange("A1:A1").getCell(1, 1).setValue(fcomStr);

  // Update Confirmation sheet
  ss.getSheetByName("Confirmation").getRange("A1:A1").getCell(1, 1).setValue(confStr);

};



/////////////////////////////////////////////////////////////////////////////////////////////////////
//
// 6 - increaseGlGVnGForNewReg
//
/////////////////////////////////////////////////////////////////////////////////////////////////////
function increaseGlGVnGForNewReg() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var response = ui.alert(
      'Warning!!!',
      'Do you want to increase glG and vnG for new registration process?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {
    increaseGlGVnGForNewRegImpl();
  }
}

function waitSeconds(iMilliSeconds) {
    var counter= 0
        , start = new Date().getTime()
        , end = 0;
    while (counter < iMilliSeconds) {
        end = new Date().getTime();
        counter = end - start;
    }
}

function increaseGlGVnGForNewRegImpl() {
  var passing_point = getPassingPoint();

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Std_VGz6v3");
  var range = sheet.getRange(1, 1, 700, 38); //row, col, numRows, numCols
  var rowStartCell = sheet.getRange("AI1:AI1").getCell(1, 1); // <= Need to update column

  ////////////////////////////////////////////////////////////////////////////////////
  ////////////////////////////////////////////////////////////////////////////////////


  var id, isReg, glG, glN, vnG, vnN, glFinalPoint, vnFinalPoint;
  // iterate through all cells in the range
  for (var cellRow = rowStartCell.getValue(); ; cellRow++) {
    id = range.getCell(cellRow, idCol).getValue();
    if(id == "") break;

    isReg   = range.getCell(cellRow, isRegCol).getValue();
    glG     = range.getCell(cellRow, glGCol).getValue();
    glN     = range.getCell(cellRow, glNCol).getValue();
    vnG     = range.getCell(cellRow, vnGCol).getValue();
    vnN     = range.getCell(cellRow, vnNCol).getValue();
    glFinalPoint  = range.getCell(cellRow, glFinalPointCol).getValue();
    vnFinalPoint  = range.getCell(cellRow, vnFinalPointCol).getValue();
    if(isReg == "x") {
      if(glG > 0 && glN != "" && glFinalPoint >= passing_point) {
        range.getCell(cellRow, glGCol).setValue(glG + 1);
      }

      if(vnG > 0 && vnN != "" && vnFinalPoint >= passing_point) {
        range.getCell(cellRow, vnGCol).setValue(vnG + 1);
      }

      // Clear the x
      range.getCell(cellRow, isRegCol).setValue('');

      // Clear glFinalPoint and vnFinalPoint

      waitSeconds(1000);
    }

    // update rowStart cell
    rowStartCell.setValue(cellRow);
  }

};




/////////////////////////////////////////////////////////////////////////////////////////////////////
//
// 7 - Clear data in external classes
//
/////////////////////////////////////////////////////////////////////////////////////////////////////
function clearDataExternalClasses() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var response = ui.alert(
      'Warning!!!',
      'Do you want to clear data in external classes?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {
    clearDataExternalClassesImpl("gl-classes");
    clearDataExternalClassesImpl("vn-classes");
  }
}


function clearDataExternalClassesImpl(sheetName) {


  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var range = sheet.getRange(2, 1, 25, 15); //row, col, numRows, numCols

  var clsName, action, clsFolder;

  // iterate through all cells in the range
  for (var cellRow = 1; cellRow <= range.getHeight(); cellRow++) {
    clsName = range.getCell(cellRow, clsNameCol).getValue();
    if( clsName == "")
      break;

    var actionCell = range.getCell(cellRow, actionCol);
    action = actionCell.getValue();

    if(action == 'x') {
      var folderId = range.getCell(cellRow, clsFolderIdCol).getValue();
      clsFolder = DriveApp.getFolderById(folderId);

      // Open target class spreadsheet
      var classId = getExternalClassSpreadsheetId(clsName, clsFolder);
      var tss = SpreadsheetApp.openById(classId);

      var delRange = tss.getSheetByName("attendance-HK1").getRange(3,6,60,20); //row, col, numRows, numCols
      delRange.clearContent();
      delRange = tss.getSheetByName("attendance-HK2").getRange(3,6,60,20); //row, col, numRows, numCols
      delRange.clearContent();
      delRange = tss.getSheetByName("grades").getRange(3,7,60,20); //row, col, numRows, numCols
      delRange.clearContent();
      delRange = tss.getSheetByName("honor-roll").getRange(3,6,20,1); //row, col, numRows, numCols
      delRange.clearContent();
      tss.getSheetByName("honor-roll").getRange("F3:F3").getCell(1, 1).setValue("1");
      tss.getSheetByName("honor-roll").getRange("F4:F4").getCell(1, 1).setValue("2");
      tss.getSheetByName("honor-roll").getRange("F5:F5").getCell(1, 1).setValue("3");
      tss.getSheetByName("honor-roll").getRange("F6:F6").getCell(1, 1).setValue("4");

      actionCell.setValue('');
    }
  }
};




/////////////////////////////////////////////////////////////////////////////////////////////////////
//
// DO NOT DELETE THIS FUNCTION
// Update classes
// This function can be used for updating individual cell in each class sheet
//
/////////////////////////////////////////////////////////////////////////////////////////////////////
function updateSheetsInThisSpreadSheet() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var response = ui.alert(
      'Warning!!!',
      'Do you want to updateSheetsInThisSpreadSheet?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {
    updateSheetsInThisSpreadSheetImpl();
  }
}


function updateSheetsInThisSpreadSheetImpl() {
  updateSheetsInThisSpreadSheetImpl2("gl-classes");
  updateSheetsInThisSpreadSheetImpl2("vn-classes");
}


function updateSheetsInThisSpreadSheetImpl2(sheetName) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var range = sheet.getRange(2, 1, 20, 15); //row, col, numRows, numCols

  var clsName, action;

  // iterate through all cells in the range
  for (var cellRow = 1; cellRow <= range.getHeight(); cellRow++) {
    clsName = range.getCell(cellRow, clsNameCol).getValue();
    if( clsName == "")
      break;

    var actionCell = range.getCell(cellRow, actionCol);
    action = actionCell.getValue();

    if(action == 'x') {
      clsFolder = DriveApp.getFolderById(range.getCell(cellRow, clsFolderIdCol).getValue());
      updateSheetsInThisSpreadSheet_ClassSheet(ss.getSheetByName(clsName), clsName, clsFolder);
      actionCell.setValue('');
    }
  }
}


function updateSheetsInThisSpreadSheet_ClassSheet(sheet, clsName, clsFolder) {

  sheet.getRange("A1:A1").getCell(1, 1).setValue(clsName);

  var newValue
  if (clsName.substr(0,2) == "GL") {
    newValue = "=query(studentsclass!1:902, \"select A,B,C,D,E,K,L,N,O,V,U,S,Q where \" & if(left(A1,1)=\"G\",\"G\",\"I\") & \"=\" & mid(A1,3,1) & \" and \" & if(left(A1,1)=\"G\",\"H\",\"J\") & \"='\" & right(A1,1) & \"' order by C,E\")";
  }
  else { // VN class sheet
    newValue = "=query(studentsclass!1:902, \"select A,B,C,D,E,K,L,N,O,V,T,S,Q where \" & if(left(A1,1)=\"G\",\"G\",\"I\") & \"=\" & mid(A1,3,1) & \" and \" & if(left(A1,1)=\"G\",\"H\",\"J\") & \"='\" & right(A1,1) & \"' order by C,E\")";
  }
  sheet.getRange("B1:B1").getCell(1, 1).setValue(newValue);


  var classId = getExternalClassSpreadsheetId(clsName, clsFolder);
  newValue = "=IMPORTRANGE(\"" + classId + "\",\"Grades!F3:F80\")";
  sheet.getRange("O2:O2").getCell(1, 1).setValue(newValue);

  newValue = "=CONCATENATE(COUNTIFS(O2:O92, \"0\"),\" | \", MIN(O2:O92),\" - \", MAX(O2:O92), \" | \", COUNTIFS(O2:O92, Q1), \":\", COUNTIFS(O2:O92, Q2)-COUNTIFS(O2:O92, Q1), \":\", COUNTIFS(O2:O92, Q3)-COUNTIFS(O2:O92, Q2), \":\", COUNTIFS(O2:O92, Q4)-COUNTIFS(O2:O92, Q3), \":\", COUNTIFS(O2:O92, Q5))";
  sheet.getRange("P1:P1").getCell(1, 1).setValue(newValue);

  newValue = "=query(AdminGrades!B2:B6, \"select B\")";
  sheet.getRange("Q1:Q1").getCell(1, 1).setValue(newValue);
}






/////////////////////////////////////////////////////////////////////////////////////////////////////
// DO NOT DELETE THIS FUNCTION
// Clone classes
// This function is for setting up new GLVN database for the first time only
/////////////////////////////////////////////////////////////////////////////////////////////////////
function cloneClassesUsingGL1AorVN1A() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var response = ui.alert(
      'Warning!!!',
      'Do you want to create new classes?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {
    cloneClassesUsingGL1AorVN1AImpl();
  }
}


function cloneClassesUsingGL1AorVN1AImpl() {
  var glClassTemplateId = getStr("GL1A_SPREADSHEET_ID");
  var vnClassTemplateId = getStr("VN1A_SPREADSHEET_ID");

  cloneClassesUsingGL1AorVN1AImpl2("gl-classes", glClassTemplateId);
  cloneClassesUsingGL1AorVN1AImpl2("vn-classes", vnClassTemplateId);
}

function debugCloneClassesUsingGL1AorVN1AImpl2 () {
  cloneClassesUsingGL1AorVN1AImpl2("gl-classes", "1NT3efSzhBashDiKfOARRvm3Jy24qfYDm62xpFkZIhlI");
}

function cloneClassesUsingGL1AorVN1AImpl2(sheetName, templateId) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var range = sheet.getRange(2, 1, 20, 15); //row, col, numRows, numCols

  var clsName, action, clsFolder;

  // iterate through all cells in the range
  for (var cellRow = 1; cellRow <= range.getHeight(); cellRow++) {
    clsName = range.getCell(cellRow, clsNameCol).getValue();
    if( clsName == "")
      break;

    var actionCell = range.getCell(cellRow, actionCol);
    action = actionCell.getValue();

    if(action == 'x' && (clsName != "GL1A" || clsName != "VN1A")) {
      var folderId = range.getCell(cellRow, clsFolderIdCol).getValue();
      clsFolder = DriveApp.getFolderById(folderId);

            /////////////////////////////////////////////////////////////////////////////
      // Create GLxx-report-cards folder if not exist
      /////////////////////////////////////////////////////////////////////////////
      var reportCardsFolderId;
      var reportCardfolders = clsFolder.getFoldersByName(clsName + "-Report-Cards");
      if (reportCardfolders.hasNext()) {
        reportCardsFolderId = reportCardfolders.next().getId();
      }
      else {
        var reportFolder = clsFolder.createFolder(clsName + "-Report-Cards");
        reportCardsFolderId = reportFolder.getId();
      }

      /////////////////////////////////////////////////////////////////////////////
      // Rename GLxx to "bk"
      /////////////////////////////////////////////////////////////////////////////
      var files = clsFolder.getFilesByName(clsName);
      if (files.hasNext()) {
        var file = files.next();
        if (file) {
          file.setName("bk");
        }
      }

      /////////////////////////////////////////////////////////////////////////////
      // Make a copy and save it into the class folder
      /////////////////////////////////////////////////////////////////////////////
      var file = DriveApp.getFileById(templateId);
      var newFile = file.makeCopy(clsName, clsFolder);

      /////////////////////////////////////////////////////////////////////////////
      // Open the new spreadsheet and setup basic functions
      /////////////////////////////////////////////////////////////////////////////
      var newss = SpreadsheetApp.openById(newFile.getId());
      newss.getSheetByName("contacts").getRange("A1:A1").getCell(1, 1).setValue(clsName);
      newss.getSheetByName("admin").getRange("B3:B3").getCell(1, 1).setValue(reportCardsFolderId);

      /////////////////////////////////////////////////////////////////////////////
      // Save new class spreadsheet id into the class worksheet (ex: GL1A) sheet in the master book
      /////////////////////////////////////////////////////////////////////////////
      var tstr = "=IMPORTRANGE(\"" + newFile.getId() + "\",\"Grades!F3:F80\")";
      ss.getSheetByName(clsName).getRange("O2:O2").getCell(1, 1).setValue(tstr);

      /////////////////////////////////////////////////////////////////////////////
      // Save new class spreadsheet id into the honor-gl-import or honor-vn-import
      // sheets in the students-extra book
      /////////////////////////////////////////////////////////////////////////////
      var imptStr = "=IMPORTRANGE(\"" + newFile.getId() + "\",\"honor-roll!B3:F` + (3+MAX_HONOR_ROLL-1) + `\")";
      var studentsExtraId = getStr("STUDENTS_EXTRA_SPREADSHEET_ID");
      var studentsExtraSs = SpreadsheetApp.openById(studentsExtraId);
      var hrSheet = studentsExtraSs.getSheetByName("honor-" + sheetName.slice(0, 2)+"-import"); // honor-gl-import or honor-vn-import sheet
      var hrRange = hrSheet.getRange(2, 1, 400, 15); //row, col, numRows, numCols
      var hrCell  = hrRange.getCell(((cellRow-1)*MAX_HONOR_ROLL)+1, 2);
      hrCell.setValue(imptStr);



      // Clear action x
      actionCell.setValue('');
    }
  }
}



/////////////////////////////////////////////////////////////////////////////////////////////////////
//
// DO NOT DELETE THIS FUNCTION
// Update classes
// This function can be used for updating individual cell in each class sheet
//
/////////////////////////////////////////////////////////////////////////////////////////////////////
function updateExternalClassSpreadSheets() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var response = ui.alert(
      'Warning!!!',
      'Do you want to updateExternalClassSpreadSheets?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {
    updateExternalClassSpreadSheetsImpl();
  }
}


function updateExternalClassSpreadSheetsImpl() {
  updateExternalClassSpreadSheetsImpl2("gl-classes");
  updateExternalClassSpreadSheetsImpl2("vn-classes");
}


function updateExternalClassSpreadSheetsImpl2(sheetName) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var range = sheet.getRange(2, 1, 20, 15); //row, col, numRows, numCols

  var clsName, action, clsFolder;

  // iterate through all cells in the range
  for (var cellRow = 1; cellRow <= range.getHeight(); cellRow++) {
    clsName = range.getCell(cellRow, clsNameCol).getValue();
    if( clsName == "")
      break;

    var actionCell = range.getCell(cellRow, actionCol);
    action = actionCell.getValue();

    if(action == 'x') {
      var folderId = range.getCell(cellRow, clsFolderIdCol).getValue();
      clsFolder = DriveApp.getFolderById(folderId);

      // Open target class spreadsheet
      var classId = getExternalClassSpreadsheetId(clsName, clsFolder);
      var tss = SpreadsheetApp.openById(classId);

      // // Update `contacts` sheet in each class spreadsheet
      // updateExternalClassSpreadSheets_contactsSheet(tss, ss.getId(), clsName);

      // // Update `attendance_HK1` sheet in each class spreadsheet
      // updateExternalClassSpreadSheets_attendanceSheet(tss, ss.getId(), "HK1");

      // // Update `attendance_HK2` sheet in each class spreadsheet
      // updateExternalClassSpreadSheets_attendanceSheet(tss, ss.getId(), "HK2");

      // Update `grades` sheet in each class spreadsheet
      updateExternalClassSpreadSheets_gradesSheet(tss, ss.getId(), clsName);

      // updateExternalClassSpreadSheets_honorSheet(tss);
      // updateExternalClassSpreadSheets_reviewSheet(tss);
      // updateExternalClassSpreadSheets_adminGradesSheet(tss, ss.getId());

      actionCell.setValue('');
    }
  }
}

function updateExternalClassSpreadSheets_contactsSheet(tss, studentMasterSpreadsheetId, clsName) {

  var sn = 'contacts';

  // sheet
  var sheet = tss.getSheetByName(sn);

  var newValue = "=IMPORTRANGE(\"" + studentMasterSpreadsheetId + "\",A1&\"!B1:N92\")";
  sheet.getRange("A2:A2").getCell(1, 1).setValue(newValue);

  // teacher names
  var newValue = "=IMPORTRANGE(\"" + studentMasterSpreadsheetId + "\",\"" + clsName.slice(0,2).toLowerCase() + "-classes!D\"&(2*mid(A1,3,1)+if(right(A1,1)=\"A\",0,1)))";
  sheet.getRange("B1:B1").getCell(1, 1).setValue(newValue);
}

function updateExternalClassSpreadSheets_attendanceSheet(tss, studentMasterSpreadsheetId, hocKy) {

  var sn = 'attendance-' + hocKy;

  // sheet
  var sheet = tss.getSheetByName(sn);

  var newValue = "=query(contacts!1:1, \"select A,B\")";
  sheet.getRange("A1:A1").getCell(1, 1).setValue(newValue);

  newValue = "=query(contacts!2:92, \"select C,E,G,I\")";
  sheet.getRange("B2:B2").getCell(1, 1).setValue(newValue);

  if (hocKy=="HK1") {
    newValue = "=IMPORTRANGE(\"" + studentMasterSpreadsheetId + "\",\"calendar!B1:S1\")";
  }
  else {
    newValue = "=IMPORTRANGE(\"" + studentMasterSpreadsheetId + "\",\"calendar!B2:S2\")";
  }
  sheet.getRange("F2:F2").getCell(1, 1).setValue(newValue);
}

function updateExternalClassSpreadSheets_gradesSheet(tss, studentMasterSpreadsheetId, clsName) {

  // sheet
  var sheet = tss.getSheetByName('grades');

  var newValue;
  // newValue = "=query(contacts!1:1, \"select A,B\")";
  // sheet.getRange("A1:A1").getCell(1, 1).setValue(newValue);

  // newValue = "=query(contacts!2:90, \"select A,B,C,E,J\")";
  // sheet.getRange("A2:A2").getCell(1, 1).setValue(newValue);

  // // teacher signature
  // var newValue = "=IMPORTRANGE(\"" + studentMasterSpreadsheetId + "\",\"" + clsName.slice(0,2).toLowerCase() + "-classes!E\"&(2*mid(A1,3,1)+if(right(A1,1)=\"A\",0,1)))";
  // sheet.getRange("F1:F1").getCell(1, 1).setValue(newValue);

  // newValue = "=if(A3<>\"\",if(G3<>\"d\",ROUND(sum(H3:L3) + sum(S3:W3),2),\"\"),\"\")";
  // sheet.getRange("F3:F3").getCell(1, 1).setValue(newValue);
  // // IMPORTANT: need to copy to the rest of the cell Fs

  newValue = "generate=> x\ngen&email=> e";
  sheet.getRange("G1:G1").getCell(1, 1).setValue(newValue);

  // sheet.getRange("H1:H1").getCell(1, 1).setValue("0-5");
  // sheet.getRange("H2:H2").getCell(1, 1).setValue("Part1");
  // sheet.getRange("I1:I1").getCell(1, 1).setValue("0-5");
  // sheet.getRange("I2:I2").getCell(1, 1).setValue("HWrk1");
  // sheet.getRange("J1:J1").getCell(1, 1).setValue("0-15");
  // sheet.getRange("J2:J2").getCell(1, 1).setValue("Quiz1");
  // sheet.getRange("K1:K1").getCell(1, 1).setValue("0-25");
  // sheet.getRange("K2:K2").getCell(1, 1).setValue("Exam1");
  // sheet.getRange("L1:L1").getCell(1, 1).setValue("blank-20");
  // sheet.getRange("L2:L2").getCell(1, 1).setValue("Extr1");
  // sheet.getRange("M1:M1").getCell(1, 1).setValue("0-20");
  // sheet.getRange("O1:O1").getCell(1, 1).setValue("0-5");
  // sheet.getRange("O2:O2").getCell(1, 1).setValue("Part2");
  // sheet.getRange("P1:P1").getCell(1, 1).setValue("0-5");
  // sheet.getRange("P2:P2").getCell(1, 1).setValue("HWrk2");
  // sheet.getRange("Q1:Q1").getCell(1, 1).setValue("0-15");
  // sheet.getRange("Q2:Q2").getCell(1, 1).setValue("Quiz2");
  // sheet.getRange("R1:R1").getCell(1, 1).setValue("0-25");
  // sheet.getRange("R2:R2").getCell(1, 1).setValue("Exam2");
  // sheet.getRange("S1:S1").getCell(1, 1).setValue("blank-20");
  // sheet.getRange("S2:S2").getCell(1, 1).setValue("Extr2");
  // sheet.getRange("T1:T1").getCell(1, 1).setValue("0-20");
}


function updateExternalClassSpreadSheets_honorSheet(tss) {

  // sheet
  var sheet = tss.getSheetByName('honor-roll');

  var newValue = "=query(contacts!1:1, \"select A,B\")";
  sheet.getRange("A1:A1").getCell(1, 1).setValue(newValue);

  newValue = "=query(grades!2:80, \"select A,B,C,D,F where F>0 order by F desc\")";
  sheet.getRange("A2:A2").getCell(1, 1).setValue(newValue);
}


function updateExternalClassSpreadSheets_reviewSheet(tss) {

  // sheet
  var sheet = tss.getSheetByName('comment-review');

  var newValue = "=query(contacts!1:1, \"select A,B\")";
  sheet.getRange("A1:A1").getCell(1, 1).setValue(newValue);

  newValue = "=query(grades!2:80, \"select A,C,D,F,N,U where F>0 order by F desc\")";
  sheet.getRange("A2:A2").getCell(1, 1).setValue(newValue);
}

function updateExternalClassSpreadSheets_adminGradesSheet(tss, studentMasterSpreadsheetId) {

  // sheet
  var sheet = tss.getSheetByName('adminGrades');

  var newValue = "=IMPORTRANGE(\"" + studentMasterSpreadsheetId + "\",\"adminGrades!1:10\")";
  sheet.getRange("A1:A1").getCell(1, 1).setValue(newValue);
}
