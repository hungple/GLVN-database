var RELEASE = "20230913"

// Std_VGz6v3
var idCol               = 1;  
var glGCol              = 7; 
var glNCol              = 8; 
var vnGCol              = 9; 
var vnNCol              = 10; 
var isRegCol            = 11;
var birthDateCol        = 21;

var euchDateCol         = 26; // Z
var euchLocationCol     = 27; // AA
var confDateCol         = 28; // AB
var confLocationCol     = 29; // AC

var glFinalPointCol     = 33; // AG 
var vnFinalPointCol     = 34; // AH


// gl-classes/vn-classes
var classes_clsNameCol          = 1;
var classes_gmailCol            = 6;
var classes_actionCol           = 7;
var classes_clsFolderIdCol      = 9;


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
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('GLVN')
      .addItem('Release: ' + RELEASE, 'showRelease')
      .addItem('Share classes (gl-classes/vn-classes)', 'shareClasses')
      .addItem('Unshare classes (gl-classes/vn-classes)', 'unShareClasses')
      .addSeparator()
      .addItem('1 - Save student final points (Std)', 'saveFinalPoints')
      .addItem('2 - Save First Communion date and location (Std)', 'saveEucharistInfo')
      .addItem('3 - Save Confirmation date and location (Std)', 'saveConfirmationInfo')
      .addItem('4 - Save students into students-past folder (should use root)', 'saveStudentsPast')
      .addItem('5 - Increase glG and vnG for new registration (Std)', 'increaseGlGVnGForNewReg')
      .addItem('6 - Clear old data in external classes (gl-classes/vn-classes : root)', 'clearDataExternalClasses')
      .addSeparator()
      .addItem('Delete old students (Std : root only)', 'deleteOldStudents')
      .addItem('Update class sheets in this student-master (gl-classes/vn-classes : root only)', 'updateSheetsInThisSpreadSheet')
      .addItem('Clone classes using GL1A or VN1A (gl-classes/vn-classes : root only', 'cloneClassesUsingGL1AorVN1A')
      .addItem('Update external class spreadSheets (gl-classes/vn-classes : root only)', 'updateExternalClassSpreadSheets')
      .addToUi();
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
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("admin");
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
// Share classes to the teachers & Un-share GL classes to the teachers
//
/////////////////////////////////////////////////////////////////////////////////////////////////////
function shareClasses() {
  var ui = SpreadsheetApp.getUi();
  
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
    shareClassesImpl("vn-classes", vnReportCardTemplateId, classLibraryId, true);
  }
}

function unShareClasses() {
  var ui = SpreadsheetApp.getUi();
  
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
    clsName = range.getCell(cellRow, classes_clsNameCol).getValue();
    gmails = range.getCell(cellRow, classes_gmailCol).getValue().trim();
    
    if( clsName == "")
      break;

    var actionCell = range.getCell(cellRow, classes_actionCol);
    action = actionCell.getValue().trim().toLowerCase();
    
    if(action == "x" && gmails != "") {

      clsFolder = DriveApp.getFolderById(range.getCell(cellRow, classes_clsFolderIdCol).getValue());

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
          //Only remove this gmail but don't remove other viewers or admins (admin also a teacher)
          if(isNotAdmin(admins, gmail)) {
            doc.removeViewer(gmail);
            libSpreadSheet.removeViewer(gmail);

            if(isShared == true){
              doc.addViewer(gmail);
              libSpreadSheet.addEditor(gmail);
            }
          }
        }
        catch(e) {
        } //ignore error
      }

      // Share the whole class folder
      try {
        // remove current editors except admins
        var editors = clsFolder.getEditors();
        for (var j = 0; j < editors.length; j++) {
          if(isNotAdmin(admins, editors[j].getEmail())){
            clsFolder.removeEditor(editors[j].getEmail());           
          }
        }
        
        // add new editors
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

  var rowStart = rowStartCell.getValue();
  
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
      else {
        range.getCell(cellRow, glFinalPointCol).setValue('');
      }
      
      if(vnName != "") {
        range.getCell(cellRow, vnFinalPointCol).setValue(getFinalGrade("VN" + vnLevel + vnName, id));
      }
      else {
        range.getCell(cellRow, vnFinalPointCol).setValue('');
      }
    }
    else {
      range.getCell(cellRow, glFinalPointCol).setValue('');
      range.getCell(cellRow, vnFinalPointCol).setValue('');
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
// 1 - Save Final Points
// 2 - Save eucharist date and location
// 3 - Save confirmation date and location
// 5 - increaseGlGVnGForNewReg
//
/////////////////////////////////////////////////////////////////////////////////////////////////////
function saveFinalPoints() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  
  var response = ui.alert(
      'Warning!!!',
      'Do you want to update final points for students?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {
    saveStdInfoImpl('saveFP');
  }
}

function saveEucharistInfo() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  
  var response = ui.alert(
      'Warning!!!',
      'Do you want to update first communion information for students?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {
    saveStdInfoImpl('eucharist');
  }
}

function saveConfirmationInfo() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  
  var response = ui.alert(
      'Warning!!!',
      'Do you want to save confirmation information for students?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {
    saveStdInfoImpl('confirmation');
  }
}

function increaseGlGVnGForNewReg() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  
  var response = ui.alert(
      'Warning!!!',
      'Do you want to increase glG and vnG for new registration process?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {
    saveStdInfoImpl('newReg');
  }
}


function saveStdInfoImpl(task) {
  
  //////////////////////// Get data from the Calendar sheet  ///////////////////////// 
  var calSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calendar");
  var calRange = calSheet.getRange(1, 1, 10, 11); //row, col, numRows, numCols 
  var commDate = calRange.getCell(3, 2).getValue(); //row, col
  var commLocation = getStr("CHURCH_INFO");
  var confDate = calRange.getCell(4, 2).getValue(); //row, col
  var confLocation = getStr("CHURCH_INFO");
  var passing_point = getPassingPoint();

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Std_VGz6v3");
  var range = sheet.getRange(1, 1, 700, 34); //row, col, numRows, numCols 
  var rowStartCell = sheet.getRange("AI1:AI1").getCell(1, 1); // <= Need to update column

  ////////////////////////////////////////////////////////////////////////////////////
  ////////////////////////////////////////////////////////////////////////////////////

  var rowStart = rowStartCell.getValue();
  
  var id, isReg, glG, glN, vnG, vnN, glFinalPoint, vnFinalPoint; 
  // iterate through all cells in the range
  for (var cellRow = rowStart; ; cellRow++) {
    id = range.getCell(cellRow, idCol).getValue(); 

    if(id == "") break;
    
    isReg  = range.getCell(cellRow, isRegCol).getValue();
    glG    = range.getCell(cellRow, glGCol).getValue();
    glN    = range.getCell(cellRow, glNCol).getValue();
    vnG    = range.getCell(cellRow, vnGCol).getValue();
    vnN    = range.getCell(cellRow, vnNCol).getValue();
    glFinalPoint = range.getCell(cellRow, glFinalPointCol).getValue();
    vnFinalPoint = range.getCell(cellRow, vnFinalPointCol).getValue();

    if(task == 'saveFP') {
      if(isReg == "x") {
        if(glN != "") {
          range.getCell(cellRow, glFinalPointCol).setValue(getFinalGrade("GL" + glG + glN, id));
        }
        else {
          range.getCell(cellRow, glFinalPointCol).setValue('');
        }
        
        if(vnN != "") {
          range.getCell(cellRow, vnFinalPointCol).setValue(getFinalGrade("VN" + vnG + vnN, id));
        }
        else {
          range.getCell(cellRow, vnFinalPointCol).setValue('');
        }
      }
      else {
        range.getCell(cellRow, glFinalPointCol).setValue('');
        range.getCell(cellRow, vnFinalPointCol).setValue('');
      }
    }
    else if(task == 'eucharist') {
      if(isReg == "x" && glG == 3 && glN != "" && glFinalPoint >= passing_point) {
        range.getCell(cellRow, euchDateCol).setValue(commDate);
        range.getCell(cellRow, euchLocationCol).setValue(commLocation);
      }
    }
    else if(task == 'confirmation') {
      if(isReg == "x" && glG == 8 && glN != "" && glFinalPoint >= passing_point) {
        range.getCell(cellRow, confDateCol).setValue(confDate);
        range.getCell(cellRow, confLocationCol).setValue(confLocation);
      }
    }
    else if(task == 'newReg') {
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
        
        waitSeconds(500);
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

function waitSeconds(iMilliSeconds) {
    var counter= 0
        , start = new Date().getTime()
        , end = 0;
    while (counter < iMilliSeconds) {
        end = new Date().getTime();
        counter = end - start;
    }
}

/////////////////////////////////////////////////////////////////////////////////////////////////////
//
// 4 - saveStudentsPast
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
// 6 - Clear data in external classes
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
  var range = sheet.getRange(2, 1, 20, 15); //row, col, numRows, numCols

  var clsName, action, clsFolder;
  
  // iterate through all cells in the range
  for (var cellRow = 1; cellRow <= range.getHeight(); cellRow++) {
    clsName = range.getCell(cellRow, classes_clsNameCol).getValue();
    if( clsName == "")
      break;

    var actionCell = range.getCell(cellRow, classes_actionCol);
    action = actionCell.getValue().trim().toLowerCase();
    
    if(action == 'x') {
      var folderId = range.getCell(cellRow, classes_clsFolderIdCol).getValue();
      clsFolder = DriveApp.getFolderById(folderId);
      
      // Clear data in the class spreadsheet
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
      tss.getSheetByName("honor-roll").getRange("F7:F7").getCell(1, 1).setValue("4");

      // Clear pdf file in report-cards folder
      var reportCardFolderId = getReportCardsFolderId(clsName, clsFolder);
      //emptyFolder(reportCardFolderId); can't delete files that were created by others

      actionCell.setValue('');
    }
  }
};

function emptyFolder(folderId) { // does not work

    const folder = DriveApp.getFolderById(folderId);

    while (folder.getFiles().hasNext()) {
      const file = folder.getFiles().next();
      file.setTrashed(true);
      // Drive.Files.remove(file.getId())
  }

}






/////////////////////////////////////////////////////////////////////////////////////////////////////
//
// DO NOT DELETE THIS FUNCTION
// Update classes
// This function can be used for updating individual cell in each class sheet
//
/////////////////////////////////////////////////////////////////////////////////////////////////////


/////////////////////////////////////////////////////////////////////////////////////////////////////
//
// deleteOldStudents
//
/////////////////////////////////////////////////////////////////////////////////////////////////////
function deleteOldStudents() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  
  var response = ui.alert(
      'Warning!!!',
      'Do you want to delete old students?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {
    deleteOldStudentsImpl();
  }
}

function calculateAge(dob) { 
    var diff_ms = Date.now() - dob.getTime();
    var age_dt = new Date(diff_ms); 
  
    return Math.abs(age_dt.getUTCFullYear() - 1970);
}

function deleteOldStudentsImpl() {
 
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Std_VGz6v3");
  var range = sheet.getRange(1, 1, 700, 38); //row, col, numRows, numCols
  var rowStartCell = sheet.getRange("AI1:AI1").getCell(1, 1); // <= Need to update column
  //var lastRow = SpreadsheetApp.getActiveSheet().getLastRow();

  ////////////////////////////////////////////////////////////////////////////////////
  ////////////////////////////////////////////////////////////////////////////////////
  
  
  var isReg, glG, glN, vnG, vnN, birthDate; 
  for (var cellRow = rowStartCell.getValue(); cellRow > 1; cellRow--) {

    isReg   = range.getCell(cellRow, isRegCol).getValue();
    glG     = range.getCell(cellRow, glGCol).getValue();
    glN     = range.getCell(cellRow, glNCol).getValue();
    vnG     = range.getCell(cellRow, vnGCol).getValue();
    vnN     = range.getCell(cellRow, vnNCol).getValue();
    birthDate = range.getCell(cellRow, birthDateCol).getValue();

    if(isReg == '') {
      if(   (glG > 8 && vnG > 8)
         || (glG > 8 && vnN == '')
         || (glN == '' && vnG > 8)
         || calculateAge(birthDate)>17) {
        sheet.deleteRow(cellRow);
      }

      waitSeconds(500);
    }
   
    // update rowStart cell
    rowStartCell.setValue(cellRow);
  }
  
};



function updateSheetsInThisSpreadSheet() {
  var ui = SpreadsheetApp.getUi();
  
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
    clsName = range.getCell(cellRow, classes_clsNameCol).getValue();
    if( clsName == "")
      break;

    var actionCell = range.getCell(cellRow, classes_actionCol);
    action = actionCell.getValue();
   
    if(action == 'x') {
      clsFolder = DriveApp.getFolderById(range.getCell(cellRow, classes_clsFolderIdCol).getValue());
      updateSheetsInThisSpreadSheet_ClassSheet(ss.getSheetByName(clsName), clsName, clsFolder);
      actionCell.setValue('');
    }
  }
}


function updateSheetsInThisSpreadSheet_ClassSheet(sheet, clsName, clsFolder) {

  sheet.getRange("A1:A1").getCell(1, 1).setValue(clsName);

  var newValue;
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
  var ui = SpreadsheetApp.getUi();
  
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
    clsName = range.getCell(cellRow, classes_clsNameCol).getValue();
    if( clsName == "")
      break;

    var classFile;
    var classSs;

    var actionCell = range.getCell(cellRow, classes_actionCol);
    action = actionCell.getValue();
    
    if(action == 'x') {
      var folderId = range.getCell(cellRow, classes_clsFolderIdCol).getValue();
      clsFolder = DriveApp.getFolderById(folderId);

      if(clsName != "GL1A" && clsName != "VN1A") { // don't replace GL1A or VN1A

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
        classFile = file.makeCopy(clsName, clsFolder);

        /////////////////////////////////////////////////////////////////////////////
        // Open the new spreadsheet and setup basic functions
        /////////////////////////////////////////////////////////////////////////////
        classSs = SpreadsheetApp.openById(classFile.getId());
        classSs.getSheetByName("contacts").getRange("A1:A1").getCell(1, 1).setValue(clsName);
        classSs.getSheetByName("admin").getRange("B3:B3").getCell(1, 1).setValue(reportCardsFolderId);
        classSs.getSheetByName("contacts").getRange("B1:B1").getCell(1, 1).setValue("=IMPORTRANGE(\"" 
            + ss.getId() + "\",\"" + clsName.slice(0,2).toLowerCase() + "-classes!D" 
            + (cellRow+1) + "\")");
        classSs.getSheetByName("grades").getRange("F1:F1").getCell(1, 1).setValue("=IMPORTRANGE(\"" 
            + ss.getId() + "\",\"" + clsName.slice(0,2).toLowerCase() + "-classes!E" 
            + (cellRow+1) + "\")");
      }

      if (! classFile) { // This must be GL1A or VN1A
        var files = clsFolder.getFilesByName(clsName);
        if (files.hasNext()) {
          classFile = files.next();
          classSs = SpreadsheetApp.openById(classFile.getId());
        }
      }


      if (classSs) {
        /////////////////////////////////////////////////////////////////////////////
        // For testing
        /////////////////////////////////////////////////////////////////////////////
        classSs.getSheetByName("grades").getRange("K3:K3").getCell(1, 1).setValue(genTestPoint(clsName, .01));
        classSs.getSheetByName("grades").getRange("K4:K4").getCell(1, 1).setValue(genTestPoint(clsName, .02));
        classSs.getSheetByName("grades").getRange("K5:K5").getCell(1, 1).setValue(genTestPoint(clsName, .03));
        classSs.getSheetByName("grades").getRange("K6:K6").getCell(1, 1).setValue(genTestPoint(clsName, .04));
        classSs.getSheetByName("grades").getRange("S3:S3").getCell(1, 1).setValue(genTestPoint(clsName, .01));
        classSs.getSheetByName("grades").getRange("S4:S4").getCell(1, 1).setValue(genTestPoint(clsName, .02));
        classSs.getSheetByName("grades").getRange("S5:S5").getCell(1, 1).setValue(genTestPoint(clsName, .03));
        classSs.getSheetByName("grades").getRange("S6:S6").getCell(1, 1).setValue(genTestPoint(clsName, .04));
      }

      if (classFile) {
        /////////////////////////////////////////////////////////////////////////////
        // Update students-master spreadsheet
        // Save new class spreadsheet id into the class worksheet (ex: GL1A) sheet
        /////////////////////////////////////////////////////////////////////////////
        var tstr = "=IMPORTRANGE(\"" + classFile.getId() + "\",\"Grades!F3:F80\")";
        ss.getSheetByName(clsName).getRange("O2:O2").getCell(1, 1).setValue(tstr);
        
        /////////////////////////////////////////////////////////////////////////////
        // Update students-extra spreadsheet
        // Save new class spreadsheet id into the honor-gl-import or honor-vn-import sheet
        /////////////////////////////////////////////////////////////////////////////
        var imptStr = "=IMPORTRANGE(\"" + classFile.getId() + "\",\"honor-roll!B3:F" + (3+MAX_HONOR_ROLL-1) + "\")";
        var studentsExtraId = getStr("STUDENTS_EXTRA_SPREADSHEET_ID");
        var studentsExtraSs = SpreadsheetApp.openById(studentsExtraId);
        var hrSheet = studentsExtraSs.getSheetByName("honor-" + sheetName.slice(0, 2)+"-import"); // honor-gl-import or honor-vn-import sheet
        var hrRange = hrSheet.getRange(2, 1, 500, 15); //row, col, numRows, numCols
        var hrCell  = hrRange.getCell(((cellRow-1)*MAX_HONOR_ROLL)+1, 2);
        hrCell.setValue(imptStr);
      }


      // Clear action x
      actionCell.setValue('');
    }
  }
}

function genTestPoint(clsName, delta) {
  var point = parseInt(clsName.charAt(2), 10);
  if (clsName.charAt(0) == 'V') {
    point = point + 8;
  }

  if (clsName.charAt(3) == 'A') {
    point = point + 0.1;
  }
  else {
    point = point + 0.2;
  }
  point = point + delta;
  return point;
}

function testGenTestPoint() {
  var point = genTestPoint("VN2B", 0.01)
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
    clsName = range.getCell(cellRow, classes_clsNameCol).getValue();
    if( clsName == "")
      break;

    var actionCell = range.getCell(cellRow, classes_actionCol);
    action = actionCell.getValue();
    
    if(action == 'x') {
      var folderId = range.getCell(cellRow, classes_clsFolderIdCol).getValue();
      clsFolder = DriveApp.getFolderById(folderId);
      
      // Open target class spreadsheet
      var classId = getExternalClassSpreadsheetId(clsName, clsFolder);
      var tss = SpreadsheetApp.openById(classId);

      // Update `contacts` sheet in each class spreadsheet
      // updateExternalClassSpreadSheets_contactsSheet(tss, ss.getId(), clsName, cellRow);

      // // Update `attendance_HK1` sheet in each class spreadsheet
      // updateExternalClassSpreadSheets_attendanceSheet(tss, ss.getId(), "HK1");

      // // Update `attendance_HK2` sheet in each class spreadsheet
      // updateExternalClassSpreadSheets_attendanceSheet(tss, ss.getId(), "HK2");

      // Update `grades` sheet in each class spreadsheet
      updateExternalClassSpreadSheets_gradesSheet(tss, ss.getId(), clsName, cellRow);

      // updateExternalClassSpreadSheets_honorSheet(tss);
      // updateExternalClassSpreadSheets_reviewSheet(tss);
      // updateExternalClassSpreadSheets_adminGradesSheet(tss, ss.getId());

      // updateStudentsExtraSpreadSheet(sheetName, classId, cellRow);

      actionCell.setValue('');
    }
  }
}

function updateExternalClassSpreadSheets_contactsSheet(tss, studentMasterSpreadsheetId, clsName, cellRow) {

  var sn = 'contacts';
      
  // sheet
  var sheet = tss.getSheetByName(sn);
  
  //var newValue = "=IMPORTRANGE(\"" + studentMasterSpreadsheetId + "\",A1&\"!B1:N92\")";
  //sheet.getRange("A2:A2").getCell(1, 1).setValue(newValue);

  // teacher names
  var newValue = "=IMPORTRANGE(\"" + studentMasterSpreadsheetId + "\",\"" + clsName.slice(0,2).toLowerCase() + "-classes!D" + (cellRow+1) + "\")";
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

function updateExternalClassSpreadSheets_gradesSheet(tss, studentMasterSpreadsheetId, clsName, cellRow) {
     
  // sheet
  var sheet = tss.getSheetByName('grades');
  
  var newValue;
  // newValue = "=query(contacts!1:1, \"select A,B\")";
  // sheet.getRange("A1:A1").getCell(1, 1).setValue(newValue);

  // newValue = "=query(contacts!2:90, \"select A,B,C,E,J\")";
  // sheet.getRange("A2:A2").getCell(1, 1).setValue(newValue);

  // teacher signature
  // var newValue = "=IMPORTRANGE(\"" + studentMasterSpreadsheetId + "\",\"" + clsName.slice(0,2).toLowerCase() + "-classes!E" + (cellRow+1) + "\")";
  // sheet.getRange("F1:F1").getCell(1, 1).setValue(newValue);

  // newValue = "=if(A3<>\"\",if(G3<>\"d\",ROUND(sum(H3:L3) + sum(S3:W3),2),\"\"),\"\")";
  // sheet.getRange("F3:F3").getCell(1, 1).setValue(newValue);
  // // IMPORTANT: need to copy to the rest of the cell Fs

  // newValue = "generate=> x\ngen&email=> e";
  // sheet.getRange("G1:G1").getCell(1, 1).setValue(newValue);

  // sheet.getRange("H1:H1").getCell(1, 1).setValue("0-5");
  // sheet.getRange("H2:H2").getCell(1, 1).setValue("Part1");
  // sheet.getRange("I1:I1").getCell(1, 1).setValue("0-5");
  // sheet.getRange("I2:I2").getCell(1, 1).setValue("HWrk1");
  // sheet.getRange("J1:J1").getCell(1, 1).setValue("0-15");
  // sheet.getRange("J2:J2").getCell(1, 1).setValue("Quiz1");
  // sheet.getRange("K1:K1").getCell(1, 1).setValue("0-25");
  // sheet.getRange("K2:K2").getCell(1, 1).setValue("Exam1");
  // sheet.getRange("L1:L1").getCell(1, 1).setValue("blank-20");
  sheet.getRange("L2:L2").getCell(1, 1).setValue("Extra1");
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
  sheet.getRange("S2:S2").getCell(1, 1).setValue("Extra2");
  // sheet.getRange("T1:T1").getCell(1, 1).setValue("0-20");

  // sheet.getRange("K3:K3").getCell(1, 1).setValue(genTestPoint(clsName, .01));
  // sheet.getRange("K4:K4").getCell(1, 1).setValue(genTestPoint(clsName, .02));
  // sheet.getRange("K5:K5").getCell(1, 1).setValue(genTestPoint(clsName, .03));
  // sheet.getRange("K6:K6").getCell(1, 1).setValue(genTestPoint(clsName, .04));

}


function updateExternalClassSpreadSheets_honorSheet(tss) {
     
  // sheet
  var sheet = tss.getSheetByName('honor-roll');
  
  // sheet.getRange("A1:A1").getCell(1, 1).setValue(
  //   "=query(contacts!1:1, \"select A,B\")"
  // );

  // sheet.getRange("A2:A2").getCell(1, 1).setValue(
  //   "=query(grades!2:80, \"select A,B,C,D,F where F>0 order by F desc\")"
  // );

  // sheet.getRange("F3:F3").getCell(1, 1).setValue('1');
  sheet.getRange("F4:F4").getCell(1, 1).setValue('2');
  // sheet.getRange("F5:F5").getCell(1, 1).setValue('3');
  // sheet.getRange("F6:F6").getCell(1, 1).setValue('4');
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

function updateStudentsExtraSpreadSheet(sheetName, classId, cellRow) {
  /////////////////////////////////////////////////////////////////////////////
  // Save new class spreadsheet id into the honor-gl-import or honor-vn-import
  // sheets in the students-extra book
  /////////////////////////////////////////////////////////////////////////////
  var imptStr = "=IMPORTRANGE(\"" + classId + "\",\"honor-roll!B3:F" + (3+MAX_HONOR_ROLL-1) + "\")";
  var studentsExtraId = getStr("STUDENTS_EXTRA_SPREADSHEET_ID");
  var studentsExtraSs = SpreadsheetApp.openById(studentsExtraId);
  var hrSheet = studentsExtraSs.getSheetByName("honor-" + sheetName.slice(0, 2)+"-import"); // honor-gl-import or honor-vn-import sheet
  var hrRange = hrSheet.getRange(2, 1, 400, 15); //row, col, numRows, numCols
  var hrCell  = hrRange.getCell(((cellRow-1)*MAX_HONOR_ROLL)+1, 2);
  hrCell.setValue(imptStr);
}
