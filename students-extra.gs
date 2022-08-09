var RELEASE = "202200809"

// students-registration
var idCol       = 1;
var snCol       = 2;
var fnCol       = 3;
var mnCol       = 4;
var lnCol       = 5;
var geCol       = 6;
var glnCol      = 7;
var vnnCol      = 8;
var tssCol      = 9;
var addrCol     = 10;
var fanCol      = 11;
var facCol      = 12;
var faeCol      = 13;
var monCol      = 14;
var mocCol      = 15;
var moeCol      = 16;
var ennCol      = 17;
var bdCol       = 18;
var bpCol       = 19;
var badCol      = 20;
var balCol      = 21;
var eudCol      = 22;
var eulCol      = 23;
var codCol      = 24;
var glgCol      = 25;

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
    name : "Create registrate forms",
    functionName : "createRegForms"
  },
  //,{
  //  name : "Merge Google docs",
  //  functionName : "mergeGoogleDocs"
  //}
  //{
  //  name : "Email New School Year Letters",
  //  functionName : "emailNewSchoolYearLetters"
  //}
  {
    name : "Create eucharist certificates",
    functionName : "createEuchCertificates"
  },
  {
    name : "Create confirmation certificates",
    functionName : "createConfCertificates"
  },
  {
    name : "Create letters",
    functionName : "createLetters"
  },
  {
    name : "Email parents",
    functionName : "emailParents"
  },
  {
    name : "Create award certificates for GL",
    functionName : "createAwardCertificatesGL"
  },

  {
    name : "Create award certificates for VN",
    functionName : "createAwardCertificatesVN"
  },,
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

// function mergeGoogleDocs() {

//   var docIDs = ['1GEbqRiuR2TW1icH5BTDN1dteios6Q6dCmMqtamOLHw4','1GG_QvGVNTixlte9D77IxaCaFVuRT7wuA4hBS2gXvrh0','1vNIb3LccROQlVecNxC00ILYrgTszCDmSNULLQDsOlHU','12hv7KicN4G-Ln9ab2EKo6_Mk-P0wDle_uAp6zb5Q9js'];
//   var baseDoc = DocumentApp.openById(docIDs[0]);

//   var body = baseDoc.getActiveSection();

//   for (var i = 1; i < docIDs.length; ++i ) {
//     var otherBody = DocumentApp.openById(docIDs[i]).getActiveSection();
//     var totalElements = otherBody.getNumChildren();
//     for( var j = 0; j < totalElements; ++j ) {
//       var element = otherBody.getChild(j).copy();
//       var type = element.getType();
//       if( type == DocumentApp.ElementType.PARAGRAPH ) {
//         //body.appendPageBreak();
//         body.appendParagraph(element);
//       }
//       else if( type == DocumentApp.ElementType.TABLE ) {
//         //body.appendPageBreak();
//         body.appendTable(element);
//       }
//       else if( type == DocumentApp.ElementType.LIST_ITEM ) {
//         //body.appendPageBreak();
//         body.appendListItem(element);
//       }
//       else
//         throw new Error("Unknown element type: "+type);
//     }
//   }
// }

function getSchoolName() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("admin");
  return sheet.getRange("B2:B2").getCell(1, 1).getValue();
}

function getRegTemplateId() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("admin");
  return sheet.getRange("B3:B3").getCell(1, 1).getValue();
}


function getRegFolderId() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("admin");
  return sheet.getRange("B4:B4").getCell(1, 1).getValue();
}


// DOES NOT WORK
//function getAllRegDocId() {
//  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("admin");
//  return sheet.getRange("B4:B4").getCell(1, 1).getValue();
//}

/////////////////////////////////////////////////////////////////////////////////////////////////////
//
// Create/Email parents
//
/////////////////////////////////////////////////////////////////////////////////////////////////////
function emailRegForms() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var response = ui.alert(
      'Warning!!!',
      'Do you want to email the registration forms to the parents?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {
    emailRegFormsImpl();
  }
}


function createRegForms() {
  createRegFormsImpl();
}

function createRegFormsImpl() {

  ////////////////////////////////////////////////////////////////////////////////////


  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("registration-print");
  var range = sheet.getRange(1, 1, 700, 25); //row, col, numRows, numCols <= need to update numRows
  var rowStartCell = sheet.getRange("Z1:Z1").getCell(1, 1); // <= need to update column
  Logger.log("rowStartCell: " + rowStartCell);

  ////////////////////////////////////////////////////////////////////////////////////

  var schName     = getSchoolName();

  var docName, id, orderby, fn, ln;

  for (var cellRow = rowStartCell.getValue()+1; ; cellRow++) {
    id = range.getCell(cellRow, idCol).getValue();
    if(id == "") break;

    var gln, vnn;
    gln = range.getCell(cellRow, glnCol).getValue();
    vnn = range.getCell(cellRow, vnnCol).getValue();

    if ((gln != '' && gln.substring(0,1) == 8) && (vnn != '' && vnn.substring(0,1) == 8)) { continue; }

    if (schName == 'OLR') {
      orderby  = range.getCell(cellRow, lnCol).getValue().substring(0, 9) + '-' + range.getCell(cellRow, addrCol).getValue().substring(0, 9);
    }
    else { //MHT
      orderby  = range.getCell(cellRow, addrCol).getValue().substring(0, 9) + '-GL' + gln;
    }

    fn       = range.getCell(cellRow, fnCol).getValue();
    ln       = range.getCell(cellRow, lnCol).getValue();

    docName = orderby + '-' + fn + '-' + ln + '-' + id + "-regForm";

    var formId      = getRegTemplateId();
    var folerId     = getRegFolderId();

    // Get document template, copy it as a new temp doc, and save the Doc’s id
    var copyId = DriveApp.getFileById(formId).makeCopy(docName).getId();

    // Open the temporary document
    var copyDoc = DocumentApp.openById(copyId);

    // Get the document’s body section
    var copyBody = copyDoc.getActiveSection();

    // Replace place holder keys,in our google doc template
    copyBody.replaceText('@id@', id);
    copyBody.replaceText('@enn@', range.getCell(cellRow, ennCol).getValue());
    copyBody.replaceText('@sn@', range.getCell(cellRow, snCol).getValue());
    copyBody.replaceText('@fn@', fn);
    copyBody.replaceText('@mn@', range.getCell(cellRow, mnCol).getValue());
    copyBody.replaceText('@ln@', ln);
    if (vnn == "") {
      copyBody.replaceText('@cls@', "GL" + gln);
    }
    else {
      copyBody.replaceText('@cls@', "GL" + gln + " - VN" + vnn);
    }
    copyBody.replaceText('@se@', range.getCell(cellRow, geCol).getValue());
    if (range.getCell(cellRow, bdCol).getValue() != '') {
      copyBody.replaceText('@bd@', Utilities.formatDate(new Date(range.getCell(cellRow, bdCol).getValue()), "GMT", "MMM d, yyyy"));
    }
    else {
      copyBody.replaceText('@bd@', ' ');
    }
    copyBody.replaceText('@bp@', range.getCell(cellRow, bpCol).getValue());
    if (range.getCell(cellRow, badCol).getValue() != '') {
      copyBody.replaceText('@ba@', 'Y');
      Logger.log("bad: '" + range.getCell(cellRow, badCol).getValue() + "'");
      copyBody.replaceText('@bad@', Utilities.formatDate(new Date(range.getCell(cellRow, badCol).getValue()), "GMT", "MMM d, yyyy"));
      copyBody.replaceText('@bal@', range.getCell(cellRow, balCol).getValue());
    }
    else {
      copyBody.replaceText('@ba@', ' ');
      copyBody.replaceText('@bad@', ' ');
      copyBody.replaceText('@bal@', ' ');
    }

    Logger.log("EUD");
    if (range.getCell(cellRow, eudCol).getValue() != '') {
      copyBody.replaceText('@eu@', 'Y');
      Logger.log("bad: '" + range.getCell(cellRow, eudCol).getValue() + "'");
      copyBody.replaceText('@eud@', Utilities.formatDate(new Date(range.getCell(cellRow, eudCol).getValue()), "GMT", "MMM d, yyyy"));
      copyBody.replaceText('@eul@', range.getCell(cellRow, eulCol).getValue());
    }
    else {
      copyBody.replaceText('@eu@', ' ');
      copyBody.replaceText('@eud@', ' ');
      copyBody.replaceText('@eul@', ' ');
    }

    if (range.getCell(cellRow, codCol).getValue() != '') {
      copyBody.replaceText('@co@', 'Y');
      copyBody.replaceText('@cod@', Utilities.formatDate(new Date(range.getCell(cellRow, codCol).getValue()), "GMT", "MMM d, yyyy"));
    }
    else {
      copyBody.replaceText('@co@', ' ');
      copyBody.replaceText('@cod@', ' ');
    }

    copyBody.replaceText('@fan@', range.getCell(cellRow, fanCol).getValue());
    copyBody.replaceText('@mon@', range.getCell(cellRow, monCol).getValue());
    copyBody.replaceText('@fac@', range.getCell(cellRow, facCol).getValue());
    copyBody.replaceText('@moc@', range.getCell(cellRow, mocCol).getValue());
    copyBody.replaceText('@fae@', range.getCell(cellRow, faeCol).getValue());
    copyBody.replaceText('@moe@', range.getCell(cellRow, moeCol).getValue());
    copyBody.replaceText('@addr@', range.getCell(cellRow, addrCol).getValue());

    // DOES NOT WORK
    //Logger.log("hpl");
    //var allRegDocApp = DocumentApp.openById(allRegDocId);
    //Logger.log("hpl2");
    //var allRegDoc = allRegDocApp.getActiveSection();
    //Logger.log("hpl3");
    //allRegDoc.appendPageBreak();
    //Logger.log("hpl4");
    //allRegDoc.appendParagraph(copyBody);
    //Logger.log("hpl5");
    //allRegDoc.saveAndClose();
    //Logger.log("hpl6");


    // Save and close the temporary document
    copyDoc.saveAndClose();

    // Convert temporary document to PDF
    var pdf = DriveApp.getFileById(copyId).getAs("application/pdf");

    // Delete temp file
    DriveApp.getFileById(copyId).setTrashed(true);

    // Delete old file
    //var files = DriveApp.getFolderById(folerId).getFilesByName(docName + ".pdf");
    //while (files.hasNext()) {
      //var file = files.next();
      //if(file.getOwner().getEmail() == Session.getActiveUser()) {
        //file.setTrashed(true);
      //}
    //}

    // Save pdf
    DriveApp.getFolderById(folerId).createFile(pdf);

    // Update current index
    rowStartCell.setValue(cellRow);

    // Attach PDF and send the email
    //if(email != "") {
    //    var body = "Kính Gửi Quý Phụ Huynh học sinh " + fName + " " + lName + ","
    //     + "<br>Xin phụ huynh xem đơn ghi danh và thông báo đính kèm. Xin cám ơn.<br>Chương Trình GLVN Andre Dũng Lạc.";

        //Logger.log("Email address:" + email + "<");
        //email = "hle007@yahoo.com";
    //    MailApp.sendEmail(email, "Đơn Ghi Danh Niên Khóa 2018-2019", body, {htmlBody: body, attachments: pdf});
    //}

  }
}


/////////////////////////////////////////////////////////////////////////////////////////////////////
//
// Email new school year letters
//
/////////////////////////////////////////////////////////////////////////////////////////////////////
function emailNewSchoolYearLetters() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var response = ui.alert(
      'Warning!!!',
      'Do you want to email to the parents?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {
    emailNewSchoolYearLettersImpl();
  }
}



function emailNewSchoolYearLettersImpl() {

  var folerId      = "0B3pmm9KdF9FjZFNsN3VQS0w4N0E";
  var formId       = "14W4HjxDuQSGqJgPNNWSOX96WvKAlyKvcvre9WED_WXk";

  var idCol       = 1;
  var sNameCol    = 2;
  var fNameCol    = 3;
  var mNameCol    = 4;
  var lNameCol    = 5;

  var glLevelCol  = 8;
  var glNameCol   = 9;
  var vnLevelCol  = 10;
  var vnNameCol   = 11;
  var tsSizeCol   = 12;
  var faNameCol   = 13;
  var moNameCol   = 14;
  var addrCol     = 15;
  var phone1Col   = 16;
  var phone2Col   = 17;
  var emailCol    = 18;
  var serviceDateCol   = 19;

  var locCol      = 23;

  var startCol    = 25;

  var folder = DriveApp.getFolderById(folerId);


  ////////////////////////////////// Get index //////////////////////////////////////
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Main");
  var range = sheet.getRange(1, 1, 700, 27); //row, col, numRows, numCols
  var rowStartCell = range.getCell(2, startCol);
  var rowStart = rowStartCell.getValue()+1;

  var reportFolder = DriveApp.getFolderById(folerId);

  var id, glName, vnName, fName, lName, fullName, email, docName, address, sdate_tmp, sdate, loc_tmp, loc;

  // iterate through all cells in the range
  for (var cellRow = rowStart; ; cellRow++) {
    id = range.getCell(cellRow, idCol).getValue();

    if(id == "") break;

    glName      = "GL" + range.getCell(cellRow, glLevelCol).getValue() + range.getCell(cellRow, glNameCol).getValue();
    vnName      = "VN" + range.getCell(cellRow, vnLevelCol).getValue() + range.getCell(cellRow, vnNameCol).getValue();
    fName       = range.getCell(cellRow, fNameCol).getValue();
    lName       = range.getCell(cellRow, lNameCol).getValue();
    address     = range.getCell(cellRow, addrCol).getValue();
    email       = range.getCell(cellRow, emailCol).getValue();
    sdate       = Utilities.formatDate(range.getCell(cellRow, serviceDateCol).getValue(), "GMT", "MMM d, yyyy");
    loc_tmp     = range.getCell(cellRow, locCol).getValue();

    if(email == "") {

    docName = glName + '_' + fName + '_' + lName + ' ' + id + "_nsyLetter";
    fullName = fName + ' ' + lName;
    if(loc_tmp == 1) {
      loc = "Nhà Thờ St. Maria Goretti.";
    }
    else {
      loc = "Nguyện Đường Các Thánh Tử Đạo.";
    }

    //sdate = new Date(sdate_tmp);


    // Get document template, copy it as a new temp doc, and save the Doc’s id
    var copyId = DriveApp.getFileById(formId).makeCopy(docName).getId();

    // Open the temporary document
    var copyDoc = DocumentApp.openById(copyId);

    // Get the document’s body section
    var copyBody = copyDoc.getActiveSection();

    // Replace place holder keys,in our google doc template
    copyBody.replaceText('@sname@', fullName);
    copyBody.replaceText('@address@', address);
    copyBody.replaceText('@glname@', glName);
    copyBody.replaceText('@vnname@', vnName);
    copyBody.replaceText('@sdate@', sdate);
    copyBody.replaceText('@loc@', loc);

    // Save and close the temporary document
    copyDoc.saveAndClose();

    // Convert temporary document to PDF
    var pdf = DriveApp.getFileById(copyId).getAs("application/pdf");

    // Delete temp file
    DriveApp.getFileById(copyId).setTrashed(true);

    // Delete old file
    //var files = folder.getFilesByName(docName + ".pdf");
    //while (files.hasNext()) {
      //var file = files.next();
      //if(file.getOwner().getEmail() == Session.getActiveUser()) {
        //file.setTrashed(true);
      //}
    //}

    // Save pdf
    folder.createFile(pdf);

    } //if(email == "")

    // Update current index
    rowStartCell.setValue(cellRow);

    // Attach PDF and send the email
    if(email != "") {
        var body = "Kính Gửi Quý Phụ Huynh Học Sinh em " + fullName + ","
         + "<br>Xin quí Phụ Huynh xem thư thông báo đính kèm. Xin cám ơn.<br>Chương Trình GLVN Andre Dũng Lạc.";

        //Logger.log("Email address:" + email + "<");
        //email = "hle007@yahoo.com";
        //MailApp.sendEmail(email, "Thông Báo Phụ Huynh, Khai Giảng và họp niên khóa 2017-2018", body, {htmlBody: body, attachments: pdf});
    }

  }
}

function createEuchCertificates() {
  // last update: 4/23/2022

  ////////////////////////////////////////////////////////////////////////////////////
  var folerId     = "1Sn5I1a0I-j0WGMA9KROhkNQww5rojHCn";
  var formId      = "1W2h4WQq5KqOlDJAVWH1pljc7SwcYfcr3eLofhY9Xqvc";

  var idCol       = 1;
  var sNameCol    = 2;
  var fNameCol    = 3;
  var mNameCol    = 4;
  var lNameCol    = 5;

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("fcom-certificates");
  var range = sheet.getRange(1, 1, 130, 12); //row, col, numRows, numCols
  var rowStartCell = sheet.getRange("G1:G1").getCell(1, 1);
  ////////////////////////////////////////////////////////////////////////////////////

  var docName, id, sName, fName, mName, lName;

  // iterate through all cells in the range
  for (var cellRow = rowStartCell.getValue()+1; ; cellRow++) {

    id = range.getCell(cellRow, idCol).getValue();
    if(id == "") { break; }

    fName = range.getCell(cellRow, fNameCol).getValue();
    sName = range.getCell(cellRow, sNameCol).getValue();
    mName = range.getCell(cellRow, mNameCol).getValue();
    lName = range.getCell(cellRow, lNameCol).getValue();

    docName = fName + '-' + lName + '-' + id + "-certificate";

    // Get document template, copy it as a new temp doc, and save the Doc’s id
    var copyId = DriveApp.getFileById(formId).makeCopy(docName).getId();

    // Open the temporary document
    var copyDoc = DocumentApp.openById(copyId);

    // Get the document’s body section
    var copyBody = copyDoc.getActiveSection();

    // Replace place holder keys,in our google doc template
    copyBody.replaceText('@SNa@', sName);
    copyBody.replaceText('@FNa@', fName);
    copyBody.replaceText('@MNa@', mName);
    copyBody.replaceText('@LNa@', lName);

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

    // Update current index
    rowStartCell.setValue(cellRow);
  }
}


function createConfCertificates() {
  // last update: 4/23/2022

  ////////////////////////////////////////////////////////////////////////////////////
  var folerId     = "153ucgAbXdJjuWgA7_60T3HktlEf69HLw";
  var formId      = "1_i2AyOOX62KcB16c97BQZLt9B5KMZ6i2XmsvSvevJdc";

  var idCol       = 1;
  var sNameCol    = 2;
  var fNameCol    = 3;
  var mNameCol    = 4;
  var lNameCol    = 5;

  var classesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("conf-certificates");
  var range = classesSheet.getRange(1, 1, 100, 12); //row, col, numRows, numCols
  var rowStartCell = sheet.getRange("G1:G1").getCell(1, 1);
  ////////////////////////////////////////////////////////////////////////////////////

  var docName, id, sName, fName, mName, lName;

  // iterate through all cells in the range
  for (var cellRow = rowStartCell.getValue()+1; ; cellRow++) {

    id = range.getCell(cellRow, idCol).getValue();
    if(id == "") { break; }

    fName = range.getCell(cellRow, fNameCol).getValue();
    sName = range.getCell(cellRow, sNameCol).getValue();
    mName = range.getCell(cellRow, mNameCol).getValue();
    lName = range.getCell(cellRow, lNameCol).getValue();

    docName = fName + '-' + lName + '-' + id + "-certificate";

    // Get document template, copy it as a new temp doc, and save the Doc’s id
    var copyId = DriveApp.getFileById(formId).makeCopy(docName).getId();

    // Open the temporary document
    var copyDoc = DocumentApp.openById(copyId);

    // Get the document’s body section
    var copyBody = copyDoc.getActiveSection();

    // Replace place holder keys,in our google doc template
    copyBody.replaceText('@SNa@', sName);
    copyBody.replaceText('@FNa@', fName);
    copyBody.replaceText('@MNa@', mName);
    copyBody.replaceText('@LNa@', lName);

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

    // Update current index
    rowStartCell.setValue(cellRow);
  }
}

function createLetters() {
  // last update: 4/23/2018

  ////////////////////////////////////////////////////////////////////////////////////
  var folerId     = "0B3pmm9KdF9FjUjhLNkx5Rkx0eVE";
  var formId      = "1UfaYxDrU1Nd75dnj6N022p22hU8eG4soKyjKnq_8OIE";

  var idCol       = 1;
  var sNameCol    = 2;
  var fNameCol    = 3;
  var mNameCol    = 4;
  var lNameCol    = 5;

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Print");
  var range = sheet.getRange(1, 1, 130, 20); //row, col, numRows, numCols <= need to update numRows
  var rowStartCell = range.getCell(2, 6); // <= need to update column
  ////////////////////////////////////////////////////////////////////////////////////

  var docName, id, fName, lName;

  // iterate through all cells in the range
  for (var cellRow = rowStartCell.getValue()+1; ; cellRow++) {

    id = range.getCell(cellRow, idCol).getValue();
    if(id == "") { break; }

    fName = range.getCell(cellRow, fNameCol).getValue();
    lName = range.getCell(cellRow, lNameCol).getValue();

    docName = fName + '-' + lName + '-' + id + "-letter";

    // Get document template, copy it as a new temp doc, and save the Doc’s id
    var copyId = DriveApp.getFileById(formId).makeCopy(docName).getId();

    // Open the temporary document
    var copyDoc = DocumentApp.openById(copyId);

    // Get the document’s body section
    var copyBody = copyDoc.getActiveSection();

    // Replace place holder keys,in our google doc template
    copyBody.replaceText('@fname@', fName);

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

    // Update current index
    rowStartCell.setValue(cellRow);
  }
}




function emailParents() {
  return;

  var folerId     = "1tErqKzDkNu20Ud9axdcbk-DfJI2p1v8Q";
  var formId      = "1UCgE_NVE0SqsjuomxdefnDJpe-C5m8dumZUEiHaAQn8";

  var docName     = "XTRLLD";

  var idCol       = 1;
  var fNameCol    = 3;
  var lNameCol    = 5;
  var addrCol     = 9;
  //var emailCol    = 13;
  //var serviceDateCol = 15;

  var folder = DriveApp.getFolderById(folerId);

  var sm = isScriptMode();
  if(sm!="yes") {
    throw new Error("Script mode: " + sm);
    return;
  }

  //////////////////// Get global variables from the Admin sheet /////////////////////
  var parentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ParrentEmails");
  var parentRange = parentSheet.getRange(1, 1, 700, 15); //row, col, numRows, numCols
  var rowStartCell = parentRange.getCell(2, 6);
  var rowStart = rowStartCell.getValue() + 1;


  var id, fname, lname, sname, addr, docPrefix;

  // iterate through all cells in the range
  for (var cellRow = rowStart; ; cellRow++) {
    id         = parentRange.getCell(cellRow, idCol).getValue();
    fname     = parentRange.getCell(cellRow, fNameCol).getValue();
    lname     = parentRange.getCell(cellRow, lNameCol).getValue();
    //glLevel    = parentRange.getCell(cellRow, glLevelCol).getValue();
    //glName     = parentRange.getCell(cellRow, glNameCol).getValue();
    //vnLevel    = parentRange.getCell(cellRow, vnLevelCol).getValue();
    //vnName     = parentRange.getCell(cellRow, vnNameCol).getValue();
    addr       = parentRange.getCell(cellRow, addrCol).getValue();
    //email      = parentRange.getCell(cellRow, emailCol).getValue();
    //serviceDate = Utilities.formatDate(parentRange.getCell(cellRow, serviceDateCol).getValue(), "GMT", "MMM d, yyyy");

    if(id == "") break;

    // prepare string for substituion
    sname = fname + " " + lname;
    //pname = "Parents of " + sname;
    //glnam = "";
    //if(glLevel != "") {
      //glnam = "GL" + glLevel + glName;
    //}
    //vnnam = "";
    //if(vnLevel != "") {
      //vnnam = "VN" + vnLevel + vnName;
    //}

    docPrefix = id;
    //if(email == "") {
      //docPrefix = "_";
    //}

    var subject = id + "_" + docName + "_" + sname;

    // Get document template, copy it as a new temp doc, and save the Doc’s id
    var copyId = DriveApp.getFileById(templateId).makeCopy(subject).getId();

    // Open the temporary document
    var copyDoc = DocumentApp.openById(copyId);

    // Get the document’s body section
    var copyBody = copyDoc.getActiveSection();

    // Replace place holder,in our google doc template
    copyBody.replaceText('@sname@', sname);
    copyBody.replaceText('@address@', addr);
    //copyBody.replaceText('@sname@', sname);
    //copyBody.replaceText('@glname@', glnam);
    //copyBody.replaceText('@vnname@', vnnam);
    //copyBody.replaceText('@date@', serviceDate);
    //copyBody.replaceText('@loc@', loc);

    // Save and close the temporary document
    copyDoc.saveAndClose();

    // Convert temporary document to PDF
    var pdf = DriveApp.getFileById(copyId).getAs("application/pdf");

    // save pdf
    folder.createFile(pdf);

    // Delete temp file
    DriveApp.getFileById(copyId).setTrashed(true);
/*
    // Log the email address of the person running the script.
    Logger.log(Session.getActiveUser().getEmail());
    var email2 = Session.getEffectiveUser().getEmail();
    Logger.log(email2);

    // Attach PDF and send the email
    if(email != "") {
      //var body = "Mến chào quí Phụ Huynh,<br>Xin quí Phụ Huynh đem theo đơn ghi danh khi đi ghi danh cho em " + fname + " " + lname + ".  Xin cám ơn. \nTrường GLVN Andre Dũng Lạc.";
      var body = "Mến chào quí Phụ Huynh,nXin quí Phụ Huynh xem thư thông báo đính kèm. Xin cám ơn.\nBan GLVN Andre Dũng Lạc.";
      Logger.log("Email address:" + email + "<");
      email = "hle007@yahoo.com";
      MailApp.sendEmail(email, subject, body, {htmlBody: body, attachments: pdf});
    }
*/

    // Update current index
    rowStartCell.setValue(cellRow);

  }
}


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

function createAwardCertificatesGL() {
  var folerId     = getStr("GL_CERTIFICATE_FOLDER_ID");

  createCertificates(folerId, "gl-all");
}

function createAwardCertificatesVN() {
  var folerId     = getStr("VN_CERTIFICATE_FOLDER_ID");

  createCertificates(folerId, "vn-all");
}


function createCertificates(folerId, sheetname) {
  // last update: 5/8/2022

  var formId                      = getStr("CERTIFICATE_TEMPLATE_ID");
  var first_title                 = getStr("FIRST_TITLE");
  var second_title                = getStr("SECOND_TITLE");
  var third_title                 = getStr("THIRD_TITLE");
  var first_second_third_message  = getStr("FIRST_SECOND_THIRD_MESSAGE");
  var fourth_title                = getStr("FOURTH_TITLE");
  var fourth_message              = getStr("FOURTH_MESSAGE");

  ////////////////////////////////////////////////////////////////////////////////////
  var cNameCol    = 1;
  var sNameCol    = 2;
  var fNameCol    = 3;
  var lNameCol    = 4;
  var rankingCol  = 6;

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetname);
  var range = sheet.getRange(1, 1, 150, 15); //row, col, numRows, numCols <= need to update numRows
  var rowStartCell = sheet.getRange("G1:G1").getCell(1, 1); // <= need to update column
  ////////////////////////////////////////////////////////////////////////////////////
  Logger.log("start cell: " + rowStartCell.getValue());

  var cName, sName, fName, lName, ranking;

  // iterate through all cells in the range
  for (var cellRow = rowStartCell.getValue(); ; cellRow++) {

    cName = range.getCell(cellRow, cNameCol).getValue();
    if(cName == "") { break; }

    sName   = range.getCell(cellRow, sNameCol).getValue();
    fName   = range.getCell(cellRow, fNameCol).getValue();
    lName   = range.getCell(cellRow, lNameCol).getValue();
    ranking = range.getCell(cellRow, rankingCol).getValue();


    docName = ranking + '-' + cName + '-' + fName + '-' + lName;

    // Get document template, copy it as a new temp doc, and save the Doc’s id
    var copyId = DriveApp.getFileById(formId).makeCopy(docName).getId();

    // Open the temporary document
    var copyDoc = DocumentApp.openById(copyId);

    // Get the document’s body section
    var copyBody = copyDoc.getActiveSection();

    // Replace place holder keys,in our google doc template
    if (ranking == 1) {
      copyBody.replaceText('@Tit@', first_title);
      copyBody.replaceText('@Mes@', first_second_third_message);
    }
    else if (ranking == 2) {
      copyBody.replaceText('@Tit@', second_title);
      copyBody.replaceText('@Mes@', first_second_third_message);
    }
    else if (ranking == 3) {
      copyBody.replaceText('@Tit@', third_title);
      copyBody.replaceText('@Mes@', first_second_third_message);
    }
    else {
      copyBody.replaceText('@Tit@', fourth_title);
      copyBody.replaceText('@Mes@', fourth_message);
    }

    if (sName != null && sName.length > 0) {
      copyBody.replaceText('@SNa@', sName + ' ' + fName + ' ' + lName);
    }
    else {
      copyBody.replaceText('@SNa@', fName + ' ' + lName);
    }

    if (cName.charAt(0) == 'G') {
      copyBody.replaceText('@CNa@', 'Lớp Giáo Lý ' + cName);
    }
    else {
      copyBody.replaceText('@CNa@', 'Lớp Việt Ngữ ' + cName);
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

    // Update current index
    rowStartCell.setValue(cellRow);
  }
}
