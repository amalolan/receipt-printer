// This is the mailer's main file
function testfinishedMail(){
  finishedMail("C13", "F13", "I5", 1456, "Add-On/Reciepts", "I5,C7");
  SpreadsheetApp.getActiveSheet().getRange("I5").setValue(1455);
}

// This is called by mail.html once the user click on the Finish button
function finishedMail(number, amount, uid, final, path, structure){
  Logger.log("Reached mail merge finish");
  var folder = getFolder(path);
  var struct = structure.split(",");
  // Hide sheets
  hide();
  // Call to the looper
  looperMail(number, amount, uid, final, folder, struct);
  // Unhides sheets
  unhide();
  
}


// The looper which iterates through all the unique numbers
function looperMail(number, amount, uid, final, folder, struct) {
  Logger.log("Reached looperMail");
  var sheet = SpreadsheetApp.getActiveSheet();
  // Get the cell with the uid
  var cell = sheet.getRange(uid);
  // Start by getting the amount in words into the amount cell.
  setINR(number, amount);
  var value;
  // Loop until you reach final+1
  while (cell.getValue() <= final){
    value = cell.getValue();
    var filename = "";
    // Get the filename using the structure parameter
    for (var i = 0; i<struct.length; i++){
      // Add a space and the corresponding cell's value according to struct
      filename += sheet.getRange(struct[i]).getValue()+" ";
    }
    // Remove the last space
    filename = filename.substring(0, filename.length - 1);
    // Call to generate the pdf using this filename and the given path
    pdfMail(folder, filename, value);
    // Once generated, change the UID to the next value
    cell.setValue(value+1); 
    // Get the amount in words
    setINR(number, amount);   
  }
  Logger.log("Finished Loop at looperMail");
}

// The pdf generator
function pdfMail(folder, filename, id){
  Logger.log("Creating PDF "+filename);
  var ss = SpreadsheetApp.getActive();
  // Creates the pdf file and puts it in attach
  var pdfFile = DriveApp.getFileById(ss.getId()),
      pdf = pdfFile.getAs('application/pdf').getBytes(),
      attach = {fileName:filename+".pdf",content:pdf, mimeType:'application/pdf'};
  // Calls email
  email(attach, id);
  // Save to pdf
  if (folder === null){
    Logger.log("No folder given-- from pdfMail");
    return null;
  }
  var theBlob = ss.getBlob().getAs('application/pdf').setName(filename);
  var newFile = folder.createFile(theBlob);
}

// The heart of this file
// This is the mailing function which sends the mail to the recipients with the attachment attach.
function email(attach, id){
  Logger.log("Reached email");
  // Initializes the variables used.
  var ss = SpreadsheetApp.getActive(),
      sheet = ss.getSheetByName("Mail Merge"),
      data = sheet.getDataRange().getValues(),
      template = ss.getSheetByName("Text").getRange("A2").getValue(),
      subject = ss.getSheetByName("Text").getRange("A1").getValue(),
      message, emailTo;
  Logger.log(template);
  Logger.log(subject);
  var templateVars = template.match(/\$\{\"[^\"]+\"\}/g);
  Logger.log(templateVars);
  for (var i=0; i < data.length; i++){
    if (data[i][0] === id){
      for (var j=0; j < data[0].length; j++){
        if (data[0][j] == "email"){
          emailTo = data[i][j];
          // If there is no email address, stops running.
          if (emailTo === null || emailTo === ""){
            return null;
          }
          message = createTextFromTemplate(template, data[i], data[0]);
          subject = createTextFromTemplate(subject, data[i], data[0]);
          Logger.log('message created from row '+ (i+1) +': '+emailTo);
          MailApp.sendEmail(emailTo, subject, message, {attachments:[attach]});
        }
      }
    }
  }
}

// Generates all the required sheets for mail merge.
function generateSheetsForMail(){
  Logger.log("Reached generator");
  if (SpreadsheetApp.getActive().getSheetByName("Mail Merge") != null){
    var sheet = SpreadsheetApp.getActive().getSheetByName("Mail Merge");
    SpreadsheetApp.getActive().deleteSheet(sheet);
  }
  // Generates sheet "Mail Merge"
  var mm = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Mail Merge");
  var row1 = mm.getRange(1, 1, 1, 26);
  row1.setVerticalAlignment("middle");
  row1.setHorizontalAlignment("center");
  row1.setBackgroundRGB(100,149,237);
  row1.setFontStyle("Trebuchet MS");
  row1.setFontColor("White");
  mm.setRowHeight(1, 32);
  mm.getRange("A1").setValue("Unique ID");
  mm.getRange("B1").setValue("First Name");
  mm.getRange("C1").setValue("Last Name");
  mm.getRange("D1").setValue("email");
  mm.setFrozenRows(1);
  if (SpreadsheetApp.getActive().getSheetByName("Text") != null){
    var sheet = SpreadsheetApp.getActive().getSheetByName("Text");
    SpreadsheetApp.getActive().deleteSheet(sheet);
  }
  // Generates sheet "Text"
  var text = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Text");
  text.deleteRows(3, text.getMaxRows()-2);
  Logger.log(text.getLastRow());
  Logger.log(text.getLastColumn());
  text.deleteColumns(2, text.getMaxColumns()-1);
  text.getRange("A1").setValue("Enter Subject...");
  text.getRange("A2").setValue("Enter Message...");
  text.setRowHeight(1, 32);
  text.setRowHeight(2, 400);
  text.setColumnWidth(1, 450);
  text.getRange("A1").setHorizontalAlignment("left");
  text.getRange("A1").setVerticalAlignment("top");
  text.getRange("A2").setHorizontalAlignment("left");
  text.getRange("A2").setVerticalAlignment("top");
}
 
