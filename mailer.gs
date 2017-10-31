// This is the mailer's main file

function testfinished(){
//  finished("C13", "F13", "I5", 1456, "Add-On/Receipts", "I5,C7", true);
  SpreadsheetApp.getActiveSheet().getRange("I5").setValue(1455);
  Logger.log(MailApp.getRemainingDailyQuota());
}

function finished(number, amount, uid, final, path, structure, sendEmail) {
  /*
  Inputs:
      number: String: Cell with the amount
      amount: String: Cell to write amount in words
      uid: String: Cell of the unique code
      final: Number: Last entry's unique code
      path: String: Destination path for the files created
      structure: String: Structure of filename (ex: I5,C7 will result a filename: val(I5) + val(C7) + '.pdf')
      sendEmail: Boolean: true if an email with the file as attachment should be sent, false otherwise

  This is called by mail.html once the user click on the Finish button. It handles everything after the user input stage
  */
  var struct = null;
  if (SpreadsheetApp.getActiveSheet().getRange(uid) == "") {
    failed("ERROR in Unique Code cell in your sheet. Please enter the Unique Code from which you want to start printing.");
    throw "shown";
  }
  if (structure) {
     struct = structure.split(",");
  }
  var folder = getFolder(path);
  var varsLocs = getMailVars(SpreadsheetApp.getActive().getRange(uid).getValue(), ["email","cc","bcc"]);
  if (Object.keys(varsLocs).length < 4) {
    failed("ERROR In email, cc, or bcc column's name in the Mail Merge sheet. Have you changed their names?");
    throw "shown";
  }
  if (varsLocs["uid"] == -1) {
    failed("ERROR: Starting Unique Code not in the first column of the Mail Merge Sheet.");
    throw "shown";
  }
  hide(); // Hide sheets
  looper(number, amount, uid, final, folder, struct, sendEmail, varsLocs); // Call to the looper
  unhide();  // Unhides sheets
}

function looper(number, amount, uid, final, folder, struct, sendEmail, varsLocs) {
  /*
  Inputs: same as finished() except for:
          struct: an array of all the cells previously comma-separated.
          folder: the folder object from the earlier path.
          varsLocs: an object consisting of the indices of the variables required for emailing.

  The looper which iterates through all the unique numbers and calls pdfMail() on all of them
  */
  var filenameCount = 1,  // The number to use to name a file if no struct is given
      ss = SpreadsheetApp.getActive(),
      sheet = SpreadsheetApp.getActiveSheet(),
      cell = sheet.getRange(uid),  // Get the cell with the uid
      value = cell.getValue(),
      mailing = {};  // This object is passed onto email()
  mailing.varsLocs = varsLocs;
  mailing.data = ss.getSheetByName("Mail Merge").getDataRange().getValues();  // This data object is used by createTextFromTemplate() and email()
  mailing.template = ss.getSheetByName("Text").getRange("A2").getValue();  // The template the user entered in the Text sheet
  mailing.subject = ss.getSheetByName("Text").getRange("A1").getValue();   // The subject the user entered

  while (value <= final) {  // Loop until you reach final+1
    // Set the amount in words
    setINR(number, amount);
    var filename = "";
    if (struct) {
      // Get the filename using the structure parameter if there is a given structure
      for (var i = 0; i<struct.length; i++){
        // Add a space and the corresponding cell's value according to struct
        try {
          filename += sheet.getRange(struct[i]).getValue()+" ";
        }
        catch(e) {
          failed("ERROR in structure of the filename input box.Please enter valid cells in it");
          throw "shown";
        }
      }
      // Remove the last space
      filename = filename.substring(0, filename.length - 1);
    }
    else {
      // just name it document along with a number to make its name unique
      filename = "document" + filenameCount;
      filenameCount++;
    }
    // Call to generate the pdf using this filename and the given path
    pdfMail(folder, filename, value, sendEmail, mailing);
    // Once generated, change the UID to the next value
    value++;
    mailing.varsLocs["uid"]++;
    cell.setValue(value);
  }
}

//
function pdfMail(folder, filename, id, sendEmail, mailing){
  /*
  Inputs:
      folder: A folder object of the dest. folder
      filename: String: the name of the pdf file
      id: Number: the current unique code
      sendEmail: Boolean: whether or not an email with the pdf should be sent
      mailing: An object passed onto email(). See email() for more info

  The pdf generator. Makes a pdf and then emails it and saves it if necessary
  */
  var ss = SpreadsheetApp.getActive();
  // Creates the pdf blob
  var pdfFile = DriveApp.getFileById(ss.getId()),
      pdf = pdfFile.getAs('application/pdf');
  pdf.setName(filename);
  // Calls email if it should send an email
  if (sendEmail) {
    email(pdf, id, mailing);
  }
  // Save to pdf if folder is given by the user
  if (folder !== null) {
    var newFile = folder.createFile(pdf);
  }
}


function email(attach, id, mailing){
  /*
  Inputs:
      attach: A blob containing the pdf
      id: the current unique code in use (ex: 1456)
      mailing: an object passed on from looper(). It contains:
        varsLocs: an object consisting of the indices of the variables required for emailing. ex: {uid: 2, email:3, cc:4, bcc:5}
        data: the data of the Mail Merge sheet as a 2D array
        template: the template for sending emails
        subject: the subject of those emails

  The heart of this file is the email function
  This is the mailing function which sends the mail to the recipients with the attachment attach.
  */
  // Initializes the variables used.
  var subject, message, emails, emailTo, options, cc, bcc, currRow;
  currRow = mailing.varsLocs["uid"];
  if (mailing.data[currRow][0] !== id) {
    failed("ERROR: Unique Code for row" + mailing.uid + " not in the first row of the Mail Merge sheet!");
    return;
  }
  message = createTextFromTemplate(mailing.template, mailing.data[currRow], mailing.data[0]);
  subject = createTextFromTemplate(mailing.subject, mailing.data[currRow], mailing.data[0]);
  emails = mailing.data[currRow][mailing.varsLocs["email"]].replace(/\s/g, "").split(",");
  emailTo = emails[0];
  if (emailTo === null || emailTo === ""){
    failed("ERROR: ID " + id + " has no corresponding email.");
    return;
  }
  cc = mailing.data[currRow][mailing.varsLocs["cc"]].replace(/\s/g, "");
  for (var i = 1; i < emails.length; i++) {
    cc +=","+ emails[i];
  }
  bcc = mailing.data[currRow][mailing.varsLocs["bcc"]].replace(/\s/g, "");
  options = {attachments:[attach], cc: cc, bcc: bcc};
  try {
    MailApp.sendEmail(emailTo, subject, message, options);
    /*
      Logger.log("Sending emailTo: " + emailTo + " with subject: " + subject + " with options: ")
      Logger.log(options);
      Logger.log("\n");
      Logger.log(message);
    */
  }
  catch(e) {
    failed("Unfortunately, the quota for sending emails for the day has been exhausted. Please try again tomorrow.\n We duly apologize for any inconvenience caused.");
    throw "shown";
  }
}
