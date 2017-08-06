// This is the printer's main file

// Tester for the function finished()
function testFinished(){
  finished("C13", "F13", "I5", 3157, "Add-On/Reciepts", "I5,D7");
}

// The main linker of this addon. It from sideBar.html and makes a call to the looper.
function finished(number, amount, uid, final, path, structure){
  Logger.log("Reached finished");
  // Call to get the folder object
  var folder = getFolder(path);
  var struct = structure.split(",");
  // Hide sheets
  hide();
  // Call to the looper
  looper(number, amount, uid, final, folder, struct);
  // Unhides sheets
  unhide();
}

// The heart of this addon. Loops through until it reaches the final+1 Unique ID.
function looper(number, amount, uid, final, folder, struct) {
  Logger.log("Reached Looper");
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
    generatePdf(folder, filename);
    // Once generated, change the UID to the next value
    cell.setValue(value+1); 
    // Get the amount in words
    setINR(number, amount);   
  }
  Logger.log("Finished Loop");
}

// The pdf generator
function generatePdf(folder, filename){
  if (folder === null){
    Logger.log("No folder given --from generatePDF()");
    return null;
  }
  Logger.log("Creating PDF "+filename);
  var ss = SpreadsheetApp.getActive();
  // Save to pdf
  var pdfName = filename;
  var theBlob = ss.getBlob().getAs('application/pdf').setName(pdfName);
  var newFile = folder.createFile(theBlob);
}
