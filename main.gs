// This is the main file

// Runs when installed
function onInstall(){
  onOpen();
}

// Runs when opened
function onOpen() {
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  menu.addItem("Print Receipts", "showSidebar");
  menu.addToUi();
}

// Displays the sidebar with sideBar.html as its html file.
function showSidebar(){
  var html = HtmlService.createTemplateFromFile("sideBarPrinter.html")
    .evaluate()
    .setTitle("Receipt Printer Parameters");
  SpreadsheetApp.getUi().showSidebar(html);
}

// Tester for the function finished()
function testFinished(){
  finished("C13", "F13", "I5", 3157, "Add-On/Reciepts", "I5,D7");
}

// This functions hides all the spreadsheets other than the active one. Used to print only the Receipt Print sheet.
function hide(){
  var ss = SpreadsheetApp.getActive();
  var name = ss.getActiveSheet().getName();
  var sheets = ss.getSheets();
  for (var i = 0; i<sheets.length ; i++) {
    var sheet = sheets[i];
    if (sheet.getName() !=  name){
      sheet.hideSheet();
    }
  }
}

// Unhides all sheets
function unhide(){
  var ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets();
  for (var i = 0; i<sheets.length ; i++) {
    var sheet = sheets[i];
    sheet.showSheet();
    }
}

// The main linker of this addon. It from sideBar.html and makes a call to the looper.
function finished(number, amount, uid, final, path, structure){
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
  Logger.log("Creating PDF "+filename);
  var ss = SpreadsheetApp.getActive();
  // Save to pdf
  var pdfName = filename;
  var theBlob = ss.getBlob().getAs('application/pdf').setName(pdfName);
  var newFile = folder.createFile(theBlob);
}
