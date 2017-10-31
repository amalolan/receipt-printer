// This is the program's main file

// Runs when installed
function onInstall(){
  onOpen();
}

// Runs when opened
function onOpen() {
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  menu.addItem("Print Receipts with Mail Merge", "showSidebarMailer");
  menu.addToUi();
}

// Displays the sidebar with mail.html as its html file.
function showSidebarMailer() {
  var ui = SpreadsheetApp.getUi();
  var html = HtmlService.createTemplateFromFile("mail.html")
    .evaluate()
    .setTitle("Receipt Printer with Mail Merge");
  SpreadsheetApp.getUi().showSidebar(html);
  ui.alert("Welcome to Receipt Printer. A wizard on the left will guide you through the proccess of printing your receipts and then mailing them.", ui.ButtonSet.OK);
}

// Before calling hide, it sets the active sheet to the given sheet containing the Receipt Template
function setAsActive(sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  sheet.showSheet();
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);

}

// This functions hides all the spreadsheets other than the active one. Used to print only the Receipt Print sheet.
function hide(){
  var ss = SpreadsheetApp.getActive();
  var name = ss.getActiveSheet().getName();
  var sheets = ss.getSheets();
  for (var i = 0; i<sheets.length ; i++) {
    var sheet = sheets[i];
    if (sheet.getName() !=  name) {
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

// Called at the end of the program
function endScript(){
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert("Thank you for using this add on. The script has finished running. You may now safely exit.", ui.ButtonSet.OK);
}

// Sends error messages to the user throught an alert box.
function failed(message){
  var ui = SpreadsheetApp.getUi();
  var msg = message || "There was an error while running the script. Please check your inputs or try again later.";
  var result = ui.alert(msg, ui.ButtonSet.OK);
}
