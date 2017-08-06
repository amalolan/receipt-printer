// This is the program's main file

// Runs when installed
function onInstall(){
  onOpen();
}

// Runs when opened
function onOpen() {
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  menu.addItem("Print Receipts", "showSidebarPrinter");
  menu.addItem("Print with Mail Merge", "showSidebarMailer");
  menu.addToUi();
}

// Displays the sidebar with sideBarPrinter.html as its html file.
function showSidebarPrinter(){
  Logger.log("At showSidebarPrinter");
  var ui = SpreadsheetApp.getUi();
  var html = HtmlService.createTemplateFromFile("sideBarPrinter.html")
    .evaluate()
    .setTitle("Receipt Printer Parameters");
  SpreadsheetApp.getUi().showSidebar(html);
  ui.alert("Welcome to Receipt Printer. A wizard on the left side will guide you through the proccess of printing your receipts.", ui.ButtonSet.OK);
}

// Displays the sidebar with mail.html as its html file.
function showSidebarMailer(){
  Logger.log("At showSidebarMailer");
  var ui = SpreadsheetApp.getUi();
  var html = HtmlService.createTemplateFromFile("mail.html")
    .evaluate()
    .setTitle("Receipt Printer with Mail Merge");
  SpreadsheetApp.getUi().showSidebar(html);
  ui.alert("Welcome to Receipt Printer. A wizard on the left side will guide you through the proccess of printing your receipts and then mailing them.", ui.ButtonSet.OK);
}

// This functions hides all the spreadsheets other than the active one. Used to print only the Receipt Print sheet.
function hide(){
  Logger.log("At hide");
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
  Logger.log("At unhide");
  var ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets();
  for (var i = 0; i<sheets.length ; i++) {
    var sheet = sheets[i];
    sheet.showSheet();
    }
}

// Called at the end of the program
function endScript(){
  Logger.log("At end script");
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert("Thank you for using this add on. The script has finished running. You may now safely exit.", ui.ButtonSet.OK);
}

function failed(){
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert("There was an error while running the script. Please check your inputs and make sure you are on the receipt printing page or try again later.", ui.ButtonSet.OK);
}
