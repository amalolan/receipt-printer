// The file which contains functions that create sheets used by this addon.

function testGetNameOfSheet() {
  Logger.log("Testing");
  Logger.log(getNameOfSheet("Mail Merge"));
}

function getNameOfSheet(sheetName) {
  /*
  Input:
    sheetName: String: Mail Merge or Text
  Output:
    String: A unique sheet name (Ex. Mail Merge (2) or Text (1) )
  If Text or Mail Merge already exists, before creating the new ones, it returns another unique sheet name which can be used for these
  */
  var count = 0;
  var final = sheetName;
  while (SpreadsheetApp.getActive().getSheetByName(sheetName) != null) {
    count++;
    sheetName = final + " (" + count + ")";
  }
  if (count) final = final + " (" + count + ")";
  return sheetName;
}

function generateSheetsForMail(){
  /*
  Generates all the required sheets for mail merge.
  */
  if (SpreadsheetApp.getActive().getSheetByName("Mail Merge") != null) {
    var sheet = SpreadsheetApp.getActive().getSheetByName("Mail Merge");
    sheet.setName(getNameOfSheet("Mail Merge"));
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
  mm.getRange("E1").setValue("cc");
  mm.getRange("F1").setValue("bcc");
  mm.setColumnWidth(4, 200);
  mm.setColumnWidth(5, 200);
  mm.setColumnWidth(6, 200);
  mm.setFrozenRows(1);
  if (SpreadsheetApp.getActive().getSheetByName("Text") != null){
    var sheet = SpreadsheetApp.getActive().getSheetByName("Text");
    sheet.setName(getNameOfSheet("Text"));
  }
  // Generates sheet "Text"
  var text = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Text");
  text.deleteRows(3, text.getMaxRows()-2);
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
