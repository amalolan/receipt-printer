function testGetMailVars() {
  var vars = getMailVars(1456, ["email","cc","bcc"]);
  Logger.log(vars);
}

function getRecord(cell) {
  /*
  Retrieve and return the information requested by the sidebar.
  */
  if (cell === 0) return SpreadsheetApp.getActiveSheet().getActiveCell().getValue();
  else return SpreadsheetApp.getActiveSheet().getActiveCell().getA1Notation();
}

function allSheets() {
  /*
   Return an array of the names of all the sheets in the spreadsheet.
  */
  var sheets = SpreadsheetApp.getActive().getSheets();
  var names = [];
  for (var i = 0; i < sheets.length; i++) {
    names.push(sheets[i].getName());
  }
  return names;
}

function getMailVars(uid, vars) {
  /*
  Inputs:
      Number: the starting unique id
      Array of Strings: the others varaibles' indices required
  Returns an object containing the variables and their indices.
  The object also contains the actual index of the uid.
  ex. {uid: 0, email: 3, cc: 4, bcc: 5}
  */
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mail Merge"),
      data = sheet.getDataRange().getValues(),
      returning = {};
  returning["uid"] = -1;
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == uid) {
      returning["uid"] = i;
    }
  }
  for (var i = 0; i < data[0].length; i++) {
    var index = vars.indexOf(data[0][i]);
    if (index > -1) {
      returning[data[0][i]] = i;
    }
  }
  return returning;
}
