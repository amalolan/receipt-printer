// Set's the final cell's value to the initial cell's amount in words
function setINR(initial, final){
  if (initial === "null"){
    Logger.log("No need to convert into amount -- from setINR");
    return
  }
  var sheet = SpreadsheetApp.getActiveSheet();
  var ans = INR(sheet.getRange(initial).getValue());
  sheet.getRange(final).setValue(ans);
  // Also returns the amount in words
  return ans;
}

function INR(input) {
  
  var a, b, c, d, e, output, outputA, outputB, outputC, outputD, outputE;
  
  var ones = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine'];
  
  if (input === 0) { // Zero  
    
    output = "Zero";
    
  } else if (input == 1) { // One
    
    output = "One only";
    
  } else { // More than one
    
    // Tens
    a = input % 100;
    outputA = oneToHundred_(a);
    
    // Hundreds
    b = Math.floor((input % 1000) / 100);
    if (b > 0 && b < 10) {
      outputB = ones[b];
    }
    
    // Thousands
    c = (Math.floor(input / 1000)) % 100;
    outputC = oneToHundred_(c);
    
    // Lakh
    d = (Math.floor(input / 100000)) % 100;
    outputD = oneToHundred_(d);
    
    // Crore
    e = (Math.floor(input / 10000000)) % 100;
    outputE = oneToHundred_(e);
    
    // Make string
    output = "";
    
    if (e > 0) {
      output = output + " " + outputE + " Crore";
    }
    
    if (d > 0) {
      output = output + " " + outputD + " Lakh";
    }
    
    if (c > 0) {
      output = output + " " + outputC + " Thousand";
    }
    
    if (b > 0) {
      output = output + " " + outputB + " Hundred";
    }
    
    if (input > 100 && a > 0) {
      output = output + " and";
    }
    
    if (a > 0) {
      output = output + " " + outputA;
    }
    
    output = output + " only";
  }
  
  return output;
  
}




function oneToHundred_(num) {
  
  var outNum;
  
  var ones = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine'];
  
  var teens = ['Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
  
  var tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
  
  if (num > 0 && num < 10) { // 1 to 9
    
    outNum = ones[num]; // ones
    
  } else if (num > 9 && num < 20) { // 10 to 19
    
    outNum = teens[(num % 10)]; // teens
    
  } else if (num > 19 && num < 100) { // 20 to 100
    
    outNum = tens[Math.floor(num / 10)]; // tens
    
    if (num % 10 > 0) {
      
      outNum = outNum + " " + ones[num % 10]; // tens + ones
      
    }
    
  }
  
  return outNum;
  
}
