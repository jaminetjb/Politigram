function onEdit(e){
  //Calculate the maybe answers
  calculateQ1();
  calculateQ2();
  calculateQ3();
  calculateQ4();
  calculateQ5();
  calculateQ6();
  calculateQ7();
  calculateQ8();
  calculateQ9();
  calculateQ10();
  calculateQ11();
  calculateQ12();
  calculateQ13();
  calculateQ14();
  calculateQ15();
  answer();
  
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];

 // This logs the value in the very last cell of this sheet
 var lastRow = sheet.getLastRow();
  //Logger.log(lastRow);
  //Emails and links
  var subject = "Your Politigram Score";
  var link1='https://politgramright.weebly.com';
  var link2='https://politgramcenter.weebly.com';
  var link3='https://politgramleft.weebly.com';
  var score = getScore().getValue();
  Logger.log(score);
  var email = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getRange(lastRow, 2);
  
  if(score < 50){
    GmailApp.sendEmail(email.getValue(),"Politgram Score", "Thank you for answering the PolitGram questionnaire! Your results have now been analyzed. According to your answers, you side most with Right. \n Click on the link below to get more informations about your political views: \n https://politgramright.weebly.com");
  }
  
  if(score >100) {
    GmailApp.sendEmail(email.getValue(),"Politgram Score", "Thank you for answering the PolitGram questionnaire! Your results have now been analyzed. According to your answers, you side most with Left. \n Click on the link below to get more informations about your political views: \n https://politgramleft.weebly.com");
  }
  
  if (score < 100 && score > 50) {
    GmailApp.sendEmail(email.getValue(),"Politgram Score", "Thank you for answering the PolitGram questionnaire! Your results have now been analyzed. According to your answers, you side most with Center. \n Click on the link below to get more informations about your political views: \n https://politgramcenter.weebly.com");
  }
}

function getScore() {
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = ss.getSheets()[0];

 // This logs the value in the very last cell of this sheet
 var lastRow = sheet.getLastRow();
 var scoreColumn = "AK"+lastRow;
 
 var scoreCell = sheet.getRange(lastRow, 37);
 var email = sheet.getRange(lastRow, 2);
 Logger.log(scoreCell.getValue());
 ////Logger.log(email.getValue());
 
 return scoreCell;
}

function getEmail() {
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = ss.getSheets()[0];

 // This logs the value in the very last cell of this sheet
 var lastRow = sheet.getLastRow();
 var lastColumn = sheet.getLastColumn();
 
 var lastCell = sheet.getRange(lastRow, lastColumn);
 var email = sheet.getRange(lastRow, 2);
 ////Logger.log(lastCell.getValue());
 //Logger.log(email.getValue());
 
 return email;
}

function answer() {
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = ss.getSheets()[0];
 var lastRow = sheet.getLastRow();
  
 // When the "numRows" argument is used, only a single column of data is returned.
 var range = sheet.getRange(lastRow, 22, 1, 15);
 var values = range.getValues();
  //Logger.log(lastRow);
  //Logger.log(range);
  //Logger.log(values);
  
  var sum = 0;
  for (var i in values[0]) {
    sum += values[0][i];
  }
  //Logger.log(sum)
  
  var rawScoreCell = "C"+lastRow;
  var totScore = sum + sheet.getRange(rawScoreCell).getValue();
  var scorecell = "AK"+lastRow;
  //c lastrow + sum 
  SpreadsheetApp.getActiveSheet().getRange(scorecell).setValue(totScore);

}


function calculateQ1() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  
  //Find the last row and set the index for the value to be written to
  var lastRow = sheet.getLastRow();
  cellQ1 = "V"+lastRow; //the cell where the calculated value will be stored
  
  //Gets value for the response to question 1
  cellQ1ans = "E"+lastRow;
  var Q1ans = sheet.getRange(cellQ1ans).getValue();
  
  //Check if the Maybe options was chosen
  var val = 0;
  if ((Q1ans != "Yes") && (Q1ans != "Yes") && (Q1ans != "No") && (Q1ans != "Increase") && (Q1ans != "Decrease")) {
    val = 5;
  }
  
  //Write the answer to the spreadsheet
  SpreadsheetApp.getActiveSheet().getRange(cellQ1).setValue(val);
}
function calculateQ2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  
  //Find the last row and set the index for the value to be written to
  var lastRow = sheet.getLastRow();
  cellQ1 = "W"+lastRow; //the cell where the calculated value will be stored
  
  //Gets value for the response to question 1
  cellQ1ans = "F"+lastRow;
  var Q1ans = sheet.getRange(cellQ1ans).getValue();
  
  //Check if the Maybe options was chosen
  var val = 0;
  if ((Q1ans != "Yes") && (Q1ans != "Yes") && (Q1ans != "No") && (Q1ans != "Increase") && (Q1ans != "Decrease")) {
    val = 5;
  }
  
  //Write the answer to the spreadsheet
  SpreadsheetApp.getActiveSheet().getRange(cellQ1).setValue(val);
}
function calculateQ3() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  
  //Find the last row and set the index for the value to be written to
  var lastRow = sheet.getLastRow();
  cellQ1 = "X"+lastRow; //the cell where the calculated value will be stored
  
  //Gets value for the response to question 1
  cellQ1ans = "G"+lastRow;
  var Q1ans = sheet.getRange(cellQ1ans).getValue();
  
  //Check if the Maybe options was chosen
  var val = 0;
  if ((Q1ans != "Yes") && (Q1ans != "Yes") && (Q1ans != "No") && (Q1ans != "Increase") && (Q1ans != "Decrease")) {
    val = 5;
  }
  
  //Write the answer to the spreadsheet
  SpreadsheetApp.getActiveSheet().getRange(cellQ1).setValue(val);
}
function calculateQ4() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  
  //Find the last row and set the index for the value to be written to
  var lastRow = sheet.getLastRow();
  cellQ1 = "Y"+lastRow; //the cell where the calculated value will be stored
  
  //Gets value for the response to question 1
  cellQ1ans = "H"+lastRow;
  var Q1ans = sheet.getRange(cellQ1ans).getValue();
  
  //Check if the Maybe options was chosen
  var val = 0;
  if ((Q1ans != "Yes") && (Q1ans != "Yes") && (Q1ans != "No") && (Q1ans != "Increase") && (Q1ans != "Decrease")) {
    val = 5;
  }
  
  //Write the answer to the spreadsheet
  SpreadsheetApp.getActiveSheet().getRange(cellQ1).setValue(val);
}
function calculateQ5() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  
  //Find the last row and set the index for the value to be written to
  var lastRow = sheet.getLastRow();
  cellQ1 = "Z"+lastRow; //the cell where the calculated value will be stored
  
  //Gets value for the response to question 1
  cellQ1ans = "I"+lastRow;
  var Q1ans = sheet.getRange(cellQ1ans).getValue();
  
  //Check if the Maybe options was chosen
  var val = 0;
  if ((Q1ans != "Yes") && (Q1ans != "Yes") && (Q1ans != "No") && (Q1ans != "Increase") && (Q1ans != "Decrease")) {
    val = 5;
  }
  
  //Write the answer to the spreadsheet
  SpreadsheetApp.getActiveSheet().getRange(cellQ1).setValue(val);
}
function calculateQ6() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  
  //Find the last row and set the index for the value to be written to
  var lastRow = sheet.getLastRow();
  cellQ1 = "AA"+lastRow; //the cell where the calculated value will be stored
  
  //Gets value for the response to question 1
  cellQ1ans = "J"+lastRow;
  var Q1ans = sheet.getRange(cellQ1ans).getValue();
  
  //Check if the Maybe options was chosen
  var val = 0;
  if ((Q1ans != "Yes") && (Q1ans != "Yes") && (Q1ans != "No") && (Q1ans != "Increase") && (Q1ans != "Decrease")) {
    val = 5;
  }
  
  //Write the answer to the spreadsheet
  SpreadsheetApp.getActiveSheet().getRange(cellQ1).setValue(val);
}
function calculateQ7() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  
  //Find the last row and set the index for the value to be written to
  var lastRow = sheet.getLastRow();
  cellQ1 = "AB"+lastRow; //the cell where the calculated value will be stored
  
  //Gets value for the response to question 1
  cellQ1ans = "K"+lastRow;
  var Q1ans = sheet.getRange(cellQ1ans).getValue();
  
  //Check if the Maybe options was chosen
  var val = 0;
  if ((Q1ans != "Yes") && (Q1ans != "Yes") && (Q1ans != "No") && (Q1ans != "Increase") && (Q1ans != "Decrease")) {
    val = 5;
  }
  
  //Write the answer to the spreadsheet
  SpreadsheetApp.getActiveSheet().getRange(cellQ1).setValue(val);
}
function calculateQ8() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  
  //Find the last row and set the index for the value to be written to
  var lastRow = sheet.getLastRow();
  cellQ1 = "AC"+lastRow; //the cell where the calculated value will be stored
  
  //Gets value for the response to question 1
  cellQ1ans = "L"+lastRow;
  var Q1ans = sheet.getRange(cellQ1ans).getValue();
  
  //Check if the Maybe options was chosen
  var val = 0;
  if ((Q1ans != "Yes") && (Q1ans != "Yes") && (Q1ans != "No") && (Q1ans != "Increase") && (Q1ans != "Decrease")) {
    val = 5;
  }
  
  //Write the answer to the spreadsheet
  SpreadsheetApp.getActiveSheet().getRange(cellQ1).setValue(val);
}
function calculateQ9() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  
  //Find the last row and set the index for the value to be written to
  var lastRow = sheet.getLastRow();
  cellQ1 = "AD"+lastRow; //the cell where the calculated value will be stored
  
  //Gets value for the response to question 1
  cellQ1ans = "M"+lastRow;
  var Q1ans = sheet.getRange(cellQ1ans).getValue();
  
  //Check if the Maybe options was chosen
  var val = 0;
  if ((Q1ans != "Yes") && (Q1ans != "Yes") && (Q1ans != "No") && (Q1ans != "Increase") && (Q1ans != "Decrease")) {
    val = 5;
  }
  
  //Write the answer to the spreadsheet
  SpreadsheetApp.getActiveSheet().getRange(cellQ1).setValue(val);
}
function calculateQ10() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  
  //Find the last row and set the index for the value to be written to
  var lastRow = sheet.getLastRow();
  cellQ1 = "AE"+lastRow; //the cell where the calculated value will be stored
  
  //Gets value for the response to question 1
  cellQ1ans = "N"+lastRow;
  var Q1ans = sheet.getRange(cellQ1ans).getValue();
  
  //Check if the Maybe options was chosen
  var val = 0;
  if ((Q1ans != "Yes") && (Q1ans != "Yes") && (Q1ans != "No") && (Q1ans != "Increase") && (Q1ans != "Decrease")) {
    val = 5;
  }
  
  //Write the answer to the spreadsheet
  SpreadsheetApp.getActiveSheet().getRange(cellQ1).setValue(val);
}
function calculateQ11() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  
  //Find the last row and set the index for the value to be written to
  var lastRow = sheet.getLastRow();
  cellQ1 = "AF"+lastRow; //the cell where the calculated value will be stored
  
  //Gets value for the response to question 1
  cellQ1ans = "O"+lastRow;
  var Q1ans = sheet.getRange(cellQ1ans).getValue();
  
  //Check if the Maybe options was chosen
  var val = 0;
  if ((Q1ans != "Yes") && (Q1ans != "Yes") && (Q1ans != "No") && (Q1ans != "Increase") && (Q1ans != "Decrease")) {
    val = 5;
  }
  
  //Write the answer to the spreadsheet
  SpreadsheetApp.getActiveSheet().getRange(cellQ1).setValue(val);
}
function calculateQ12() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  
  //Find the last row and set the index for the value to be written to
  var lastRow = sheet.getLastRow();
  cellQ1 = "AG"+lastRow; //the cell where the calculated value will be stored
  
  //Gets value for the response to question 1
  cellQ1ans = "P"+lastRow;
  var Q1ans = sheet.getRange(cellQ1ans).getValue();
  
  //Check if the Maybe options was chosen
  var val = 0;
  if ((Q1ans != "Yes") && (Q1ans != "Yes") && (Q1ans != "No") && (Q1ans != "Increase") && (Q1ans != "Decrease")) {
    val = 5;
  }
  
  //Write the answer to the spreadsheet
  SpreadsheetApp.getActiveSheet().getRange(cellQ1).setValue(val);
}
function calculateQ13() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  
  //Find the last row and set the index for the value to be written to
  var lastRow = sheet.getLastRow();
  cellQ1 = "AH"+lastRow; //the cell where the calculated value will be stored
  
  //Gets value for the response to question 1
  cellQ1ans = "Q"+lastRow;
  var Q1ans = sheet.getRange(cellQ1ans).getValue();
  
  //Check if the Maybe options was chosen
  var val = 0;
  if ((Q1ans != "Yes") && (Q1ans != "Yes") && (Q1ans != "No") && (Q1ans != "Increase") && (Q1ans != "Decrease")) {
    val = 5;
  }
  
  //Write the answer to the spreadsheet
  SpreadsheetApp.getActiveSheet().getRange(cellQ1).setValue(val);
}
function calculateQ14() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  
  //Find the last row and set the index for the value to be written to
  var lastRow = sheet.getLastRow();
  cellQ1 = "AI"+lastRow; //the cell where the calculated value will be stored
  
  //Gets value for the response to question 1
  cellQ1ans = "R"+lastRow;
  var Q1ans = sheet.getRange(cellQ1ans).getValue();
  
  //Check if the Maybe options was chosen
  var val = 0;
  if ((Q1ans != "Yes") && (Q1ans != "Yes") && (Q1ans != "No") && (Q1ans != "Increase") && (Q1ans != "Decrease")) {
    val = 5;
  }
  
  //Write the answer to the spreadsheet
  SpreadsheetApp.getActiveSheet().getRange(cellQ1).setValue(val);
}
function calculateQ15() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  
  //Find the last row and set the index for the value to be written to
  var lastRow = sheet.getLastRow();
  cellQ1 = "AJ"+lastRow; //the cell where the calculated value will be stored
  
  //Gets value for the response to question 1
  cellQ1ans = "S"+lastRow;
  var Q1ans = sheet.getRange(cellQ1ans).getValue();
  
  //Check if the Maybe options was chosen
  var val = 0;
  if ((Q1ans != "Yes") && (Q1ans != "Yes") && (Q1ans != "No") && (Q1ans != "Increase") && (Q1ans != "Decrease")) {
    val = 5;
  }
  
  //Write the answer to the spreadsheet
  SpreadsheetApp.getActiveSheet().getRange(cellQ1).setValue(val);
}