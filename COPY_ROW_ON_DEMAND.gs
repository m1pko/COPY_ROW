/**********************************************************************|
* Project           : Copy specified row from one Google Spreadsheep to another
*
* Script name       : COPY_ROW_ON_DEMAND.gs
*
* Date        Author      Ref    Revision (Date in YYYYMMDD format) 
* 20180310    MBarba      1      Script created
*
|**********************************************************************/

//Global variables - START

//number of columns intended to copy
var columnCount = 3;

//sheet number, in the source spreadsheet, where the row to copy is, 0 being the first sheet
var sourceSheetNumber = 0;

//sheet number, in the destiny spreadsheet, where to copy the row, 0 being the first sheet
var destinationSheetNumber = 0;

//destination spreadsheet URL
var destinationUrl = "DESTINATION_SPREADSHEET_URL";

//menu name
var menuName = "MENU";

//sub menu name
var subMenuName = "SUB_MENU";

//Global variables - END

//Auxiliary functions - START

function columnToLetter(column) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/******************************************************************|
* NAME :            setDropdown
*
* DESCRIPTION :     Inserts a Drop Down in the specified location (sheet, row, column) and using a specific data range
*
* INPUTS :
*       PARAMETERS:
*           rownNumber           Integer representing the row number
*           columnLetter         String representing the column where to insert the object
*           refDataOriginSheet   Sheet where the reference data to be used to populate the Drop Down is
*           refDataRanhe         Data range to be considered
|*******************************************************************/
function setDropdown(rowNumber,columnLetter,refDataOriginSheet,refDataRange){
  var sheet = SpreadsheetApp.openByUrl(destinationUrl);
  var dropDownList = sheet.getSheetByName(refDataOriginSheet).getRange(refDataRange);
  var options = SpreadsheetApp.newDataValidation().requireValueInRange(dropDownList).build();
  sheet.getRange(columnLetter + rowNumber).setDataValidation(options);
}

//Auxiliary functions - END

// Main functions - START

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu(menuName).addItem(subMenuName, 'requestRowNumber').addToUi();
}

function requestRowNumber() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.prompt('Please specify the row number.', ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  
  //test if input is a positive integer greater than 1
  var regExp = new RegExp("[2-9]+[0-9]*"); // "i" is for case insensitive
  var rowNumber = regExp.exec(text);
  
  if (rowNumber == null) {
    Browser.msgBox('\'' + text + '\' it\'s not a valid row number.');
    return;
  }
    
  copyRow(rowNumber);

}

function copyRow(rowNumber) {
    
  //source sheet
  var source = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = source.getSheets()[sourceSheetNumber];
  
  //get the row
  var refCell = "A"+rowNumber;
  var refColumn = columnToLetter(columnCount+1);
  var row = sourceSheet.getRange(refCell + ":" + refColumn).getValues();
  
  //destiny sheet
  var destination = SpreadsheetApp.openByUrl(destinationUrl);
  var destinationSheet = destination.getSheets()[destinationSheetNumber];
  
  var rowToAppend = [];
  for (var i = 0; i < columnCount; i++)  {
    rowToAppend[i] = row[0][i];
  }
  
  //copy from source to destination, appending as a new line at the end of the table
  destinationSheet.appendRow(rowToAppend);
  var destinationRowNumber = destinationSheet.getLastRow();

  //will create a dropdown box in column D in the destination spreadsheet with data that exists in the sheet "RefData"
  setDropdown(destinationRowNumber,"D","RefData","A2:A"); 

  //finally remove the line from the source file
  sourceSheet.deleteRow(rowNumber);
  
}

// Main functions - END
