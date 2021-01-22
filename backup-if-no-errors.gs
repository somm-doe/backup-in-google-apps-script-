/**
 * This function checks for errors and blank cells in Google Sheets before making a duplicate.
 * Make sure that "dataviz" is your first sheet when you import it to your dataviz tool.
 *
 * The function needs to be triggered with a Google Apps Scripts trigger
 * at https://script.google.com/home/all
 **/

function backupIfNoError() {
  var ss = SpreadsheetApp.getActive();
  var sourceSheet = ss.getSheetByName("source"); // Source sheet name
  var targetSheet = ss.getSheetByName("dataviz"); // Target sheet name
  var A1notation = "A1:E12"; // A1 notation of the source range (only for errors)
  var A1notationA = "C2:C12"; // check empty cells in C
  var A1notationB = "E2:E12"; // check empty cells in E
  var sourceRange = sourceSheet.getRange(A1notation);
  var sourceRangeA = sourceSheet.getRange(A1notationA);
  var sourceRangeB = sourceSheet.getRange(A1notationB);
  var sourceValues = sourceRange.getValues();
  var sourceValuesA = sourceRangeA.getValues();  
  var sourceValuesB = sourceRangeB.getValues(); 
  var errors = ["#REF!", "#N/A", "#VALUE!", "#DIV/0!", "#NAME?"]; // Array of errors (translate when neccessary if your local Google language is not English)
  for (var i = 0; i < sourceValues.length; i++) { // Iterate through all values in source range
    for (var j = 0; j < sourceValues[i].length; j++) {
      if (errors.includes(sourceValues[i][j])) return; // End execution if error found in a cell
    }
  }
  // Check if there is an inbetween empty row:
  var emptyRowInBetweenA = sourceValuesA.map(rowValues => { 
    return rowValues.every(cellValue => {
      return cellValue === "";
    });
  }).join().indexOf("true,false") !== -1;
  var emptyRowInBetweenB = sourceValuesB.map(rowValues => { 
    return rowValues.every(cellValue => {
      return cellValue === "";
    });
  }).join().indexOf("true,false") !== -1;
  var targetRange = targetSheet.getRange("A1"); // Range to copy to (please change if necessary)
  if (!emptyRowInBetweenA && !emptyRowInBetweenB) sourceRange.copyTo(targetRange, {contentsOnly:true}); // Copy range to specified destination
}
