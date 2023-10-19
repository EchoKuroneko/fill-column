function FillColumn(dataRange, repetition, sheetName){
    // Get the active spreadsheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // If sheetName is provided, get the sheet with that name, else get the active sheet
    var sheet = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
    
    // Get the range object based on the provided dataArrayRange
    var selectedRange = sheet.getRange(dataRange);

    // Get the values of the selected range
    var rangeValues = selectedRange.getValues();
    
    // Get the number of rows and columns in the selected range
    var numRows = selectedRange.getNumRows();
    var numCols = selectedRange.getNumColumns();
    
    var result = [];

    // Iterate through each column and row of the range
    for (var col = 0; col<numCols; col++){
      for(var row =0; row<numRows; row++){
        // Repeat the value 'repetition' number of times
        for(var iteration=0; iteration<repetition; iteration++){
          result.push(rangeValues[row][col]);
        }
        
      }
    }
    return result;
  }