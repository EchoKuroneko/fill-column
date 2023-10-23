function FillColumn(dataRange, repetition){
    // Get the active spreadsheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Get the formula of the active range
    var formula = SpreadsheetApp.getActiveRange().getFormula();
    
    // Extract the arguments from the formula
    var args = formula.match(/=\w+\((.*)\)/i)[1].split(',');

    // Extract sheet name and range from the arguments
    var args = args[0].split('!');

    // If sheet name is provided, get the sheet with that name, else get the active sheet
    if (args.length === 2) {
      var sheetName = args[0].replace(/^'/, '').replace(/'$/, '');
      var dataRange = args[1];
      var sheet = ss.getSheetByName(sheetName);
    }
    else {
      dataRange = args[0];
      sheet = ss.getActiveSheet();
    }
    
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