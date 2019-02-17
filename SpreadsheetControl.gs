function createNewSheet(spreadsheet, sheetname) {
  var sheet = spreadsheet.insertSheet(sheetname); 
  sheet.activate(); 
  spreadsheet.moveActiveSheet(1); 
  sheet.getRange(1, 1).setValue('Manufacturer Part Number'); 
  sheet.getRange(1, 2).setValue('Manufacturer'); 
  sheet.getRange(1, 3).setValue('Description'); 
  sheet.getRange(1, 4).setValue('Quantity'); 
  sheet.getRange(1, 5).setValue('Packaging'); 
  sheet.getRange(1, 6).setValue('OrderNumber'); 
  
  return sheet; 
}

function createSummarizeSheet(spreadsheet) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheet = spreadsheet.getSheetByName('TotalCount'); 
  if(sheet != null) {
    spreadsheet.deleteSheet(sheet); 
  }
  sheet = createNewSheet(spreadsheet, 'TotalCount'); 
  
  return sheet; 
}
  
function checkSheetLengh() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); 
  var sourceSheets = spreadsheet.getSheets(); 
  Logger.log("Sheets Length = " + sourceSheets.length); 
}  

function summarizeOrders() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); 

  var pivotSheet = spreadsheet.getSheetByName('Pivot_TotalCount'); 
  if(pivotSheet == null) {
    sheet = createNewSheet(spreadsheet, 'Pivot_TotalCount'); 
  }
  pivotSheet.activate(); 
  spreadsheet.moveActiveSheet(1); 

  var targetSheet = createSummarizeSheet(spreadsheet); 
  targetSheet.activate(); 
  spreadsheet.moveActiveSheet(1); 
  rowCount = 2; 

  var sourceSheets = spreadsheet.getSheets(); 
  Logger.log("Sheets Length = " + sourceSheets.length); 
  
  for(var i = 1; i < sourceSheets.length; i++) {
    row = sourceSheets[i].getLastRow(); 
    col = sourceSheets[i].getLastColumn(); 
    data = sourceSheets[i].getRange(2, 1, row, col).getValues(); 
    Logger.log("Summarize: " + sourceSheets[i].getSheetName() + " / GetRange to (" + row + ", " + col + ")"); 
    for(var j = 0; j < row - 1; j++) {
      for(var k = 0; k < 6; k++) {
        targetSheet.getRange(rowCount, k + 1).setValue(data[j][k]); 
      }
      rowCount++; 
    }
  }

  
}  
  