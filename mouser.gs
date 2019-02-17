var urlBase_mouser = 'https://www.mouser.jp/api/v1'; 

function getOrderFromMouserOrderNumber(orderNumber) {
  var urlApi = urlBase_mouser + '/order/' + orderNumber + '?' + 'apiKey=' + apiKey_mouser; 
  
  var response = UrlFetchApp.fetch(urlApi); 
  var json = JSON.parse(response.getContentText());
  Logger.log(json); 
  var orderDetails = json.OrderLines; 
  var itemCount = Object.keys(orderDetails).length; 
  
  // Create New Sheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheet = spreadsheet.getSheetByName('mouser-' + orderNumber); 
  if(sheet != null) {
    // If OrderNumber sheet is existed, remove OrderNumber sheet
    spreadsheet.deleteSheet(sheet); 
  }
  // Create Ordernumber sheet
  sheet = createNewSheet(spreadsheet, 'mouser-' + orderNumber); 
  
  // Set Value to sheet
  for(var i = 0; i < itemCount; i++) {
    sheet.getRange(i + 2, 1).setValue(orderDetails[i].MfrPartNumber); 
    sheet.getRange(i + 2, 2).setValue(orderDetails[i].Manufacturer); 
    sheet.getRange(i + 2, 3).setValue(orderDetails[i].Description); 
    sheet.getRange(i + 2, 4).setValue(orderDetails[i].Quantity); 
    sheet.getRange(i + 2, 5).setValue(""); 
    sheet.getRange(i + 2, 6).setValue('mouser-' + orderNumber); 
  }
  
}


