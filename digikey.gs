// OAuth2 URL
var urlBase_digikeyAuth = 'https://sso.digikey.com/as/authorization.oauth2'; 
var urlBase_digikeyToken = 'https://sso.digikey.com/as/token.oauth2'; 
// API URL
var urlBase_digikeyOrderStatus = 'https://api.digikey.com/services/orderStatus/v2/orderStatus/'
var urlBase_digikeyPartDetails = 'https://api.digikey.com/services/partsearch/v2/partdetails'

function getPartDetailsFromDigikeyPartNumber(partNumber) {
  var token = getDigikeyService().getAccessToken(); 
  
  var urlApi = urlBase_digikeyPartDetails; 
  var headers = {
    'x-ibm-client-id': clientId_digikey, 
    'authorization' : token, 
    'content-type' : 'application/json', 
    'accept' : 'application/json'
  };
  var data = {
    "Part": partNumber,
    "IncludeAllAssociatedProducts" : "false",
    "IncludeAllForUseWithProducts" : "false", 
  };
  var payload = JSON.stringify(data);
  var options = {
    "method" : "post",
    "headers" : headers,
    "payload":payload, 
  };
  
  var response = UrlFetchApp.fetch(urlApi, options); 
  var json = JSON.parse(response.getContentText());
  
  var partDetails = {
    'Manufacture' : json.PartDetails.ManufacturerName.Text, 
    'ManufactuerPartNumber' : json.PartDetails.ManufacturerPartNumber, 
    'Packaging' : json.PartDetails.Packaging.Value, 
  }; 
  Logger.log(partDetails); 
  
  return partDetails; 
  
}

function getOrderStatus(salesOrderId) {
  var token = getDigikeyService().getAccessToken(); 
  
  var urlApi = urlBase_digikeyOrderStatus + customerId + '/' + salesOrderId + '?rootAccount=false'; 
  var headers = {
    'x-ibm-client-id': clientId_digikey, 
    'authorization' : token, 
    'accept' : 'application/json'
  };
  var options = {
      "method" : "get",
      "headers" : headers,
  };
  
  var response = UrlFetchApp.fetch(urlApi, options); 
  var json = JSON.parse(response.getContentText());
  var orderDetails = json.OrderDetails; 
  var itemCount = Object.keys(orderDetails).length; 
  
  // Create New Sheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheet = spreadsheet.getSheetByName('digikey-' + salesOrderId); 
  if(sheet != null) {
    // If salesOrderId sheet is existed, remove OrderNumber sheet
    spreadsheet.deleteSheet(sheet); 
  }
  // Create Ordernumber sheet
  sheet = createNewSheet(spreadsheet, 'digikey-' + salesOrderId); 
  
  // Set Value to sheet
  for(var i = 0; i < itemCount; i++) {
    var partDetails = getPartDetailsFromDigikeyPartNumber(orderDetails[i].PartNumber); 
    sheet.getRange(i + 2, 1).setValue(partDetails.ManufactuerPartNumber); 
    sheet.getRange(i + 2, 2).setValue(partDetails.Manufacture); 
    sheet.getRange(i + 2, 3).setValue(orderDetails[i].Description); 
    sheet.getRange(i + 2, 4).setValue(orderDetails[i].Quantity); 
    sheet.getRange(i + 2, 5).setValue(partDetails.Packaging); 
    sheet.getRange(i + 2, 6).setValue('digikey-' + salesOrderId); 
  }

}

// Callback
function authCallback_digikey(e) {
  var digikeyService = getDigikeyService(); 
  var isAuthorized = digikeyService.handleCallback(e); 
  Logger.log(e); 
  if(isAuthorized) {
//    var salesOrderId = '58460725'; 
    var salesOrderId = PropertiesService.getUserProperties().getProperty('orderNumber'); 
    getOrderStatus(salesOrderId); 
//    clearDigikeyService(); 
    var template = HtmlService.createTemplate('Success: OAuth2.0 Authorization. <br/>' + 'Sheet is created. Digikey Order Number: ' + salesOrderId + '<br/>' + '<a href="<?= returnUrl ?>" target="_top">戻る</a>'); 
    template.returnUrl = scriptUrl; 
    return HtmlService.createHtmlOutput(template.evaluate()); 
//    return HtmlService.createHtmlOutput('Success!'); 
  }
  else {
    var template = HtmlService.createTemplate('Fail: OAuth2.0 Authorization. <br/>' + '<a href="<?= returnUrl ?>" target="_top">戻る</a>'); 
    template.returnUrl = scriptUrl; 
    return HtmlService.createHtmlOutput(template.evaluate()); 
//    return HtmlService.createHtmlOutput('Denied.'); 
  }
}

// Reset Token
function clearDigikeyService() {
  OAuth2.createService('digikey')
  .setPropertyStore(PropertiesService.getUserProperties())
  .reset(); 
}

// OAuth2 Setting
function getDigikeyService() {
  return OAuth2.createService('digikey')
  .setAuthorizationBaseUrl(urlBase_digikeyAuth)
  .setTokenUrl(urlBase_digikeyToken)
  .setClientId(clientId_digikey)
  .setClientSecret(clientSecret_digikey)
  .setCallbackFunction('authCallback_digikey')
  .setPropertyStore(PropertiesService.getUserProperties())
}
  