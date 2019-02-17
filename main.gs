function doGet(e) {
  if(e.parameter.submit == 'summarize') {
    summarizeOrders(); 
    var template = HtmlService.createTemplate('Purchased List is summarized. <br/>' + '<a href="<?= returnUrl ?>" target="_top">戻る</a>'); 
    template.returnUrl = scriptUrl; 
    return HtmlService.createHtmlOutput(template.evaluate()); 
  }
  else if(e.parameter.supplier == 'mouser') {
    getOrderFromMouserOrderNumber(e.parameter.orderNumber); 
    var template = HtmlService.createTemplate('Sheet is created. Mouser Order Number: ' + e.parameter.orderNumber + '<br/>' + '<a href="<?= returnUrl ?>" target="_top">戻る</a>'); 
    template.returnUrl = scriptUrl; 
    return HtmlService.createHtmlOutput(template.evaluate()); 
  }
  else if(e.parameter.supplier == 'digikey') {
    //clearDigikeyService(); 
    PropertiesService.getUserProperties().setProperty('orderNumber', e.parameter.orderNumber); 
    var digikeyService = getDigikeyService(); 
    if(!digikeyService.hasAccess()) {
      var authorizationUrl = digikeyService.getAuthorizationUrl(); 
      var template = HtmlService.createTemplate('<a href="<?= authorizationUrl ?>" target="_top">Proceed to digikey API Authorize</a>.'); 
      template.authorizationUrl = authorizationUrl; 
      return HtmlService.createHtmlOutput(template.evaluate()); 
    }
    else {
      var salesOrderId = PropertiesService.getUserProperties().getProperty('orderNumber'); 
      getOrderStatus(salesOrderId); 
      
      var template = HtmlService.createTemplate('O.K. Token is already issued. <br/>' + 'Sheet is created. Digikey Order Number: ' + salesOrderId + '<br/>' + '<a href="<?= returnUrl ?>" target="_top">戻る</a>'); 
      template.returnUrl = scriptUrl; 
      return HtmlService.createHtmlOutput(template.evaluate()); 
    }
  }
  else {
    var template = HtmlService.createTemplateFromFile('index'); 
    template.returnUrl = scriptUrl; 
    return HtmlService.createHtmlOutput(template.evaluate()); 
  }
}

function doPost(e) {
  
}
