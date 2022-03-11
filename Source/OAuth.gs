const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
const clientId = configSheet.getRange(2, 2).getValue();
const clientSecret = configSheet.getRange(3, 2).getValue();
 
function getSplitwiseService() {
  return OAuth2.createService('Splitwise')
    .setAuthorizationBaseUrl('https://secure.splitwise.com/oauth/authorize')
    .setTokenUrl('https://secure.splitwise.com/oauth/token')
    .setClientId(clientId)
    .setClientSecret(clientSecret)
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
}
 
function authCallback(request) {
  var splitwiseService = getSplitwiseService();
  var isAuthorized = splitwiseService.handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput('Success! You can close this tab.');
  } else {
    return HtmlService.createHtmlOutput('Denied. You can close this tab');
  }
}
