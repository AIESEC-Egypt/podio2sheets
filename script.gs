// Replace these values with your Podio app credentials
var sheet3 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SUOGV');

// make sure to put your info in the belw variables
var CLIENT_ID = '';
var CLIENT_SECRET = '';
// include this url in your keyAPIs settings
var REDIRECT_URI = 'script.google.com';
var SCOPE = 'view modify';
var PODIO_APP_ID = '';

function authenticatePodioGV() {
  // Initialize the OAuth2 library
  var oauth2Service = OAuth2.createService('podio')
    .setAuthorizationBaseUrl('https://podio.com/oauth/authorize')
    .setTokenUrl('https://podio.com/oauth/token')
    .setClientId(CLIENT_ID)
    .setClientSecret(CLIENT_SECRET)
    .setCallbackFunction('authCallbackGV')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope(SCOPE);

  // Check if the user has authorized the script
  if (!oauth2Service.hasAccess()) {
    var authorizationUrl = oauth2Service.getAuthorizationUrl();
    Logger.log('Open the following URL and re-run the script: %s', authorizationUrl);
  } else {
    // Your authenticated code here
    Logger.log('Successfully authenticated!');
    // Now, you can make Podio API calls using the access token.
    var accessToken = oauth2Service.getAccessToken();
    // Make Podio API calls here using the accessToken
    Logger.log(accessToken)

    // Make a GET request to retrieve data from Podio
    var headers = {
        'Authorization': 'Bearer ' + accessToken,
        'Content-Type': 'application/json'
    };
    var options = {
        'method': 'get',
        'headers': headers,
    };

    Logger.log(headers['Authorization']); // Log the authorization header
    for(let offset = 10000; offset >= 0 ; offset-=500){
      var podioUrl = `https://api.podio.com/item/app/${PODIO_APP_ID}?limit=500&offset=${offset}`;
      try {
          
          var response = UrlFetchApp.fetch(podioUrl, options);
          var responseData = JSON.parse(response.getContentText());
          if(responseData['items'].length == 0){
            Logger.log(responseData['items'].length);
            continue;
          }
          let data = [];
          let ids = sheet3.getRange(2,1,sheet3.getLastRow(),1).getValues().flat(1);
          // Logger.log(ids)
          let fields = ['Created at', 'Full Name', 'Sign-up By', 'Home LC', 'Email', 'Phone', 'City you currently live in', 'University you currently study or the last one you finished studying', 'Department', 'Birthdate', 'Age', 'EP Manager', 'Status', 're APD?', 'Has fee reduction?', 'Has ERASMUS scholarship?', 'Countries the EP is interested in', 're-approach later', "Reason EP didin't continue", 'How did you hear about AIESEC?']
          
          for(let i = 0; i < responseData['items'].length; i++){
            Logger.log(responseData['items'][i]['app_item_id'])
              if(ids.indexOf(responseData['items'][i]['app_item_id'])>-1){
                Logger.log("old")
                let row = [responseData['items'][i]['app_item_id']];
                for(let field of fields){
                  let flag = 0;
                  for(let j = 0; j < 30; j++){
                      let x = responseData['items'][i]['fields'][j]?responseData['items'][i]['fields'][j]['label']:"-";
                      if(x == "-")continue;
                      if(x == field){
                          flag = 1;
                          if(field == 'Created at')row.push(responseData['items'][i]['fields'][j]['values'][0]['start'])
                          else if(field == 'Full Name')row.push(responseData['items'][i]['fields'][j]['values'][0]['value'])
                          else if(field == 'Sign-up By')row.push(responseData['items'][i]['fields'][j]['values'][0]['value']['name'])
                          else if(field == 'Home LC')row.push(responseData['items'][i]['fields'][j]['values'][0]['value']['text'])
                          else if(field == 'Email')row.push(responseData['items'][i]['fields'][j]['values'][0]['value'])
                          else if(field == 'Phone')row.push(responseData['items'][i]['fields'][j]['values'][0]['value'])
                          else if(field == 'City you currently live in')row.push(responseData['items'][i]['fields'][j]['values'][0]['value']['text'])
                          else if(field == 'University you currently study or the last one you finished studying')row.push(responseData['items'][i]['fields'][j]['values'][0]['value']['text'])
                          else if(field == 'Department')row.push(responseData['items'][i]['fields'][j]['values'][0]['value']['text'])
                          else if(field == 'Birthdate')row.push(responseData['items'][i]['fields'][j]['values'][0]['start_date'])
                          else if(field == 'Age')row.push(responseData['items'][i]['fields'][j]['values'][0]['value'])
                          else if(field == 'EP Manager')row.push(responseData['items'][i]['fields'][j]['values'][0]['value']['name'])
                          else if(field == 'Status')row.push(responseData['items'][i]['fields'][j]['values'][0]['value']['text'])
                          else if(field == 're APD?')row.push(responseData['items'][i]['fields'][j]['values'][0]['value']['text'])
                          else if(field == 'Has fee reduction?')row.push(responseData['items'][i]['fields'][j]['values'][0]['value']['text'])
                          else if(field == 'Has ERASMUS scholarship?')row.push(responseData['items'][i]['fields'][j]['values'][0]['value']['text'])
                          else if(field == 'Countries the EP is interested in')row.push(responseData['items'][i]['fields'][j]['values'][0]['value']['text'])
                          else if(field == 're-approach later')row.push(responseData['items'][i]['fields'][j]['values'][0]['value']['text'])
                          else if(field == "Reason EP didin't continue")row.push(responseData['items'][i]['fields'][j]['values'][0]['value']['text'])
                          else if(field == 'How did you hear about AIESEC?')row.push(responseData['items'][i]['fields'][j]['values'][0]['value']['text'])
                      }
                  }
                  if(flag == 0)row.push("-")
                }
                let index = ids.indexOf(responseData['items'][i]['app_item_id']);
                Logger.log(responseData['items'][i]['app_item_id']);
                Logger.log(index);
                let rows = [];
                rows.push(row);
                sheet3.getRange(index + 2, 1, 1, rows[0].length).setValues(rows)
              }else{
                Logger.log("new")
                let row = [responseData['items'][i]['app_item_id']];
                for(let field of fields){
                  let flag = 0;
                  for(let j = 0; j < 30; j++){
                      let x = responseData['items'][i]['fields'][j]?responseData['items'][i]['fields'][j]['label']:"-";
                      if(x == "-")continue;
                      if(x == field){
                          flag = 1;
                          if(field == 'Created at')row.push(responseData['items'][i]['fields'][j]['values'][0]['start'])
                          else if(field == 'Full Name')row.push(responseData['items'][i]['fields'][j]['values'][0]['value'])
                          else if(field == 'Sign-up By')row.push(responseData['items'][i]['fields'][j]['values'][0]['value']['name'])
                          else if(field == 'Home LC')row.push(responseData['items'][i]['fields'][j]['values'][0]['value']['text'])
                          else if(field == 'Email')row.push(responseData['items'][i]['fields'][j]['values'][0]['value'])
                          else if(field == 'Phone')row.push(responseData['items'][i]['fields'][j]['values'][0]['value'])
                          else if(field == 'City you currently live in')row.push(responseData['items'][i]['fields'][j]['values'][0]['value']['text'])
                          else if(field == 'University you currently study or the last one you finished studying')row.push(responseData['items'][i]['fields'][j]['values'][0]['value']['text'])
                          else if(field == 'Department')row.push(responseData['items'][i]['fields'][j]['values'][0]['value']['text'])
                          else if(field == 'Birthdate')row.push(responseData['items'][i]['fields'][j]['values'][0]['start_date'])
                          else if(field == 'Age')row.push(responseData['items'][i]['fields'][j]['values'][0]['value'])
                          else if(field == 'EP Manager')row.push(responseData['items'][i]['fields'][j]['values'][0]['value']['name'])
                          else if(field == 'Status')row.push(responseData['items'][i]['fields'][j]['values'][0]['value']['text'])
                          else if(field == 're APD?')row.push(responseData['items'][i]['fields'][j]['values'][0]['value']['text'])
                          else if(field == 'Has fee reduction?')row.push(responseData['items'][i]['fields'][j]['values'][0]['value']['text'])
                          else if(field == 'Has ERASMUS scholarship?')row.push(responseData['items'][i]['fields'][j]['values'][0]['value']['text'])
                          else if(field == 'Countries the EP is interested in')row.push(responseData['items'][i]['fields'][j]['values'][0]['value']['text'])
                          else if(field == 're-approach later')row.push(responseData['items'][i]['fields'][j]['values'][0]['value']['text'])
                          else if(field == "Reason EP didin't continue")row.push(responseData['items'][i]['fields'][j]['values'][0]['value']['text'])
                          else if(field == 'How did you hear about AIESEC?')row.push(responseData['items'][i]['fields'][j]['values'][0]['value']['text'])
                      }
                  }
                  if(flag == 0)row.push("-")
                }
                data.unshift(row);
              }            
          }
          Logger.log(data)
          Logger.log(data.length)
          if(data.length > 0){
            Logger.log('log fuckkk')
              sheet3.getRange(sheet3.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
          }
      } catch (error) {
          Logger.log(error.toString());
      }
    }
  }
}

function authCallbackGV(request) {
  var oauth2Service = OAuth2.createService('podio')
    .setAuthorizationBaseUrl('https://podio.com/oauth/authorize')
    .setTokenUrl('https://podio.com/oauth/token')
    .setClientId(CLIENT_ID)
    .setClientSecret(CLIENT_SECRET)
    .setCallbackFunction('authCallbackGV')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope(SCOPE);

  var isAuthorized = oauth2Service.handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput('Success! You can close this tab.');
  } else {
    return HtmlService.createHtmlOutput('Denied. You can close this tab');
  }
}
