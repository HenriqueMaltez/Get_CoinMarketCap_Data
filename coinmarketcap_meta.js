function getmetadata() {

  //Making the tabs we are going to use into variables
  let sheet = SpreadsheetApp.getActiveSpreadsheet()
  let sheetTarget = sheet.getSheetByName("api_response_meta");
  let sheetSource = sheet.getSheetByName("to_api");

  //Making the ranges we are going to use into variables
  let API_KEY = sheetSource.getRange('B6').getDisplayValue().trim();
  let tokens = sheetSource.getRange('B4').getDisplayValue().trim();
  let tokens_range = sheetSource.getRange('A2:A').getValues();
  let tokens_arr = []

  //Checking if you have added at leat one symbol
  if (!String(tokens_range[0])) sheet.msgBox("You need to put at least one symbol in the tab 'data'")
  if (!String(tokens_range[0])) { return; }

  //Checking if you have added the API KEY in the 'to_api' tab
  if (!String(API_KEY)) sheet.msgBox("You need to put your API KEY in the tab 'to_api'")
  if (!String(API_KEY)) { return; }

  for (let y=0; y<tokens_range.length; y++){
    
    if (!String(tokens_range[y])) break

    tokens_arr.push(tokens_range[y])
  }

  //Setting up the request we are going to send to the API  
  let response = UrlFetchApp.fetch("https://pro-api.coinmarketcap.com/v1/cryptocurrency/info?symbol="+String(tokens), {
    "method": "GET",
    headers: {
    'X-CMC_PRO_API_KEY': String(API_KEY),
    'Content-Type': 'application/json'
  },
    "muteHttpExceptions": true
  }) 

  let tokens_data = JSON.parse(response);
  Logger.log(tokens_data);  

  let data = tokens_data.data;
  let status = tokens_data.status;

  //Clearing table to receive new values
  sheetTarget.getRange('A2:D').clearContent();

  //Adding time of update
  sheetTarget.getRange('D2').setValue(Object.values(status));

  //Making a loop to add the values of each token to the table
  for (let x=0; x<tokens_arr.length; x++){

    sheetTarget.getRange(x+2, 1).setValue((Object.values(data)[x].symbol));
    sheetTarget.getRange(x+2, 2).setValue((Object.values(data)[x].urls.website));
    sheetTarget.getRange(x+2, 3).setValue((Object.values(data)[x].urls.twitter));

  }
}