function getdata() {

  //Making the tabs we are going to use into variables
  let sheet = SpreadsheetApp.getActiveSpreadsheet()
  let sheetTarget = sheet.getSheetByName("api_response");
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
  let response = UrlFetchApp.fetch("https://pro-api.coinmarketcap.com/v1/cryptocurrency/quotes/latest?symbol="+String(tokens), {
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
  sheetTarget.getRange('A2:T').clearContent();

  //Adding time of update
  sheetTarget.getRange('T2').setValue(Object.values(status));

  //Making a loop to add the values of each token to the table
  for (let x=0; x<tokens_arr.length; x++){
        
    sheetTarget.getRange(x+2, 1).setValue(Object.values(data)[x].symbol);
    sheetTarget.getRange(x+2, 2).setValue(Object.values(data)[x].name);
    sheetTarget.getRange(x+2, 3).setValue(Object.values(data)[x].quote.USD.price);
    sheetTarget.getRange(x+2, 4).setValue(Object.values(data)[x].quote.USD.volume_24h);
    sheetTarget.getRange(x+2, 5).setValue((Object.values(data)[x].quote.USD.volume_change_24h)/100);
    sheetTarget.getRange(x+2, 6).setValue((Object.values(data)[x].quote.USD.percent_change_1h)/100);
    sheetTarget.getRange(x+2, 7).setValue((Object.values(data)[x].quote.USD.percent_change_24h)/100);
    sheetTarget.getRange(x+2, 8).setValue((Object.values(data)[x].quote.USD.percent_change_7d)/100);
    sheetTarget.getRange(x+2, 9).setValue((Object.values(data)[x].quote.USD.percent_change_30d)/100);
    sheetTarget.getRange(x+2, 10).setValue((Object.values(data)[x].quote.USD.percent_change_60d)/100);
    sheetTarget.getRange(x+2, 11).setValue((Object.values(data)[x].quote.USD.percent_change_90d)/100);
    sheetTarget.getRange(x+2, 12).setValue(Object.values(data)[x].quote.USD.market_cap);
    sheetTarget.getRange(x+2, 13).setValue(Object.values(data)[x].quote.USD.fully_diluted_market_cap);
    sheetTarget.getRange(x+2, 14).setValue(Object.values(data)[x].self_reported_circulating_supply);
    sheetTarget.getRange(x+2, 15).setValue(Object.values(data)[x].self_reported_market_cap);
    sheetTarget.getRange(x+2, 16).setValue(Object.values(data)[x].max_supply);
    sheetTarget.getRange(x+2, 17).setValue(Object.values(data)[x].circulating_supply);
    sheetTarget.getRange(x+2, 18).setValue(Object.values(data)[x].total_supply);
    sheetTarget.getRange(x+2, 19).setValue(Object.values(data)[x].tags);
    
  }
}