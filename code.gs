/*
  =============================================================================
  Project Page: https://github.com/cmenon12/contemporary-choir
  Copyright:    (c) 2021 by Christopher Menon
  License:      GNU General Public License, version 3 (GPL-3.0)
                http://www.opensource.org/licenses/gpl-3.0.html
  =============================================================================
 */


/**
 * Make a request to the URL using the params.
 */
function makeRequest(url, params) {

  // Make the POST request
  const response = UrlFetchApp.fetch(url, params);
  const status = response.getResponseCode();
  const responseText = response.getContentText();

  // If successful then return the response text
  if (status === 200) {
    return responseText;

  // Otherwise log and throw an error
  } else {
    Logger.log(`There was a ${status} error fetching ${url}.`);
    Logger.log(responseText);
    throw Error(`There was a ${status} error fetching ${url}.`)
  }

}


/**
 * Downloads and returns all transactions.
 */
function downloadAllTransactions() {

  // Prepare the request body
  const body = {
  "client_id": getSecrets().CLIENT_ID,
  "secret": getSecrets().SECRET,
  "access_token": getSecrets().ACCESS_TOKEN,
  "options": {
    "count": 500,
    "offset": 0
  },
  "start_date": "2017-01-01",
  "end_date": "2030-01-01" 
  };

  // Condense the above into a single object
  const params = {
    "contentType": "application/json",
    "method": "post",
    "payload": JSON.stringify(body),
    "muteHttpExceptions": true
  };

  // Make the first POST request
  const result = JSON.parse(makeRequest(getSecrets().URL, params));
  const total_count = result.total_transactions;
  let offset = 0;
  let r;

  Logger.log(`There are ${total_count} transactions.`)

  // Make repeated requests
  while (offset <= total_count-1) {
    offset = offset+500;
    body.options.offset = offset;
    params.payload = JSON.stringify(body);
    r = JSON.parse(makeRequest(getSecrets().URL, params));
    result.transactions = result.transactions.concat(r.transactions);
  }

  Logger.log(`We downloaded ${result.transactions.length} transactions.`)
  return result;

}
