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
 * 
 * @param {string} url the URL to make the request to.
 * @param {Object} params the params to use with the request.
 * @return {string} the text of the response if successful.
 * @throws {Error} response status code was not 200.
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
 * 
 * @return {Object} the result of transactions.get, with all transactions.
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

  Logger.log(`There are ${total_count} transactions in Plaid.`);

  // Make repeated requests
  while (offset <= total_count - 1) {
    offset = offset + 500;
    body.options.offset = offset;
    params.payload = JSON.stringify(body);
    r = JSON.parse(makeRequest(getSecrets().URL, params));
    result.transactions = result.transactions.concat(r.transactions);
  }

  // Replace the dates with JavaScript dates
  for (let i = 0; i < result.transactions.length; i++) {
    result.transactions[i].date = Date.parse(result.transactions[i].date);
  }

  Logger.log(`We downloaded ${result.transactions.length} transactions from Plaid.`);
  return result;

}


/**
 * Fetch the transactions that are currently on the sheet.
 * 
 * @param {SpreadsheetApp.Sheet} sheet the sheet to fetch the transactions from.
 * @return {Object} the transactions.
 */
function getTransactionsFromSheet(sheet) {

  const result = {};
  result.transactions = [];
  result.available = 0.0;
  result.current = 0.0;

  // Get the headers
  result.headers = sheet.getRange(7, 1, 1, sheet.getLastColumn()).getValues().flat();
  result.headers = result.headers.map(item => item.replace("?", ""));
  result.headers = result.headers.map(item => item.toLowerCase());

  // Don't bother if it's empty
  if (sheet.getLastRow() === 7) {
    Logger.log(`We fetched ${result.transactions.length} transactions from the sheet named ${sheet.getName()}.`);
    return result;
  }

  // Get the transactions, starting with most recent
  const values = sheet.getRange(8, 1, sheet.getLastRow() - 7, sheet.getLastColumn()).getValues();
  for (let i = 0; i < values.length; i++) {
    const newTransaction = {}
    for (let j = 0; j < result.headers.length; j++) {
      newTransaction[result.headers[j].toLowerCase()] = values[i][j];
    }
    result.transactions.push(newTransaction);

    // Increment the balance(s)
    result.current += Number(values[i][6]);
    if (values[i][7] === false) {
      result.available += Number(values[i][6]);
    }

  }

  Logger.log(`We fetched ${result.transactions.length} transactions from the sheet named ${sheet.getName()}.`);

  return result;


}


/**
 * Convert a Plaid transaction to a transaction for the sheet.
 * 
 * @param {Object} transaction the transaction to convert.
 * @param {Object} existing the existing transaction to update.
 * @return {Object} the converted transaction.
 */
function plaidToSheet(transaction, existing = undefined) {

  // Use existing values if we have them
  let internal;
  let notes;
  let category;
  let subcategory;
  let channel;
  if (existing === undefined) {
    internal = false;
    notes = "";
    category = transaction.category[0];
    subcategory = "";
    for (const subcat of transaction.category.slice(1)) subcategory = subcategory + subcat + " ";
    subcategory = subcategory.slice(0, -1);
    channel = transaction.payment_channel;

  } else {
    internal = existing.internal;
    notes = existing.notes;
    category = existing.category;
    subcategory = existing.subcategory;
    channel = existing.channel;
  }

  // Return the transaction for the sheet
  return {
    "id": transaction.transaction_id,
    "date": transaction.date,
    "name": transaction.name,
    "category": category,
    "subcategory": subcategory,
    "channel": transaction.payment_channel,
    "amount": -transaction.amount,
    "pending": transaction.pending,
    "internal": internal,
    "notes": notes
  }

}


/** 
 * Searches transactions for the transaction with the ID, and returns its index.
 * Painfully inefficient.
 * 
 * @param {Object[]} transactions the transactions to search.
 * @param {string} the ID to search for.
 * @return {Number} the index of the transaction, or -1 if it doesn't exist.
*/
function getExisitingIndexById(transactions, id) {

  for (let i = 0; i < transactions.length; i++) {
    if (transactions[i].id === id) {
      return i;
    }
  }
  return -1
}


/** 
 * Searches transactions for the transaction with the ID, and returns its index.
 * Painfully inefficient.
 * 
 * @param {Object[]} transactions the transactions to search.
 * @param {string} the ID to search for.
 * @return {Number} the index of the transaction, or -1 if it doesn't exist.
*/
function getPlaidIndexById(transactions, id) {

  for (let i = 0; i < transactions.length; i++) {
    if (transactions[i].transaction_id === id) {
      return i;
    } else if (transactions[i].pending_transaction_id === id) {
      return i;
    }
  }
  return -1
}


/**
 * Inserts the transaction into transactions in the correct place.
 * 
 * @param {Object[]} transactions the list of transactions.
 * @param {Object} transaction the transaction to insert.
 * @return {Object[]} the updated transactions.
 */
function insertNewTransaction(transactions, transaction) {

  // Insert it when we first encounter an exisiting one with a smaller date
  for (let i = 0; i < transactions.length; i++) {
    if (transaction.date >= transactions[i].date) {
      transactions.splice(i, 0, transaction);
      return transactions;
    }
  }

  // If the new transaction is the oldest then add it at the end
  transactions.push(transaction);
  return transactions;

}


/**
 * Writes the transactions to the sheet.
 * 
 * @param {SpreadsheetApp.Sheet} sheet the sheet to write the transactions to.
 * @param {Object[]} transaction the transactions to write.
 * @param {string[]} headers the headers of the sheet.
 */
function writeTransactionsToSheet(sheet, transactions, headers) {

  const result = []
  for (let i = 0; i < transactions.length; i++) {

    const row = headers.slice();
    for (const [key, value] of Object.entries(transactions[i])) {
      if (key === "date") {
        let date = new Date();
        date.setTime(value);
        row[row.indexOf(key)] = date;
      } else {
        row[row.indexOf(key)] = value;
      }
    }
    result.push(row);

  }

  sheet.deleteRows(9, sheet.getLastRow()-8);
  sheet.insertRowsAfter(8, result.length-1);
  sheet.getRange(8, 1, result.length, sheet.getLastColumn()).setValues(result);

}


/**
 * Formats the date as a nice string.
 * 
 * @param {Date} date the date to parse.
 * @return {string} the nicely formatted date.
 */
function formatDate(date) {

  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const days = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  return `${days[date.getDay()]} ${date.getDate()} ${months[date.getMonth()]} ${date.getFullYear()}`;

}


function updateTransactions() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions");

  const existing = getTransactionsFromSheet(sheet);
  const result = downloadAllTransactions();
  // Logger.log(JSON.stringify(existing));
  // Logger.log(JSON.stringify(result));

  // Prepare to determine changes
  const changes = {
    "added": [],
    "removed": []
  }

  for (let i = 0; i < result.transactions.length; i++) {

    let existingTransaction = undefined;
    let existingIndex;

    // If it has a pending ID then see if we have that
    if (result.transactions[i].pending_transaction_id !== null) {
      existingIndex = getExisitingIndexById(existing.transactions, result.transactions[i].pending_transaction_id);
      if (existingIndex >= 0) {
        existingTransaction = existing.transactions[existingIndex]

        // If it has a pending ID but we don't have it from when it was pending
      } else {
        existingIndex = getExisitingIndexById(existing.transactions, result.transactions[i].transaction_id);

        // If a transaction with that transaction_id already exists
        if (existingIndex >= 0) {
          existingTransaction = existing.transactions[existingIndex]
        }
      }

      // If it doesn't have a pending ID
    } else {
      existingIndex = getExisitingIndexById(existing.transactions, result.transactions[i].transaction_id);

      // If a transaction with that transaction_id already exists
      if (existingIndex >= 0) {
        existingTransaction = existing.transactions[existingIndex]
      }
    }

    // Update existing with the transaction
    const newTransaction = plaidToSheet(result.transactions[i], existingTransaction);
    if (existingIndex >= 0) {
      existing.transactions[existingIndex] = newTransaction;
    } else {
      existing.transactions = insertNewTransaction(existing.transactions, newTransaction);
      changes.added.push(newTransaction);
    }

  }
  Logger.log("Finished iterating through Plaid transactions.");

  // Find which old transactions have been removed
  for (const transaction of existing.transactions) {
    if (getPlaidIndexById(result.transactions, transaction.id) == -1) {
      existing.transactions.splice(existing.transactions.indexOf(transaction), 1);
      changes.removed.push(transaction);
    }
  }

  if (changes.added.length == 0 && changes.removed.length == 0) {
    Logger.log("No transactions were added or removed.");

    // Tell the user that there were no new transactions
    // An error is raised if this is called by the trigger
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast("No new changes to the transactions were found.");
    } catch (error) {

    }
  } else {

    // Write the transactions to the sheet
    Logger.log(`There are ${existing.transactions.length} transactions to write.`);
    writeTransactionsToSheet(sheet, existing.transactions, existing.headers);
    Logger.log(`Finished writing transactions to the sheet named ${sheet.getName()}.`)

    // Format the sheet neatly
    formatNeatlyTransactions();
    Logger.log(`Finished formatting the sheet named ${sheet.getName()} neatly.`);

    // Produce a message to tell the user of the changes
    // An error is raised if this is called by the trigger
    try {
      const ui = SpreadsheetApp.getUi();
      let message = "";
      if (changes.added.length > 0) {
        for (const transaction of changes.added) {
          let date = new Date();
          date.setTime(transaction.date);
          message = `${message}ADDED: £${transaction.amount} on ${formatDate(date)} from ${transaction.name}.\r\n`
        }
      }
      if (changes.removed.length > 0) {
        for (const transaction of changes.removed) {
          let date = new Date();
          date.setTime(transaction.date);
          message = `${message}REMOVED: £${transaction.amount} on ${formatDate(date)} from ${transaction.name}.\r\n`
        }
      }
      ui.alert(`${changes.added.length} added | ${changes.removed.length} removed`, message, ui.ButtonSet.OK);
    } catch (error) {

    }

  }

  // Update when this script was last run
  const range = sheet.getRange("TransactionsScriptLastRun");
  if (range !== undefined) {
    const date = new Date();
    let minutes = date.getMinutes().toString();
    if (parseInt(minutes) < 10) minutes = "0" + minutes;
    const dateString = `Last updated on ${formatDate(date)} at ${date.getHours()}:${minutes}.`;
    range.setValue(dateString);
  }

}


/** 
 * Formats the 'Transactions' sheet neatly.
*/
function formatNeatlyTransactions() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions");

  // Get the headers
  let headers = sheet.getRange(7, 1, 1, sheet.getLastColumn()).getValues().flat();
  headers = headers.map(item => item.replace("?", ""));
  headers = headers.map(item => item.toLowerCase());

  // Get column letters (for A1 notation)
  const amountColNum = headers.indexOf("amount") + 1;

  // Create named ranges
  for (let i = 0; i < headers.length; i++) {
    const range = sheet.getRange(8, i + 1, sheet.getLastRow() - 7, 1);
    SpreadsheetApp.getActiveSpreadsheet().setNamedRange(`${headers[i]}s`, range)
  }

  // Add the total titles, merge them, and hide the currently unused rows
  sheet.getRange(6, 2, 1, amountColNum - 2).setValue("AVAILABLE BALANCE");
  sheet.getRange(5, 2, 1, amountColNum - 2).setValue("AMOUNT PENDING");
  sheet.getRange(4, 2, 1, amountColNum - 2).setValue("CURRENT BALANCE");
  sheet.getRange(1, 2, 6, amountColNum - 2).mergeAcross();
  sheet.hideRows(1, 3);

  // Add the totals themselves
  sheet.getRange(6, amountColNum).setValue(`=SUM(amounts)`);
  sheet.getRange(5, amountColNum).setValue(`=SUMIF(pendings, "=TRUE", amounts)`);
  sheet.getRange(4, amountColNum).setValue(`=SUMIF(pendings, "=FALSE", amounts)`);

  // Convert the TRUE/FALSE columns to checkboxes
  sheet.getRange(`pendings`).insertCheckboxes();
  sheet.getRange(`internals`).insertCheckboxes();

  // Add conditional formatting to the amount column
  const amountRange = sheet.getRange(`amounts`);
  const positiveRule = SpreadsheetApp.newConditionalFormatRule().setFontColor("#1B5E20").whenNumberGreaterThan(0).setRanges([amountRange]).build();
  const negativeRule = SpreadsheetApp.newConditionalFormatRule().setFontColor("#B71C1C").whenNumberLessThan(0).setRanges([amountRange]).build();
  sheet.setConditionalFormatRules([positiveRule, negativeRule]);

  // Add data validation for the categories, subcategories, and channels
  let range = sheet.getRange("categorys");
  let values = sheet.getRange("Categories")
  let rule = SpreadsheetApp.newDataValidation().requireValueInRange(values, true).setAllowInvalid(false).build();
  range.setDataValidation(rule);

  range = sheet.getRange("subcategorys");
  values = sheet.getRange("Subcategories")
  rule = SpreadsheetApp.newDataValidation().requireValueInRange(values, true).setAllowInvalid(false).build();
  range.setDataValidation(rule);

  range = sheet.getRange("channels");
  values = sheet.getRange("ChannelsValues")
  rule = SpreadsheetApp.newDataValidation().requireValueInRange(values, true).setAllowInvalid(false).build();
  range.setDataValidation(rule);

  // Freeze the top rows and hide the first column
  sheet.setFrozenRows(7);
  sheet.hideColumn(sheet.getRange("A1"));

  // Add protection for ranges that shouldn't be edited
  for (const protection of sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)) protection.remove();
  for (const name of ["ids", "dates", "names", "amounts", "pendings"]) {
    sheet.getRange(name).protect().setWarningOnly(true);
  }

  // Recreate the filter
  amountRange.getFilter().remove();
  sheet.getRange(7, 1, sheet.getLastRow() - 6, sheet.getLastColumn()).createFilter();
}


/** 
 * Formats the 'Weekly Summary' sheet neatly.
*/
function formatNeatlyWeeklySummary() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Weekly Summary");

  // Hide rows in the future
  sheet.showRows(1, sheet.getLastRow());
  const now = new Date();
  for (let i = 3; i < sheet.getLastRow() - 1; i++) {
    if (sheet.getRange(i, 2).getValue().getTime() <= now.getTime()) {
      sheet.hideRows(3, i - 3)
      break;
    }
  }

}


/**
 * Runs all the formatNeatly functions.
 */
function formatAll() {
  formatNeatlyTransactions()
  formatNeatlyWeeklySummary()
}


/**
 * Updates transactions and then formats everything neatly.
 */
function doEverything() {
  updateTransactions()
  formatNeatlyWeeklySummary()
}


/**
 * Adds the Scripts menu to the menu bar at the top.
 */
function onOpen() {
  const menu = SpreadsheetApp.getUi().createMenu("Scripts");
  menu.addItem("Update Transactions", "updateTransactions");
  menu.addItem("Format the Transactions sheet neatly", "formatNeatlyTransactions");
  menu.addItem("Format the Weekly Summary sheet neatly", "formatNeatlyWeeklySummary");
  menu.addSeparator();
  menu.addItem("Format all sheets neatly", "formatAll");
  menu.addSeparator();
  menu.addItem("Do everything", "doEverything");
  menu.addToUi();
}
