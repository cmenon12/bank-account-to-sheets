/*
  =============================================================================
  Project Page: https://github.com/cmenon12/bank-account-to-sheets
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
    throw Error(`There was a ${status} error fetching ${url}.`);
  }

}


/**
 * Downloads and returns all transactions from Plaid.
 * 
 * @return {Object} the result of transactions.get, with all transactions.
 */
function downloadAllTransactionsFromPlaid() {

  /*// Force Plaid to refresh the transactions
  let params = {
    "contentType": "application/json",
    "method": "post",
    "payload": JSON.stringify({
      "client_id": getSecrets().CLIENT_ID,
      "secret": getSecrets().SECRET,
      "access_token": getSecrets().ACCESS_TOKEN
    }),
    "muteHttpExceptions": true
  };
  makeRequest(`${getSecrets().URL}/transactions/refresh`, params);
*/

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
  params = {
    "contentType": "application/json",
    "method": "post",
    "payload": JSON.stringify(body),
    "muteHttpExceptions": true
  };

  // Make the first POST request
  const result = JSON.parse(makeRequest(`${getSecrets().URL}/transactions/get`, params));
  const total_count = result.total_transactions;
  let offset = 0;
  let r;

  Logger.log(`There are ${total_count} transactions in Plaid.`);

  // Make repeated requests
  while (offset <= total_count - 1) {
    offset = offset + 500;
    body.options.offset = offset;
    params.payload = JSON.stringify(body);
    r = JSON.parse(makeRequest(`${getSecrets().URL}/transactions/get`, params));
    result.transactions = result.transactions.concat(r.transactions);
  }

  // Replace the dates with JavaScript dates
  for (const plaidTxn of result.transactions) plaidTxn.date = Date.parse(plaidTxn.date);

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
  result.headers = sheet.getRange(getHeaderRowNumber(sheet), 1, 1, sheet.getLastColumn()).getValues().flat();
  result.headers = result.headers.map(item => item.replace("?", ""));
  result.headers = result.headers.map(item => item.toLowerCase());

  // Don't bother if it's empty
  if (sheet.getLastRow() === getHeaderRowNumber(sheet)) {
    Logger.log(`We fetched ${result.transactions.length} transactions from the sheet named ${sheet.getName()}.`);
    return result;
  }

  // Get the transactions, starting with most recent
  const values = sheet.getRange(getHeaderRowNumber(sheet) + 1, 1, sheet.getLastRow() - getHeaderRowNumber(sheet), sheet.getLastColumn()).getValues();
  for (let i = 0; i < values.length; i++) {
    const newSheetTxn = {};
    for (let j = 0; j < result.headers.length; j++) {
      newSheetTxn[result.headers[j].toLowerCase()] = values[i][j];
    }
    if (typeof newSheetTxn.date === "number") {
      newSheetTxn.date = new Date(newSheetTxn.date)
    }
    result.transactions.push(newSheetTxn);

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
 * @param {Object} plaidTxn the transaction to convert.
 * @param {Object} sheetTxn the existing sheet transaction to update.
 * @return {Object} the converted transaction.
 */
function plaidToSheet(plaidTxn, sheetTxn = undefined) {

  // Use existing values if we have them
  let internal;
  let notes;
  let category;
  let subcategory;
  let channel;
  if (sheetTxn === undefined) {
    internal = false;
    notes = "";
    category = plaidTxn.category[0];
    subcategory = "";
    for (const subcat of plaidTxn.category.slice(1)) subcategory = subcategory + subcat + " ";
    subcategory = subcategory.slice(0, -1);
    channel = plaidTxn.payment_channel;

  } else {
    internal = sheetTxn.internal;
    notes = sheetTxn.notes;
    category = sheetTxn.category;
    subcategory = sheetTxn.subcategory;
    channel = sheetTxn.channel;
  }

  // Return the transaction for the sheet
  return {
    "id": plaidTxn.transaction_id,
    "date": plaidTxn.date,
    "name": plaidTxn.name,
    "category": category,
    "subcategory": subcategory,
    "channel": channel,
    "amount": -plaidTxn.amount,
    "pending": plaidTxn.pending,
    "internal": internal,
    "notes": notes
  };

}


/**
 * Searches the transactions from the sheet to see if a given Plaid transaction already exists.
 * Painfully inefficient.
 *
 * @param {Object[]} sheetTxns the sheet transactions to search.
 * @param {Object} plaidTxn the Plaid transaction to search for.
 * @return {Number} the index of the plaidTxn, or -1 if it doesn't exist.
 */
function getIndexOfPlaidFromSheet(sheetTxns, plaidTxn) {

  const sameDateAndAmount = [];

  for (let i = 0; i < sheetTxns.length; i++) {

    // Check the IDs
    if (sheetTxns[i].id === plaidTxn.pending_transaction_id) {
      return i;
    } else if (sheetTxns[i].id === plaidTxn.transaction_id) {
      return i;
    }


    /* Only enable when the ACCESS_TOKEN has been changed
    // Check the date, name, and amount
    let date = sheetTxns[i].date
    if (typeof date === "number") {
      date = new Date(date)
    }
    if (date.getTime() === plaidTxn.date &&
      sheetTxns[i].name === plaidTxn.name &&
      sheetTxns[i].amount === -plaidTxn.amount) {
      return i;
    }

    // For if the name has changed
    if (date.getTime() === plaidTxn.date &&
      sheetTxns[i].amount === -plaidTxn.amount) {
      sameDateAndAmount.push(i)
    }
    */
  }

  // If there was only one with that date and amount
  if (sameDateAndAmount.length === 1) {
    return sameDateAndAmount[0];
  }

  return -1;
}


/**
 * Searches the transactions from plaid for the transaction with the ID, and returns its index.
 * Painfully inefficient.
 *
 * @param {Object[]} plaidTxns the Plaid transactions to search.
 * @param {string} id ID to search for.
 * @return {Number} the index of the transaction, or -1 if it doesn't exist.
 */
function getIndexOfIdFromPlaid(plaidTxns, id) {

  for (let i = 0; i < plaidTxns.length; i++) {
    if (plaidTxns[i].transaction_id === id) {
      return i;
    } else if (plaidTxns[i].pending_transaction_id === id) {
      return i;
    }
  }
  return -1;
}


/**
 * Inserts the sheet transaction into the sheet transactions in the correct place.
 *
 * @param {Object[]} sheetTxns the list of transactions from the sheet.
 * @param {Object} sheetTxn the sheet transaction to insert.
 * @return {Object[]} the updated sheet transactions.
 */
function saveNewSheetTransaction(sheetTxns, sheetTxn) {

  // Insert it when we first encounter an existing one with a smaller date
  for (let i = 0; i < sheetTxns.length; i++) {
    if (sheetTxn.date >= sheetTxns[i].date) {
      sheetTxns.splice(i, 0, sheetTxn);
      return sheetTxns;
    }
  }

  // If the new transaction is the oldest then add it at the end
  sheetTxns.push(sheetTxn);
  return sheetTxns;

}


/**
 * Writes the sheet transactions to the sheet.
 *
 * @param {SpreadsheetApp.Sheet} sheet the sheet to write the transactions to.
 * @param {Object[]} sheetTxns the sheet transactions to write.
 * @param {string[]} headers the headers of the sheet.
 */
function writeTransactionsToSheet(sheet, sheetTxns, headers) {

  const result = [];
  for (let i = 0; i < sheetTxns.length; i++) {

    const row = headers.slice();
    for (const [key, value] of Object.entries(sheetTxns[i])) {
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

  sheet.deleteRows(getHeaderRowNumber(sheet) + 2, sheet.getLastRow() - (getHeaderRowNumber(sheet) + 1));
  sheet.insertRowsAfter(getHeaderRowNumber(sheet) + 1, result.length - 1);
  sheet.getRange(getHeaderRowNumber(sheet) + 1, 1, result.length, sheet.getLastColumn()).setValues(result);

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


/**
 * Updates the transactions in the Transactions sheet.
 */
function updateTransactions() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions");

  const existing = getTransactionsFromSheet(sheet);
  const plaid = downloadAllTransactionsFromPlaid();

  // Prepare to determine changes
  const changes = {
    "added": [],
    "removed": []
  };

  for (let i = 0; i < plaid.transactions.length; i++) {

    let existingTxn = undefined;
    let existingIndex;

    // Search for it in existing
    existingIndex = getIndexOfPlaidFromSheet(existing.transactions, plaid.transactions[i]);
    if (existingIndex >= 0) {
      existingTxn = existing.transactions[existingIndex]
    }

    // Update existing with the transaction
    const newSheetTxn = plaidToSheet(plaid.transactions[i], existingTxn);
    if (existingIndex >= 0) {
      existing.transactions[existingIndex] = newSheetTxn;
    } else {
      existing.transactions = saveNewSheetTransaction(existing.transactions, newSheetTxn);
      changes.added.push(newSheetTxn);
    }

  }
  Logger.log("Finished iterating through Plaid transactions.");

  // Find which old transactions have been removed
  for (const sheetTxn of existing.transactions) {
    if (getIndexOfIdFromPlaid(plaid.transactions, sheetTxn.id) === -1) {
      existing.transactions.splice(existing.transactions.indexOf(sheetTxn), 1);
      changes.removed.push(sheetTxn);
    }
  }

  if (changes.added.length === 0 && changes.removed.length === 0) {
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
    formatNeatlyTransactions(plaid);
    Logger.log(`Finished formatting the sheet named ${sheet.getName()} neatly.`);

    // Produce a message to tell the user of the changes
    // An error is raised if this is called by the trigger
    try {
      const ui = SpreadsheetApp.getUi();
      let message = "";
      if (changes.added.length > 0) {
        for (const sheetTxn of changes.added) {
          let date = new Date();
          date.setTime(sheetTxn.date);
          message = `${message}ADDED: £${sheetTxn.amount} on ${formatDate(date)} from ${sheetTxn.name}.\r\n`
        }
      }
      if (changes.removed.length > 0) {
        for (const sheetTxn of changes.removed) {
          let date = new Date();
          date.setTime(sheetTxn.date);
          message = `${message}REMOVED: £${sheetTxn.amount} on ${formatDate(date)} from ${sheetTxn.name}.\r\n`
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
 * Extract and return the totals for the given account.
 * 
 * @param {Object} account the account from Plaid.
 * @return {Object} the totals.
 */
function getPlaidAccountTotals(account) {

  const result = {};

  // For a credit card account
  if (account.type === "credit") {
    result.available = account.balances.limit - account.balances.available;
    result.current = -account.balances.current;
    result.pending = result.available - result.current;

    // For a depository (normal current) account
  } else {
    result.available = account.balances.available;
    result.current = account.balances.current;
    result.pending = result.available - result.current;
  }

  return result;

}


/**
 * Formats the 'Transactions' sheet neatly.
 * 
 * @param {Object} plaidResult the result of transactions.get from Plaid.
 */
function formatNeatlyTransactions(plaidResult = undefined) {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions");

  // Get the headers
  let headers = sheet.getRange(getHeaderRowNumber(sheet), 1, 1, sheet.getLastColumn()).getValues().flat();
  headers = headers.map(item => item.replace("?", ""));
  headers = headers.map(item => item.toLowerCase());

  // Get column letters (for A1 notation)
  const amountColNum = headers.indexOf("amount") + 1;

  // Create named ranges
  for (let i = 0; i < headers.length; i++) {
    const range = sheet.getRange(getHeaderRowNumber(sheet) + 1, i + 1, sheet.getLastRow() - getHeaderRowNumber(sheet), 1);
    SpreadsheetApp.getActiveSpreadsheet().setNamedRange(`${headers[i]}s`, range)
  }

  if (plaidResult !== undefined) {
    sheet.deleteRows(1, getHeaderRowNumber() - 1);
    if (plaidResult.accounts.length === 1) {
      sheet.insertRows(1, 2)

      // Add the total titles and merge them
      sheet.getRange(1, 2, 1, amountColNum - 2).setValue("CURRENT BALANCE");
      sheet.getRange(2, 2, 1, amountColNum - 2).setValue("AMOUNT PENDING (UNACCOUNTED FOR)");
      sheet.getRange(3, 2, 1, amountColNum - 2).setValue("AMOUNT PENDING (ACCOUNTED FOR)");
      sheet.getRange(4, 2, 1, amountColNum - 2).setValue("AVAILABLE BALANCE");
      sheet.getRange(1, 2, 4, amountColNum - 2).mergeAcross();

      // Extract the totals
      const totals = getPlaidAccountTotals(plaidResult.accounts[0]);

      // Add the totals themselves
      sheet.getRange(1, amountColNum).setValue(`${totals.current}`);
      sheet.getRange(2, amountColNum).setValue(`=${totals.pending}-SUMIF(pendings, "=TRUE", amounts)`);
      sheet.getRange(3, amountColNum).setValue(`=SUMIF(pendings, "=TRUE", amounts)`);
      sheet.getRange(4, amountColNum).setValue(`=${totals.current}-${totals.pending}`);

    } else {
      sheet.insertRows(1, (plaidResult.accounts.length * 3) + 4);

      // Prepare to track the grand totals
      const grandTotals = {};
      grandTotals.available = 0;
      grandTotals.current = 0;
      grandTotals.pending = 0;

      // For each account
      for (let i = 1; i < plaidResult.accounts.length - 1; i++) {

        // Add the total titles and merge them
        sheet.getRange((i * 3) - 2, 2, 1, amountColNum - 2).setValue(`${plaidResult.accounts[i - 1].name} CURRENT BALANCE`);
        sheet.getRange((i * 3) - 1, 2, 1, amountColNum - 2).setValue(`${plaidResult.accounts[i - 1].name} AMOUNT PENDING (ACCOUNTED FOR)`);
        sheet.getRange(i * 3, 2, 1, amountColNum - 2).setValue(`${plaidResult.accounts[i - 1].name} AVAILABLE BALANCE`);
        sheet.getRange((i * 3) - 2, 2, 3, amountColNum - 2).mergeAcross();

        // Extract the totals
        const totals = getPlaidAccountTotals(plaidResult.accounts[i - 1]);
        grandTotals.available = totals.curravailableent + grandTotals.available;
        grandTotals.current = totals.current + grandTotals.current;
        grandTotals.pending = totals.pending + grandTotals.pending;

        // Add the totals themselves
        sheet.getRange((i * 3) - 2, amountColNum).setValue(`${totals.current}`);
        sheet.getRange((i * 3) - 1, amountColNum).setValue(`=SUMIFS(amounts, pendings, "=TRUE", accounts, "${plaidResult.accounts[i - 1].name}")`);
        sheet.getRange(i * 3, amountColNum).setValue(`=${totals.current}-SUMIFS(amounts, pendings, "=TRUE", accounts, "${plaidResult.accounts[i - 1].name}")`);

      }

      const startingRow = (plaidResult.accounts.length * 3) + 1;

      // Add the total titles and merge them
      sheet.getRange(startingRow, 2, 1, amountColNum - 2).setValue("TOTAL CURRENT BALANCE");
      sheet.getRange(startingRow + 1, 2, 1, amountColNum - 2).setValue("TOTAL AMOUNT PENDING (UNACCOUNTED FOR)");
      sheet.getRange(startingRow + 2, 2, 1, amountColNum - 2).setValue("TOTAL AMOUNT PENDING (ACCOUNTED FOR)");
      sheet.getRange(startingRow + 3, 2, 1, amountColNum - 2).setValue("TOTAL AVAILABLE BALANCE");
      sheet.getRange(1, 2, 4, amountColNum - 2).mergeAcross();

      // Add the totals themselves
      sheet.getRange(1, amountColNum).setValue(`${grandTotals.current}`);
      sheet.getRange(2, amountColNum).setValue(`=${grandTotals.pending}-SUMIF(pendings, "=TRUE", amounts)`);
      sheet.getRange(3, amountColNum).setValue(`=SUMIF(pendings, "=TRUE", amounts)`);
      sheet.getRange(4, amountColNum).setValue(`=${grandTotals.current}-${grandTotals.pending}`);

    }
  }

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
  sheet.setFrozenRows(getHeaderRowNumber(sheet));
  sheet.hideColumn(sheet.getRange("A1"));

  // Add protection for ranges that shouldn't be edited
  for (const protection of sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)) protection.remove();
  for (const name of ["ids", "dates", "names", "amounts", "pendings"]) {
    sheet.getRange(name).protect().setWarningOnly(true);
  }

  // Recreate the filter
  amountRange.getFilter().remove();
  sheet.getRange(getHeaderRowNumber(sheet), 1, sheet.getLastRow() - (getHeaderRowNumber(sheet) - 1), sheet.getLastColumn()).createFilter();
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
 * Searches for and returns the row number of the header row.
 * 
 * @param {SpreadsheetApp.Sheet} sheet the sheet to search.
 * @return {number} the row number, or -1 if it can't be found.
*/
function getHeaderRowNumber(sheet) {

  const range = sheet.getRange(1, 1, sheet.getLastRow()).getValues();
  for (let i = 0; i < range.length; i++) {
    if (range[i][0] === "ID") {
      return i + 1;
    }
  }

  return -1;

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
