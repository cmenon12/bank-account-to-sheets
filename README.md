# Bank Account to Google Sheets
This is a Google Apps script that imports transactions from a bank account into Google Sheets via Plaid.


## Description


## Setup
You'll need to sign-up for a developer account on Plaid ([here](https://dashboard.plaid.com/signup)). You'll  be prompted to create a 'team' and then granted immediate access to the 'Sandbox' environment, which uses test data. However, in order to use this app you'll need to request access to the 'Development' environment ([here](https://dashboard.plaid.com/overview/development)), which will allow you to use up to 100 real bank accounts. This will take a day or so to be approved.


### Getting an Access Token
You'll find the client ID and the secret for the development environment on the keys page for your team ([here](https://dashboard.plaid.com/team/keys)). Once you've got access to the 'Development' environment, you'll need to create an access token that grants your 'team' (or app) access to your own bank account. 

The simplest way to do this is by following their Quickstart guide here: [https://plaid.com/docs/quickstart/](https://plaid.com/docs/quickstart/). You'll need to use the 'Development' environment instead of the 'Sandbox', and set `PLAID_PRODUCTS` to `auth,transactions`. I did it without Docker, and whilst it is a bit fiddly, it does work eventually. Once you've granted it access to your own bank account, you'll be presented with an access token that you can use in these scripts.  


### Creating the Google Sheet
These scripts are designed to work with a specific Google Sheets template, which you can make your own copy of here: [https://docs.google.com/spreadsheets/d/1ZkSQwviwqdIXWb2EefQwYWmjq2sUUs6U3EB5LTqgyhU/copy](https://docs.google.com/spreadsheets/d/1ZkSQwviwqdIXWb2EefQwYWmjq2sUUs6U3EB5LTqgyhU/copy). Once you've done that, you can easily add the scripts and start importing your transactions.

1. Make a copy of the Google Sheet above and open it. Don't worry about the example transactions, they'll be removed when you import your own.
2. Go to Tools -> Script Editor, which should open a new Apps Script project.
3. Remove anything that's in the editor already, and copy in the contents of [`code.gs`](/code.gs).
4. At the top-left, click the `+` next to `Files` to create a new script file, and name it `secrets`.
5. Copy the contents of [`secrets.gs`](/secrets.gs) into this new file.
6. Add your own client ID, secret, and access token.
7. Save it.
8. Refresh the Google Sheet. This might close the editor (which is fine).
9. A new menu named `Scripts` should appear in the menu bar. 
10. Clicking the `Update Transactions` option within this menu should download all of your transactions into the 'Transactions' sheet, with a summary automatically generated in the 'Weekly Summary' sheet.

Optionally, you can create a [time-driven trigger](https://developers.google.com/apps-script/guides/triggers/installable#time-driven_triggers) to run the `updateTransactions()` function automatically on a regular basis (e.g., hourly).


## Assumptions & Limitations
- You should only grant Plaid access to one bank account, as the sheet won't separate the transactions by account.
- Plaid only updates its list of transactions (from the bank) [every six hours or so](https://plaid.com/docs/transactions/webhooks/#:~:text=typically%2C%20plaid%20will%20check%20for%20transactions%20once%20every%206%20hours%2C%20but%20may%20check%20less%20frequently%20(such%20as%20once%20every%2024%20hours)%20depending%20on%20factors%20such%20as%20the%20institution%20and%20account%20type.), which means that new transactions might not appear instantly. I've been unable to get access to [/transactions/refresh](https://plaid.com/docs/api/products/#transactionsrefresh) to try to remove this delay.
- The access token doesn't expire, however it might need updating if the user (you) changes their password, or when working with European institutions that comply with PSD2's 90-day consent window. This can be done by re-authorising with the quickstart.


## Important Points to Note About the Template
- This template has been designed to work specifically with this script, so you should avoid modifying it (unless you're willing to modify the script too).
- You can edit a transaction's category, subcategory, channel, internal status, and notes, and these changes will all be preserved when the transactions are updated. Changes to anything else (like the date, or the name) won't be.
- Transactions marked as internal won't be included on the Weekly Summary sheet (except as part of the ending balance). This is designed for transactions that are transfers to or from your own accounts, and therefore don't represent money gained or spent by you.  
- The option to format the Weekly Summary sheet neatly will adjust the hidden rows so that the current week is displayed at the top, and weeks in the future are hidden. 
- You can change the categories on the Weekly Summary sheet to any category or subcategory that you want.
  - Note that `Other Shops` isn't a category, but instead is the sum of everything categorised as `Shops`, minus those with the subcategories of `Supermarkets and Groceries` or `Clothing and Accessories`.
  - The final categories column, `Other`, is simply the sum of all transactions, minus those in the displayed categories.
- The Values sheet simply has a static list of all the available categories and subcategories, as well as list of which category each subcategory maps to.
  - On the Transactions sheet, you don't have to enforce that a transaction's subcategory must come from its category, but it's probably sensible to.


## License
[GNU GPLv3](https://choosealicense.com/licenses/gpl-3.0/)
