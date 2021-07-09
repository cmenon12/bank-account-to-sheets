/*
  =============================================================================
  Project Page: https://github.com/cmenon12/bank-account-to-sheets
  Copyright:    (c) 2021 by Christopher Menon
  License:      GNU General Public License, version 3 (GPL-3.0)
                http://www.opensource.org/licenses/gpl-3.0.html
  =============================================================================
 */


function getSecrets() {

  const secrets = {};
  secrets.URL = "https://development.plaid.com";
  secrets.CLIENT_ID = "your client ID";
  secrets.SECRET = "your secret";
  secrets.ACCESS_TOKEN = "your access token";

  return secrets;

}
