function getSecrets() {

  const secrets = {};
  secrets.URL = "https://development.plaid.com/transactions/get";
  secrets.CLIENT_ID = "your client ID";
  secrets.SECRET = "your secret";
  secrets.ACCESS_TOKEN = "your access token";

  return secrets;

}
