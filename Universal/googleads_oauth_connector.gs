const DEVELOPER_TOKEN = PropertiesService.getScriptProperties().getProperty('GADS_DEVELOPER_TOKEN');

function getOAuthService() {
  var scriptId = ScriptApp.getScriptId();
  var redirectUri = 'https://script.google.com/macros/d/' + scriptId + '/usercallback';

  return OAuth2.createService('GoogleAds')
      .setAuthorizationBaseUrl('https://accounts.google.com/o/oauth2/auth')
      .setTokenUrl('https://oauth2.googleapis.com/token')
      .setClientId(PropertiesService.getScriptProperties().getProperty('GOOGLE_OAUTH_CLIENT_ID'))
      .setClientSecret(PropertiesService.getScriptProperties().getProperty('GOOGLE_OAUTH_CLIENT_SECRET'))
      .setCallbackFunction('authCallback')
      .setPropertyStore(PropertiesService.getUserProperties())
      .setScope('https://www.googleapis.com/auth/adwords')
      .setParam('access_type', 'offline')
      .setParam('approval_prompt', 'force')
      .setRedirectUri(redirectUri);
}

function authCallback(request) {
  var service = getOAuthService();
  var isAuthorized = service.handleCallback(request);
  if (isAuthorized) {
    var token = service.getAccessToken();
    saveAccessToken(token);
    return HtmlService.createHtmlOutput('Success! You can close this tab.');
  } else {
    return HtmlService.createHtmlOutput('Denied. You can close this tab');
  }
}

function saveAccessToken(token) {
  PropertiesService.getScriptProperties().setProperty('GOOGLE_ADS_ACCESS_TOKEN', token);
}

function getAccessToken() {
  var service = getOAuthService();
  if (service.hasAccess()) {
    var token = service.getAccessToken();
    Logger.log(token);
    saveAccessToken(token);
    return token;
  }
  // return PropertiesService.getScriptProperties().getProperty('GOOGLE_ADS_ACCESS_TOKEN');
}

function showOAuthPopup() {
  // no access token
  var service = getOAuthService();
  if (!service.hasAccess()) {
    var authorizationUrl = service.getAuthorizationUrl();
    var htmlOutput = HtmlService.createHtmlOutput('<div style=width:500px;height:350px;border:solid;border-width:1px;margin:auto;text-align:center;padding-top:20px><h1 style=text-align:center>Authorize Google Ads</h1><a href="' + authorizationUrl + '"onclick=google.script.run.showwaiting() target=_blank><button style="text-align:center;margin-top:50px;padding:15px 40px;background:#fea539;font-weight:700;font-size:16px;cursor:pointer">Authorize</button></a></div>')
      .setWidth(550)
      .setHeight(400);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Authorize Google Ads');
  } else {
    //check saved account
    if (PropertiesService.getScriptProperties().getProperty('customer_id') != '-') {
      var htmlOutput = HtmlService.createHtmlOutput('<div style=width:500px;height:350px;border:solid;border-width:1px;margin:auto;text-align:center;padding-top:20px><h1 style=text-align:center>Account Selected</h1><h2>' + PropertiesService.getScriptProperties().getProperty('customer_id') + '-' + PropertiesService.getScriptProperties().getProperty('customer_name') + '</h2><button style="text-align:center;margin-top:50px;padding:15px 40px;background:#fea539;font-weight:700;font-size:16px;cursor:pointer" onclick="google.script.run.showaccountslist()">Change Account</button> <button style="text-align:center;margin-top:50px;margin-left:10px;padding:15px 40px;background:#fff;font-weight:700;font-size:16px;cursor:pointer" onclick="google.script.run.revokeAccessToken()">Revoke Access</button></div>')
        .setWidth(550)
        .setHeight(400);
      SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Ads Connector');
    } else {
      //show dropdown
      showaccountslist();
    }
  }

  createTimeDrivenTrigger();
}

function showwaiting() {
  var htmlOutput = HtmlService.createHtmlOutput('<div><div style=width:500px;height:250px;border:solid;border-width:1px;margin:auto;text-align:center;padding-top:20px><h2 style=text-align:center>Waiting for authentication..</h2><svg id=airport-waiting-lounge viewBox="0 0 660.35 376.15"xmlns=http://www.w3.org/2000/svg><g><path d=M1588.54,383.43s36.92,128.16-102.31,195.88c-32.93,16-105.38-17.63-164.38-14.29-111.63,6.32-196.23,44.3-251.21,21.86-61.51-25.11-84.22-49.47-93.55-66.56-2.44-4.46-5.78-9.14-8.33-13.4-15.91-26.51-66.09-124.89-9.66-189.43,48.08-55,142.69-54.73,189.92-50.65a134.77,134.77,0,0,1,48.48,13.73c29.66,14.76,83.83,40.92,106.69,46.71,43.11,10.93,49.71,9,130-25.34S1572.56,269.19,1588.54,383.43Z fill=#e9edfb transform="translate(-932.63 -265.24)"></path></g><g><path d=M1011,478.58S1071,362.32,1140,360.13s68,76.73,87.69,88.79,40.56,8.77,55.91-17.54,32.78-75.23,65.66-73,70.27,104.83,90,104.83,56-78.27,88.46-77.09c33.61,1.22,56.67,92.67,56.67,92.67l-51.95-.24Z fill=#e1e6fb transform="translate(-932.63 -265.24)"></path></g><g><g><path d=M949.62,478.45,1148.77,481s16.57-1,33.57-14c12.94-9.88,17.43-17,17.43-24s-6.48-43-70-50L935.4,389.3s-1.72,6-2.59,17.94a98.62,98.62,0,0,0,.44,16.64,180.16,180.16,0,0,0,3,19.08C939.91,459.72,944.89,470.73,949.62,478.45Z fill=#bac5f4 transform="translate(-932.63 -265.24)"></path><ellipse cx=193.64 cy=154.31 fill=#fff rx=38.5 ry=10.5></ellipse><circle cx=40.64 cy=170.31 fill=#fff r=10.5></circle><circle cx=71.14 cy=170.81 fill=#fff r=11></circle><circle cx=71.14 cy=170.81 fill=#fff r=11></circle><circle cx=103.14 cy=170.81 fill=#fff r=11></circle><circle cx=353.13 cy=118.33 fill=#fff r=15.86></circle><circle cx=336.14 cy=119.52 fill=#fff r=12.67></circle><circle cx=542.42 cy=82.46 fill=#fff r=18.44></circle><circle cx=562.08 cy=84.43 fill=#fff r=11.22></circle><circle cx=523.35 cy=86.32 fill=#fff r=11.22></circle><circle cx=134.14 cy=170.81 fill=#fff r=11></circle><path d=M954,489l628.94.48s-21.61,66.16-36.71,78.89,13.94,73.85-97.58,34.38S1275.55,604,1243,609.11s-118.49,47.11-153.34,26.74S1071.09,591.29,1006,590s-61.57-66.21-61.57-66.21L954,489 fill=#f5f5fc transform="translate(-932.63 -265.24)"></path></g><polygon fill=#fff points="650.3 224.2 21.36 223.76 16.58 213.25 652.95 213.36 650.3 224.2"></polygon><polygon fill=#fff points="110.78 214.48 104.9 214.48 104.83 8.24 110.75 6.95 110.78 214.48"></polygon><polygon fill=#fff points="281.4 216.83 275.52 216.83 275.47 20.54 281.29 23.27 281.4 216.83"></polygon><polygon fill=#fff points="448.5 219.19 442.61 219.19 442.62 60.5 448.55 58.48 448.5 219.19"></polygon><polygon fill=#fff points="607.06 216.83 601.18 216.83 601.23 20.95 607.18 23.42 607.06 216.83"></polygon></g><g><polygon fill=#bac5f4 points="418.41 276.47 433.13 276.47 433.09 265.3 576.28 265.28 576.28 277.06 589.24 277.06 589.24 247.02 418.41 247.61 418.41 276.47"></polygon><polygon fill=#bac5f4 points="418.77 242.36 589.32 241.97 589.32 234.69 418.39 235.07 418.77 242.36"></polygon><polygon fill=#bac5f4 points="418.39 231.62 588.94 231.24 588.94 223.96 418.01 224.34 418.39 231.62"></polygon><polygon fill=#bac5f4 points="418.77 220.51 589.32 220.13 589.32 212.84 418.39 213.23 418.77 220.51"></polygon><polygon fill=#bac5f4 points="418.77 209.01 589.32 208.63 589.32 201.35 418.39 201.73 418.77 209.01"></polygon></g></svg></div></div><script>setInterval(function(){google.script.run.checkaccesstoken()}, 3000);</script>')
    .setHeight(400)
    .setWidth(550);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '.');
}

function checkaccesstoken() {
  if (getAccessToken()) {
    showOAuthPopup();
  }
}

function createTimeDrivenTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  if (triggers.length === 0) {
    ScriptApp.newTrigger('get_current_month')
      .timeBased()
      .everyDays(1)
      .atHour(10)
      .create();
  }
}

let headers = {
  Authorization: "Bearer " + ScriptApp.getOAuthToken(),
  "developer-token": DEVELOPER_TOKEN
};

/*
function showaccountslist() {
  let headers1 = {
    Authorization: "Bearer " + getAccessToken(),
    "developer-token": DEVELOPER_TOKEN
  };

  let requestParams = {
    method: "GET",
    headers: headers1,
    muteHttpExceptions: true
  };
  let response = JSON.parse(UrlFetchApp.fetch("https://googleads.googleapis.com/v22/customers:listAccessibleCustomers", requestParams));

  var customer_list = [];
  for (var i = 0; i < response.resourceNames.length; i++) {
    var customer_id = response.resourceNames[i].toString().split('/')[1];

    var cust_list = getcustomerName(customer_id);
    for (var j = 0; j < cust_list.length; j++) {
      customer_list.push(cust_list[j]);
    }
  }

  var html = '<div style=width:500px;height:350px;border:solid;border-width:1px;margin:auto;text-align:center;padding-top:20px><h1 style=text-align:center>Change Your Account</h1><select id="customer_id" style=width:50%;display:block;margin:auto;padding:10px><option value="0">Select</option>';
  for (var a = 0; a < customer_list.length; a++) {
    var mcc = customer_list[a].mcc;
    var mcc_id;
    if (mcc) {
      mcc_id = ' (MCC - ' + mcc + ')';
    } else {
      mcc_id = '';
    }
    html = html + '<option value="' + customer_list[a].id + '-' + customer_list[a].name + '-' + customer_list[a].mcc + '">' + customer_list[a].name + mcc_id + '</option>';
  }
  var html = html + '</select><br><br><button style="text-align:center; margin-top:50px; padding: 15px 40px; background:#fea539; font-weight:bold; font-size:16px; cursor:pointer;" onclick="getData(); google.script.host.close();">Get Data</button></div><script>function getData(){if(document.getElementById("customer_id").value !=0){google.script.run.getData(document.getElementById("customer_id").value)}}</script>';

  var htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(550)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Select Account');
}
*/
function showaccountslist() {
  const API_VERSION = 'v22'; // v22/v23 will cause 404 errors

  let headers1 = {
    Authorization: "Bearer " + getAccessToken(),
    "developer-token": DEVELOPER_TOKEN
  };

  let requestParams = {
    method: "GET",
    headers: headers1,
    muteHttpExceptions: true
  };

  // 1. Fetch the data
  let fetchUrl = `https://googleads.googleapis.com/${API_VERSION}/customers:listAccessibleCustomers`;
  let responseRaw = UrlFetchApp.fetch(fetchUrl, requestParams);
  let response = JSON.parse(responseRaw.getContentText());

  var customer_list = [];

  // 2. USE THE RESPONSE: Loop through and populate customer_list
  if (response.resourceNames && response.resourceNames.length > 0) {
    for (var i = 0; i < response.resourceNames.length; i++) {
      // resourceNames looks like "customers/1234567890"
      var customer_id = response.resourceNames[i].split('/')[1];

      try {
        // Fetch the details (Name/MCC status) for each ID
        var cust_details = getcustomerName(customer_id);
        if (cust_details && cust_details.length > 0) {
          customer_list = customer_list.concat(cust_details);
        }
      } catch (e) {
        Logger.log("Error processing customer " + customer_id + ": " + e.message);
      }
    }
  } else {
    Logger.log("No accounts found in response: " + responseRaw.getContentText());
  }

  // 3. Build the HTML UI
  var html = `
    <div style="width:500px; height:350px; border:solid 1px; margin:auto; text-align:center; padding-top:20px; font-family: sans-serif;">
      <h1 style="text-align:center">Change Your Account</h1>
      <select id="customer_id" style="width:80%; display:block; margin:auto; padding:10px;">
        <option value="0">Select an Account</option>`;

  // This loop turns the customer_list into the dropdown options
  for (var a = 0; a < customer_list.length; a++) {
    var item = customer_list[a];
    var mcc_label = item.mcc ? ` (MCC - ${item.mcc})` : '';

    // escaping quotes for safety
    var safeValue = `${item.id}-${item.name}-${item.mcc}`.replace(/"/g, '&quot;');
    var safeName = `${item.name}${mcc_label}`.replace(/"/g, '&quot;');

    html += `<option value="${safeValue}">${safeName}</option>`;
  }

  html += `
      </select>
      <br><br>
      <button style="text-align:center; margin-top:50px; padding:15px 40px; background:#fea539; font-weight:bold; font-size:16px; cursor:pointer; border:none; border-radius:5px;" 
              onclick="sendData()">Set Account</button>
    </div>
    <script>
      function sendData(){
        var val = document.getElementById("customer_id").value;
        if(val !== "0"){
          google.script.run.withSuccessHandler(function() {
             google.script.host.close();
          }).getData(val);
        } else {
          alert("Please select an account first.");
        }
      }
    </script>`;

  var htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(550)
    .setHeight(450);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Select Account');
}

const formatDate = (date) => {
  let d = new Date(date);
  let month = (d.getMonth() + 1).toString().padStart(2, '0');
  let day = d.getDate().toString().padStart(2, '0');
  let year = d.getFullYear();
  return [day, month, year].join('-');
};

function getcustomerName(id) {
  let headers = {
    Authorization: "Bearer " + getAccessToken(),
    "developer-token": DEVELOPER_TOKEN
  };

  let requestParams = {
    method: "POST",
    contentType: "application/json",
    headers: headers,
    muteHttpExceptions: true,
    payload: JSON.stringify({
      query: `SELECT customer.descriptive_name, customer.id, customer.manager
FROM customer`
    })
  };

  let response = JSON.parse(UrlFetchApp.fetch("https://googleads.googleapis.com/v23/customers/" + id + "/googleAds:searchStream", requestParams));

  if (response[0].results[0].customer.manager == true) {
    let headers = {
      Authorization: "Bearer " + getAccessToken(),
      "developer-token": DEVELOPER_TOKEN
      //"login-customer-id": PARENT_MCC_ID
    };

    let requestParams = {
      method: "POST",
      contentType: "application/json",
      headers: headers,
      muteHttpExceptions: true,
      payload: JSON.stringify({
        query: `SELECT
            customer_client.client_customer,
            customer_client.id,
            customer_client.level,
            customer_client.descriptive_name
        FROM
            customer_client`
      })
    };
    let response = JSON.parse(UrlFetchApp.fetch("https://googleads.googleapis.com/v23/customers/" + id + "/googleAds:searchStream", requestParams));

    var customer = [];
    for (var i = 1; i < response[0].results.length; i++) {
      customer.push({ id: response[0].results[i].customerClient.id, name: response[0].results[i].customerClient.descriptiveName, mcc: id });
    }

    return customer;
  } else {
    var customer = [{ id: response[0].results[0].customer.id, name: response[0].results[0].customer.descriptiveName, mcc: undefined }];
    return customer;
  }
}

function getData(x) {
  //if (PropertiesService.getScriptProperties().getProperty('customer_id') == '-') {
  var cust_id = x.split('-')[0];
  var cust_name = x.split('-')[1];
  var mcc_id = x.split('-')[2];

  PropertiesService.getScriptProperties().setProperty('customer_id', cust_id);
  PropertiesService.getScriptProperties().setProperty('customer_name', cust_name);
  PropertiesService.getScriptProperties().setProperty('mcc_id', mcc_id);
}

function revokeAccessToken1(accessToken) {
  var url = 'https://oauth2.googleapis.com/revoke';
  var payload = 'token=' + accessToken;

  var options = {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: payload
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    Logger.log(response.getContentText());

    if (response.getResponseCode() === 200) {
      Logger.log('Access token successfully revoked.');
    } else {
      Logger.log('Failed to revoke access token. Response code: ' + response.getResponseCode());
    }
  } catch (e) {
    Logger.log('Error: ' + e.message);
  }
}

function revokeAccessToken() {
  var service = getOAuthService();
  var accessToken = service.getAccessToken();
  revokeAccessToken1(accessToken);
  service.reset();
  showOAuthPopup();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data");
  var sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Main");
  sheet.clear();
  transpose();
  PropertiesService.getScriptProperties().setProperty('customer_id', '-');
  PropertiesService.getScriptProperties().setProperty('customer_name', '-');
  PropertiesService.getScriptProperties().setProperty('mcc_id', '-');
  sheet1.getRange('E4').setValue('');
  sheet1.getRange('G4').setValue('');
  sheet1.getRange('I4').setValue('');
}
