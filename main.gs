// main.gs -- Supports downloading all intraday data
// Currently fetches steps, floors, calories and heart rate
//
// Change activities if you want more stuff
var activities = [
    //"activities/steps",
    //"activities/calories",
    //"activities/floors", // not working
    //"activities/distance",
    "activities/heart", // heart rate must be last
];

// Set the sheet name where data will be downloaded. Nothing else should be in this sheet
var sheetName = "Data";

// If you want want to filter out empty rows from the data, set this to true. If heartrate and steps is zero, the row is considered empty.
var filterEmptyRows = true;


/*
 * Do not change these key names. These are just keys to access these properties once you set them up by running the Setup function from the Fitbit menu
 */
// Key of userProperties for Fitbit consumer key.
var CONSUMER_KEY_PROPERTY_NAME = "fitbitConsumerKey";
// Key of userProperties for Fitbit consumer secret.
var CONSUMER_SECRET_PROPERTY_NAME = "fitbitConsumerSecret";

var SERVICE_IDENTIFIER = 'fitbit';

var userProperties = PropertiesService.getUserProperties();

function onOpen() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var menuEntries = [
        { name: "Setup", functionName: "setup" },
        { name: "Authorize", functionName: "showSidebar" },
        { name: "Reset", functionName: "clearService" },
        { name: "Download data", functionName: "downloadData" }
    ];
    ss.addMenu("Fitbit", menuEntries);
}


function isConfigured() {
    return getConsumerKey() != "" && getConsumerSecret() != "";
}

function setConsumerKey(key) {
    userProperties.setProperty(CONSUMER_KEY_PROPERTY_NAME, key);
}

function getConsumerKey() {
    var key = userProperties.getProperty(CONSUMER_KEY_PROPERTY_NAME);
    if (key == null) {
        key = "";
    }
    return key;
}

function setLoggables(loggable) {
    userProperties.setProperty("loggables", loggable);
}

function getLoggables() {
    var loggable = userProperties.getProperty("loggables");
    if (loggable == null) {
        loggable = LOGGABLES;
    } else {
        loggable = loggable.split(',');
    }
    return loggable;
}

function setConsumerSecret(secret) {
    userProperties.setProperty(CONSUMER_SECRET_PROPERTY_NAME, secret);
}

function getConsumerSecret() {
    var secret = userProperties.getProperty(CONSUMER_SECRET_PROPERTY_NAME);
    if (secret == null) {
        secret = "";
    }
    return secret;
}

// function saveSetup saves the setup params from the UI
function saveSetup(e) {
    setConsumerKey(e.parameter.consumerKey);
    setConsumerSecret(e.parameter.consumerSecret);
    setLoggables(e.parameter.loggables);
    setFirstDate(e.parameter.firstDate);
    var app = UiApp.getActiveApplication();
    app.close();
    return app;
}

function setFirstDate(firstDate) {
    userProperties.setProperty("firstDate", firstDate);
}

function getFirstDate() {
    var firstDate = userProperties.getProperty("firstDate");
    if (firstDate == null) {
        firstDate = "today";
    }
    return firstDate;
}

function getFitbitService() {
    // Create a new service with the given name. The name will be used when
    // persisting the authorized token, so ensure it is unique within the
    // scope of the property store
    Logger.log(PropertiesService.getUserProperties());
    return OAuth2.createService(SERVICE_IDENTIFIER)

        // Set the endpoint URLs, which are the same for all Google services.
        .setAuthorizationBaseUrl('https://www.fitbit.com/oauth2/authorize')
        .setTokenUrl('https://api.fitbit.com/oauth2/token')

        // Set the client ID and secret, from the Google Developers Console.
        .setClientId(getConsumerKey())
        .setClientSecret(getConsumerSecret())

        // Set the name of the callback function in the script referenced
        // above that should be invoked to complete the OAuth flow.
        .setCallbackFunction('authCallback')

        // Set the property store where authorized tokens should be persisted.
        .setPropertyStore(PropertiesService.getUserProperties())

        .setScope('activity profile heartrate nutrition weight')

        .setTokenHeaders({
            'Authorization': 'Basic ' + Utilities.base64Encode(getConsumerKey() + ':' + getConsumerSecret())
        });

}

function clearService() {
    OAuth2.createService(SERVICE_IDENTIFIER)
        .setPropertyStore(PropertiesService.getUserProperties())
        .reset();
}

function showSidebar() {
    var service = getFitbitService();
    if (true) {//!service.hasAccess()) {
        var authorizationUrl = service.getAuthorizationUrl();
        var template = HtmlService.createTemplate(
            '<a href="<?= authorizationUrl ?>" target="_blank">Authorize</a>. ' +
            'Reopen the sidebar when the authorization is complete.');
        template.authorizationUrl = authorizationUrl;
        var page = template.evaluate();
        SpreadsheetApp.getUi().showSidebar(page);
    } else {
        Logger.log("Has access!!!!");
    }
}

function authCallback(request) {
    Logger.log("authcallback");
    var service = getFitbitService();
    var isAuthorized = service.handleCallback(request);
    if (isAuthorized) {
        Logger.log("success");
        return HtmlService.createHtmlOutput('Success! You can close this tab.');
    } else {
        Logger.log("denied");
        return HtmlService.createHtmlOutput('Denied. You can close this tab');
    }
}

function yesterday() {
  var dateString = Utilities.formatDate(new Date(new Date() - (24 * 3600 * 1000)), "GMT-5", "yyyy-MM-dd");
  return dateString;
}

function downloadData() {
  if (!isConfigured()) {
    setup();
    return;
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var lastrow = sheet.getLastRow();
  //sheet.clear();
  //if (lastrow > 3) { sheet.deleteRows(2, lastrow - 2); }
  sheet.setFrozenRows(1);

  var options = {
    headers: {
      "Authorization": "Bearer " + getFitbitService().getAccessToken(),
      "method": "GET"
    }
  };

    var table = {};

    var titleCell = sheet.getRange("a1");
    titleCell.setValue("time");

    var dateString = yesterday();
    for (var activity in activities) {
        var currentActivity = activities[activity];
        if (currentActivity == "activities/steps") {
            var stepsColumn = parseInt(activity) + 1;
        }
        try {
            if (currentActivity == "activities/heart") {
                var heartColumn = parseInt(activity) + 1;
                var result = UrlFetchApp.fetch("https://api.fitbit.com/1/user/-/activities/heart/date/" + dateString + "/1d/1sec.json", options);
            } else {
                var result = UrlFetchApp.fetch("https://api.fitbit.com/1/user/-/" + currentActivity + "/date/" + dateString + "/1d.json", options);
            }
        } catch (exception) {
            Logger.log(exception);
        }
        var o = JSON.parse(result.getContentText());
        console.log(o);

        var title = currentActivity.split("/")[1];
        titleCell.offset(0, 1 + parseInt(activity)).setValue(title);
        var intradays_field = "activities-" + title + "-intraday"
        var row = o[intradays_field]["dataset"];

        Logger.log(row.length);
        for (var j in row) {
            var val = row[j];

            index = val["time"];
            if (table[index] instanceof Array) {} else {
                table[index] = new Array()
            }
            table[index][0] = dateString + " " + val["time"];
            table[index].push(val["value"])

        }
      console.log(table);


    }
  var al = activities.length + 1;
  if (sheet.getMaxColumns() > al) {
    sheet.deleteColumns(al + 1, sheet.getMaxColumns() - al);
  }

    // Pad the array - setValues needs a value in each field
    Object.keys(table).forEach(function(key) {
        var tl = table[key].length
        if (tl < al) {
            table[key].push(0)
        }

    });

    // Convert the object to an array - setValues needs an array
    var tablearray = Object.keys(table).map(function(key) {
        return table[key];
    })
    Logger.log(tablearray);
    
    if (filterEmptyRows) {
      tablearray = tablearray.filter(function(currarr) { 
              return (currarr[heartColumn] > 0 || currarr[stepsColumn] > 0);
       });
    }

    if (tablearray[1]) {
      var range = sheet.getRange(lastrow + 1, 1, tablearray.length, al);
      console.log(range.getNumColumns());
      console.log(range.getNumRows());
      console.log(range.getA1Notation());
      range.setValues(tablearray);
    } else {
       SpreadsheetApp.getUi().alert('No data found for chosen day: ' + dateString);
    }
}
