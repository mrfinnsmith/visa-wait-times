/* Not using this right now. Might be useful to figure out what the correspondence is between cids (the consulate IDs) and the consulates, but there is no documenation. The HTML code works fine.

This code updates the API sheet, which is probably hidden.

function fetchVisaWaitTimes() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  var apiSheet = null;

  var apiSheetId = PropertiesService.getScriptProperties().getProperty('API_SHEET_ID');
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === apiSheetId) {
      apiSheet = sheets[i];
      break;
    }
  }

  if (!apiSheet) {
    Logger.log("Did not find a sheet with an ID that matches the script property API_SHEET_ID.");
    return;
  }

  var cids = [];
  for (var j = 1; j <= 400; j++) {
    cids.push('P' + j);
  }

  var allData = [];
  var currentDate = new Date();
  var formattedDate = Utilities.formatDate(currentDate, spreadsheet.getSpreadsheetTimeZone(), "yyyy-MM-dd");

  for (var j = 0; j < cids.length; j++) {
    var url = "https://travel.state.gov/content/travel/resources/database/database.getVisaWaitTimes.html?cid=" + cids[j];
    var response = UrlFetchApp.fetch(url);
    var data = response.getContentText();
    if (!data || data.toLowerCase().includes("not found")) {
      Logger.log("No response for " + cids[j]);
      continue;
    }
    var waitTimes = data.split('|');
  
    var visaTypes = [
      "Interview Required Visitors (B1/B2)",
      "Interview Required Students/Exchange Visitors (F, M, J)",
      "",
      "Interview Required Crew and Transit (C, D, C1/D)",
      "Interview Waiver Visitors (B1/B2)",
      "",
      "",
      ""
    ];

    for (var i = 0; i < waitTimes.length; i++) {
      if (visaTypes[i] && waitTimes[i]) {
        var calendarDays = "";
        if (/days calendar days/i.test(waitTimes[i] + " Calendar Days")) {
          calendarDays = parseInt(waitTimes[i].trim().split(' ')[0]);
        }
        allData.push([formattedDate, cids[j], visaTypes[i], waitTimes[i] + " Calendar Days", calendarDays]);
      }
    }
  }

  if (allData.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, allData.length, allData[0].length).setValues(allData);
  }
}
*/