/* 
This works.

What it does: Gets the latest excel sheet that has visa issuances by post, then appends that data to this spreadsheet.

Note: It depends on the Drive API v2 (not v3!). Make sure it's enabled.

Todo:
* Test that it works in this spreadsheet. I copied it from another spreadsheet, so make a copy before running it again.
* Make sure it checks if the latest excel's data is already on the data sheet.
* Email and log more in general, but especially if it doesn't find expected data, as in it's past the 5th of the month, and there isn't new data (need to think about what woudl be a good trigger.)
* Set a trigger so that it runs every day.
* Get the Grand Total from the spreadsheet and put it on the summary sheet.
* Store the link to both the Excel sheet and PDF on the summary sheet.

*/

function fetchAndAppendLatestVisaDataByPost() {
  const stateDeptUrl = 'https://travel.state.gov/content/travel/en/legal/visa-law0/visa-statistics/nonimmigrant-visa-statistics/monthly-nonimmigrant-visa-issuances.html';
  const response = UrlFetchApp.fetch(stateDeptUrl);
  const html = response.getContentText();

  // Regex to find all Excel file links that are for "by post"
  const regex = /<li><a href="[^"]+"[^>]*>([A-Z][a-z]+ \d{4}) - NIV Issuances by Post and Visa Class<\/a>&nbsp;- \(<a href="([^"]+)"[^>]*>Excel<\/a>\)<\/li>/g;
  let match;
  let latestLink = '';
  let latestDateStr = '';
  let latestDate = new Date(0);

  while (match = regex.exec(html)) {
    const dateStr = match[1];
    const link = match[2];
    const date = new Date(dateStr + ' 1, ' + dateStr.split(' ')[1]);

    if (date > latestDate) {
      latestDate = date;
      latestLink = link;
      latestDateStr = dateStr;
    }
  }

  if (latestLink) {
    if (latestLink.startsWith('/')) {
      latestLink = 'https://travel.state.gov' + latestLink;
    }
    Logger.log('Latest Excel link: ' + latestLink);

    const excelResponse = UrlFetchApp.fetch(latestLink);
    const blob = excelResponse.getBlob();

    const tempFile = DriveApp.createFile(blob).setName('TemporaryFile');
    const tempFileId = tempFile.getId();

    const tempSpreadsheet = Drive.Files.copy(
      {
        title: 'NIV Issuances',
        mimeType: MimeType.GOOGLE_SHEETS
      },
      tempFileId
    );

    const spreadsheetId = tempSpreadsheet.id;
    const tempSheet = SpreadsheetApp.openById(spreadsheetId).getSheets()[0];
    const data = tempSheet.getDataRange().getValues();

    const startRow = data.findIndex(row => row[0] === 'Post' && row[1] === 'Visa Class' && row[2] === 'Issuances') + 1;
    const filteredData = data.slice(startRow + 1).filter(row => {
      const rowString = row.join(' ');
      return !row[0].startsWith('*') && !/grand total/i.test(rowString);
    }).map(row => [new Date(latestDate.getFullYear(), latestDate.getMonth(), 1), ...row]);

    const targetSheet = getSheetById(218073541); // Replace with your target sheet ID

    if (targetSheet) {
      appendDataToSheet(targetSheet, filteredData);
      DriveApp.getFileById(tempFileId).setTrashed(true); // Cleanup the original Excel file
      DriveApp.getFileById(spreadsheetId).setTrashed(true); // Cleanup the converted Google Sheet
    } else {
      Logger.log('Sheet with ID 218073541 not found.');
    }
  } else {
    Logger.log('No link found');
  }
}

function getSheetById(sheetId) {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() == sheetId) {
      return sheets[i];
    }
  }
  return null;
}

function appendDataToSheet(sheet, data) {
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, data.length, data[0].length).setValues(data);
}